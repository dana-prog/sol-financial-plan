const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function fixture(solutionValues) {
  const engine = {
    variables: {}, constraints: [],
    addVariables(names, mins, maxes) { names.forEach((name, i) => { this.variables[name] = {min: mins[i], max: maxes[i]}; }); },
    addConstraints(mins, maxes, names, coefficients) {
      mins.forEach((min, i) => this.constraints.push({min, max: maxes[i], terms: Object.fromEntries(names[i].map((name, j) => [name, coefficients[i][j]]))}));
    },
    setObjectiveCoefficient() {}, setMinimization() {},
    addConstraint(min, max) {
      const c = {min, max, terms: {}, setCoefficient(name, value) { this.terms[name] = value; return this; }};
      this.constraints.push(c); return c;
    },
    solve() {
      if (!solutionValues) return {isValid: () => false};
      for (const [name, bounds] of Object.entries(this.variables)) {
        const value = solutionValues[name] || 0;
        assert.ok(value >= bounds.min - 1e-9 && value <= bounds.max + 1e-9, name);
      }
      for (const c of this.constraints) {
        const actual = Object.entries(c.terms).reduce((sum, [name, coefficient]) => sum + coefficient * (solutionValues[name] || 0), 0);
        assert.ok(actual >= c.min - 1e-9 && actual <= c.max + 1e-9, `constraint ${actual} not within ${c.min}, ${c.max}`);
      }
      return {isValid: () => true, getVariableValue: name => solutionValues[name] || 0};
    }
  };
  const ctx = vm.createContext({Date, console, LinearOptimizationService: {createEngine: () => engine, VariableType: {CONTINUOUS: 'CONTINUOUS', INTEGER: 'INTEGER'}}});
  for (const file of ['RepaymentCalculations.js', 'CompanyTax.js', 'RepaymentOptimizer.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', file), 'utf8'), ctx);
  }
  const months = [0, 1, 2].map(i => ctx.repaymentDay(2026, i, 1));
  const model = {months, loans: [{row: 4, name: 'Loan', priority: 1, date: months[0], rate: 0,
    draws: [100000000, 0, 0], interestRepaymentStart: null, maturity: months[2]}],
    receipts: [0, 0, 100000000], expenses: [0, 80000000, 0], reserve: 0,
    principalRepaymentStart: months[0], tax: {profitBeforeInterest: [0, 0, 0], rate: 0, yearEndMonth: 12}};
  return {ctx, model, engine};
}

test('a verified optimized schedule holds cash for a future expense and meets maturity', () => {
  const values = {'c0': 80, 'c1': 0, 'c2': 20, 'p0_0': 20, 'p0_2': 80, 'b0_0': 80, 'b0_1': 80};
  const {ctx, model} = fixture(values);
  const initial = ctx.calculateRepaymentsWithCompanyTax(model);
  assert.equal(initial.cash[1], -80000000);
  const search = ctx.searchRepaymentSchedule(model, initial);
  assert.equal(search.found, true);
  assert.deepEqual(Array.from(search.result.cash), [80000000, 0, 20000000]);
  assert.deepEqual(Array.from(search.result.payments[0].principal), [20000000, 0, 80000000]);
  assert.equal(search.result.warnings.length, 0);
});

test('an unsolved optimization does not claim success', () => {
  const {ctx, model} = fixture(null);
  const initial = ctx.calculateRepaymentsWithCompanyTax(model);
  const search = ctx.searchRepaymentSchedule(model, initial);
  assert.equal(search.found, false);
  assert.match(search.reason, /did not return a usable schedule/);
  assert.equal(search.result.cash[1], -80000000);
});

test('optimizer prevents interest payments before the loan deadline', () => {
  const {ctx, model, engine} = fixture(null);
  model.loans[0].interestRepaymentStart = model.months[2];
  ctx.optimizeRepaymentPlan(model, [[100000, 100000, 100000]], [0, 0, 0]);
  assert.equal(engine.variables.i0_0.max, 0);
  assert.equal(engine.variables.i0_1.max, 0);
  assert.ok(engine.variables.i0_2.max > 0);
});

test('cash shortfalls take precedence over principal and interest deadlines', () => {
  const {ctx} = fixture(null);
  const result = (cash, principal, missedInterest) => ({cash: [cash], warningDetails: [
    ...(principal ? [{kind:'maturity', principal}] : []),
    ...Array.from({length: missedInterest}, () => ({kind:'interest'}))]});
  assert.equal(ctx.repaymentScoreImproves(result(0, 100, 1), result(-1, 0, 0)), true);
  assert.equal(ctx.repaymentScoreImproves(result(0, 0, 1), result(0, 1, 0)), true);
  assert.equal(ctx.repaymentScoreImproves(result(-1, 0, 0), result(0, 100, 1)), false);
});
