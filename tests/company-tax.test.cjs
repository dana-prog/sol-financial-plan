const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const context = vm.createContext({Date});
for (const file of ['RepaymentCalculations.js', 'CompanyTax.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src', file), 'utf8'), context);
}
const day = (year, month, date = 1) => context.repaymentDay(year, month - 1, date);
const timeline = (count, year = 2026, month = 1) => Array.from({length: count}, (_, i) => day(year, month + i));
const near = (a, b) => assert.ok(Math.abs(a - b) < 0.00001, `${a} != ${b}`);
const sum = values => values.reduce((a, b) => a + b, 0);

test('depreciation uses each asset life, each completion month, and stops at cost', () => {
  const actual = context.calculateUnitDepreciation([
    {name: 'Short', lifeYears: 0.25, costs: [90, 30]},
    {name: 'Long', lifeYears: 1, costs: [0, 120]},
    {name: 'Nondepreciable', lifeYears: 0, costs: [500]}
  ], 18);
  assert.deepEqual(Array.from(actual.slice(0, 5)), [30, 50, 50, 20, 10]);
  near(sum(actual), 240);
  near(actual[13], 0);
  assert.throws(() => context.calculateUnitDepreciation([{name: 'Bad', lifeYears: -1, costs: [1]}], 12), /Invalid useful life/);
});

test('tax is annual, with August prepayment and settlement 150 days after year end', () => {
  const months = timeline(24);
  const tax = context.calculateCompanyTax(months, months.map(() => 100), 0.2, 12);
  near(tax.years[0].tax, 240);
  near(tax.cashPayments[7], 120);
  near(tax.cashPayments[16], 120);
  near(tax.cashPayments[19], 120);
  near(sum(tax.cashPayments), 360);
  near(sum(tax.taxExpense), 480);
  assert.equal(tax.years[0].settlementDate, day(2027, 5, 30));
  assert.equal(tax.years[1].settlementDate, day(2028, 5, 29));
});

test('a later loss reverses current-year tax accrual without creating a tax refund', () => {
  const profits = [100, -200, ...Array(10).fill(0)];
  const tax = context.calculateCompanyTax(timeline(12), profits, 0.2, 12);
  near(tax.taxExpense[0], 20);
  near(tax.taxExpense[1], -20);
  near(tax.years[0].tax, 0);
  near(sum(tax.cashPayments), 0);
  near(tax.years[0].closingLoss, 100);
});

test('losses remain available in the fifth subsequent period and expire in the sixth', () => {
  const profits = Array(84).fill(0);
  profits[0] = -100;
  profits[60] = 30;
  profits[72] = 100;
  const tax = context.calculateCompanyTax(timeline(84), profits, 0.2, 12);
  near(tax.years[5].used, 30);
  near(tax.years[6].expired, 70);
  near(tax.years[6].tax, 20);
});

test('oldest losses are consumed first and losses never carry back', () => {
  const profits = Array(84).fill(0);
  profits[0] = -100;
  profits[12] = -70;
  profits[60] = 110;
  profits[72] = 60;
  const tax = context.calculateCompanyTax(timeline(84), profits, 0.2, 12);
  near(tax.years[5].used, 110);
  near(tax.years[6].expired, 0);
  near(tax.years[6].used, 60);
  near(sum(tax.taxExpense), 0);
  const laterLoss = context.calculateCompanyTax(timeline(24), [...Array(12).fill(10), ...Array(12).fill(-10)], 0.2, 12);
  near(laterLoss.years[0].tax, 24);
  near(laterLoss.years[1].closingLoss, 120);
});

test('a June year end shifts prepayments and settlement and marks partial periods', () => {
  const months = timeline(24, 2026, 7);
  const tax = context.calculateCompanyTax(months, months.map(() => 100), 0.2, 6);
  assert.equal(tax.years[0].yearEnd, day(2027, 6, 30));
  near(tax.cashPayments[7], 120);
  near(tax.cashPayments[16], 120);
  const partial = context.calculateCompanyTax(timeline(6), Array(6).fill(100), 0.2, 6);
  assert.equal(partial.years[0].months, 6);
  near(partial.years[0].prepayment, 0);
  assert.throws(() => context.calculateCompanyTax(months, months.map(() => 0), 0.2, 13), /year-end month/);
});

test('tax and repayments converge, reserve is respected, and cash reconciles', () => {
  const months = timeline(96);
  const model = {months, loans: [{row: 4, name: 'Loan', priority: 1, date: months[0], rate: 0.12,
    draws: months.map((_, i) => i ? 0 : 10000), interestRepaymentStart: months[0]}],
    receipts: months.map(() => 500), expenses: months.map(() => 100), reserve: 1000,
    principalRepaymentStart: day(2027, 1),
    tax: {rate: 0.2, yearEndMonth: 12, profitBeforeInterest: months.map(() => 400)}};
  const original = JSON.stringify(model);
  const result = context.calculateRepaymentsWithCompanyTax(model);
  let cash = 0;
  months.forEach((month, i) => {
    const paid = sum(result.payments.map(p => p.principal[i] + p.interest[i]));
    cash += model.receipts[i] + model.loans[0].draws[i] - model.expenses[i] - paid - result.tax.cashPayments[i];
    near(result.cash[i], cash);
    assert.ok(paid === 0 || result.cash[i] >= model.reserve - 0.00001);
  });
  const check = context.calculateCompanyTax(months, result.interestAccrued.map(interest => 400 - interest), 0.2, 12);
  assert.deepEqual(result.tax, check);
  assert.deepEqual(result, context.calculateRepaymentsWithCompanyTax(model));
  assert.equal(JSON.stringify(model), original);
  assert.ok(sum(result.tax.cashPayments) > 0);
});

test('tax deducts accrued interest even when repayments have not started', () => {
  const months = timeline(12);
  const model = {months, loans: [{row: 4, name: 'Loan', priority: 1, date: months[0], rate: 0.1,
    draws: months.map((_, i) => i ? 0 : 1000), interestRepaymentStart: day(2030, 1)}],
    receipts: months.map(() => 100), expenses: months.map(() => 0), reserve: 0,
    principalRepaymentStart: day(2030, 1), tax: {rate: 0.2, yearEndMonth: 12, profitBeforeInterest: months.map(() => 100)}};
  const result = context.calculateRepaymentsWithCompanyTax(model);
  near(sum(result.payments[0].interest), 0);
  near(sum(result.interestAccrued), 100);
  near(result.tax.years[0].tax, 220);
});

test('tax output updates the five matching cash-flow rows and annual summary only', () => {
  const writes = [];
  const labels = Array(43).fill('');
  for (const [row, label] of [[11, 'company tax paid'], [36, 'construction depreciation'],
    [41, 'loss carryforward'], [42, 'taxable profit'], [43, 'company tax expense']]) labels[row - 1] = label;
  const cash = {getLastRow: () => 43, getRange(row, column, height, width) {
    if (column === 2) return {getValues: () => labels.map(label => [label])};
    return {setValues(values) {writes.push({sheet: 'Cash Flow', row, column, height, width, values});}};
  }};
  const summary = {getMaxRows: () => 30, getRange(row, column, height, width) {
    return {clearContent() {}, setValues(values) {writes.push({sheet: 'Company Tax', row, column, height, width, values});}};
  }};
  const spreadsheet = {getSheetByName(name) {
    assert.ok(['Cash Flow', 'Company Tax'].includes(name));
    return name === 'Cash Flow' ? cash : summary;
  }};
  const months = timeline(12);
  const tax = context.calculateCompanyTax(months, months.map(() => 100), 0.2, 12);
  context.writeCompanyTaxResults(spreadsheet, {months, tax: {depreciation: months.map(() => 10)}}, {tax});
  assert.deepEqual(writes.filter(w => w.sheet === 'Cash Flow').map(w => w.row), [11, 36, 41, 42, 43]);
  assert.equal(writes.at(-1).values[0][8], 240);
  writes.length = 0;
  labels[10] = 'Unexpected label';
  assert.throws(() => context.writeCompanyTaxResults(spreadsheet, {months, tax: {depreciation: []}}, {tax}), /Expected one Cash Flow row/);
  assert.equal(writes.length, 0);
});
