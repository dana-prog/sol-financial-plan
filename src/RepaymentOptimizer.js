/** Find a whole-forecast schedule using fixed accrual and tax estimates.
 * The caller recalculates and verifies the proposed schedule against the actual model.
 * Amounts are expressed in millions of baht for numerical stability.
 */
function optimizeRepaymentPlan(model, accrual, taxPayments) {
  const scale = 1000000, big = 1000000;
  const engine = LinearOptimizationService.createEngine();
  const names = [], lowers = [], uppers = [], types = [];
  const constraints = [];
  const cashShortfalls = [], missedPrincipal = [], missedInterest = [], reserveShortfalls = [];
  const variable = (name, lower, upper, type = LinearOptimizationService.VariableType.CONTINUOUS) => {
    names.push(name); lowers.push(lower); uppers.push(upper);
    types.push(type);
  };
  const constraint = (lower, upper) => {
    const row = {lower, upper, names: [], coefficients: [], setCoefficient(name, value) {
      this.names.push(name); this.coefficients.push(value); return this;
    }};
    constraints.push(row);
    return row;
  };
  const equality = value => constraint(value, value);
  const n = model.months.length;
  for (let t = 0; t < n; t++) {
    variable('c' + t, 0, big);
    variable('n' + t, 0, big);
    variable('s' + t, 0, big);
    cashShortfalls.push('n' + t); reserveShortfalls.push('s' + t);
    const reserve = constraint(model.reserve / scale, big);
    reserve.setCoefficient('c' + t, 1).setCoefficient('s' + t, 1);
    const net = (model.receipts[t] - model.expenses[t] - taxPayments[t]
      + model.loans.reduce((sum, loan) => sum + loan.draws[t], 0)) / scale;
    const cash = equality(net);
    cash.setCoefficient('c' + t, 1).setCoefficient('n' + t, -1);
    if (t) cash.setCoefficient('c' + (t - 1), -1).setCoefficient('n' + (t - 1), 1);
    model.loans.forEach((loan, i) => {
      const month = model.months[t];
      const startInterest = loan.interestRepaymentStart !== null && month >= loan.interestRepaymentStart;
      const startPrincipal = model.principalRepaymentStart !== null && month >= model.principalRepaymentStart;
      const tag = i + '_' + t;
      variable('p' + tag, 0, startPrincipal ? big : 0);
      variable('i' + tag, 0, startInterest ? big : 0);
      variable('b' + tag, 0, big);
      if (loan.maturity === month) missedPrincipal.push('b' + tag);
      variable('a' + tag, 0, big);
      cash.setCoefficient('p' + tag, 1).setCoefficient('i' + tag, 1);
      const principal = equality(loan.draws[t] / scale);
      principal.setCoefficient('b' + tag, 1).setCoefficient('p' + tag, 1);
      const interest = equality(accrual[i][t] / scale);
      interest.setCoefficient('a' + tag, 1).setCoefficient('i' + tag, 1);
      if (t) {
        principal.setCoefficient('b' + i + '_' + (t - 1), -1);
        interest.setCoefficient('a' + i + '_' + (t - 1), -1);
      }
    });
  }
  model.loans.forEach((loan, i) => {
    const deadline = model.months.indexOf(loan.interestRepaymentStart);
    if (deadline < 0 || accrual[i].slice(0, deadline + 1).reduce((a, b) => a + b, 0) <= 0.01) return;
    const started = constraint(1 / scale, big);
    variable('m' + i, 0, 1, LinearOptimizationService.VariableType.INTEGER);
    missedInterest.push('m' + i);
    started.setCoefficient('m' + i, 1 / scale);
    for (let t = 0; t <= deadline; t++) started.setCoefficient('i' + i + '_' + t, 1);
  });
  engine.addVariables(names, lowers, uppers, types, names.map(() => 0));
  engine.addConstraints(constraints.map(row => row.lower), constraints.map(row => row.upper),
    constraints.map(row => row.names), constraints.map(row => row.coefficients));
  engine.setMinimization();
  let solution, previous = [];
  // Solve each goal in order and constrain the best result before moving to the next.
  for (const group of [cashShortfalls, missedPrincipal, missedInterest, reserveShortfalls]) {
    if (!group.length) continue;
    previous.forEach(name => engine.setObjectiveCoefficient(name, 0));
    group.forEach(name => engine.setObjectiveCoefficient(name, 1));
    solution = engine.solve(8);
    if (!solution.isValid()) return null;
    const value = group.reduce((sum, name) => sum + solution.getVariableValue(name), 0);
    const lock = engine.addConstraint(0, value + 0.000000001);
    group.forEach(name => lock.setCoefficient(name, 1));
    previous = group;
  }
  if (!solution) return null;
  return model.loans.map((loan, i) => ({row: loan.row,
    principal: model.months.map((_, t) => Math.max(0, solution.getVariableValue('p' + i + '_' + t) * scale)),
    interest: model.months.map((_, t) => Math.max(0, solution.getVariableValue('i' + i + '_' + t) * scale))}));
}

function repaymentPriorityScore(result) {
  return [result.cash.reduce((sum, value) => sum + Math.max(0, -value), 0),
    result.warningDetails.filter(issue => issue.kind === 'maturity').reduce((sum, issue) => sum + issue.principal, 0),
    result.warningDetails.filter(issue => issue.kind === 'interest').length];
}

function repaymentScoreImproves(candidate, prior) {
  const a = repaymentPriorityScore(candidate), b = repaymentPriorityScore(prior);
  for (let i = 0; i < a.length; i++) {
    if (Math.abs(a[i] - b[i]) > (i === 2 ? 0 : 0.01)) return a[i] < b[i];
  }
  return false;
}

/** Search using updated exact accrual and tax estimates; never assert infeasibility. */
function searchRepaymentSchedule(model, initial) {
  let current = initial, best = initial;
  const started = Date.now();
  let attempts = 0, rejected = 0, reason = 'The search reached its iteration limit.';
  for (let attempt = 0; attempt < 12 && Date.now() - started < 90000; attempt++) {
    attempts++;
    const plan = optimizeRepaymentPlan(model, current.interestByLoan, current.tax.cashPayments);
    if (!plan) {
      reason = 'The optimizer did not return a usable schedule.';
      break;
    }
    const next = calculateRepaymentsWithCompanyTax({...model, paymentPlan: plan});
    const matches = plan.every((payment, i) => ['principal', 'interest'].every(kind =>
      payment[kind].every((value, t) => Math.abs(value - next.payments[i][kind][t]) <= 0.01)));
    if (!matches) rejected++;
    if (matches && repaymentScoreImproves(next, best)) best = next;
    if (matches && !next.warnings.length && next.cash.every(value => value >= -0.01)) {
      return {result: next, found: true, attempts};
    }
    current = next;
  }
  if (Date.now() - started >= 90000) reason = 'The search reached its time limit.';
  if (rejected && attempts === rejected) reason += ' Proposed payments could not be reconciled with recalculated interest.';
  console.log('Repayment search: ' + JSON.stringify({attempts, rejected, reason, elapsedMs: Date.now() - started}));
  return {result: best, found: false, attempts, reason};
}
