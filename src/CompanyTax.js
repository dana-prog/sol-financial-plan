/** Straight-line depreciation by completion-month cohort, capped at each asset's cost. */
function calculateUnitDepreciation(assets, monthCount) {
  const depreciation = Array(monthCount).fill(0);
  assets.forEach(asset => {
    if (!Number.isFinite(asset.lifeYears) || asset.lifeYears < 0) {
      throw new Error('Invalid useful life for ' + asset.name + '. Nothing was written.');
    }
    if (asset.lifeYears === 0) return;
    const lifeMonths = asset.lifeYears * 12;
    asset.costs.forEach((cost, start) => {
      if (!Number.isFinite(cost) || cost < 0) throw new Error('Invalid construction cost for ' + asset.name);
      for (let age = 0; age < lifeMonths && start + age < monthCount; age++) {
        depreciation[start + age] += cost / lifeMonths * Math.min(1, lifeMonths - age);
      }
    });
  });
  return depreciation;
}

/** Annual tax with FIFO losses from the preceding five accounting periods. */
function calculateCompanyTax(months, profits, rate, yearEndMonth) {
  if (!Number.isFinite(rate) || rate < 0 || rate > 1) throw new Error('Company tax rate must be between 0% and 100%.');
  if (!Number.isInteger(yearEndMonth) || yearEndMonth < 1 || yearEndMonth > 12) {
    throw new Error('Financial year-end month must be a whole number from 1 to 12.');
  }
  if (profits.length !== months.length || !profits.every(Number.isFinite)) throw new Error('Invalid monthly tax profits.');
  const result = {cashPayments: months.map(() => 0), lossCarryforward: [], taxableProfit: [], taxExpense: [], years: []};
  const periods = new Map();
  months.forEach((month, index) => {
    const date = new Date((month - 25569) * 86400000);
    const year = date.getUTCFullYear() + (date.getUTCMonth() + 1 > yearEndMonth ? 1 : 0);
    if (!periods.has(year)) periods.set(year, []);
    periods.get(year).push(index);
  });
  let losses = [];
  const schedule = (dueDate, amount) => {
    const index = months.indexOf(repaymentMonthStart(dueDate));
    if (index >= 0) result.cashPayments[index] += amount;
  };
  periods.forEach((indexes, year) => {
    const expired = losses.filter(loss => loss.year < year - 5).reduce((sum, loss) => sum + loss.amount, 0);
    losses = losses.filter(loss => loss.year >= year - 5);
    const openingLoss = losses.reduce((sum, loss) => sum + loss.amount, 0);
    let profit = 0;
    let previousTaxable = 0;
    indexes.forEach(index => {
      profit += profits[index];
      const taxable = Math.max(0, profit - openingLoss);
      result.taxableProfit[index] = taxable - previousTaxable;
      result.taxExpense[index] = (taxable - previousTaxable) * rate;
      result.lossCarryforward[index] = Math.max(0, openingLoss - profit);
      previousTaxable = taxable;
    });
    const used = Math.min(openingLoss, Math.max(0, profit));
    let remaining = used;
    losses.forEach(loss => {
      const applied = Math.min(loss.amount, remaining);
      loss.amount -= applied;
      remaining -= applied;
    });
    if (profit < 0) losses.push({year, amount: -profit});
    const tax = previousTaxable * rate;
    // A forecast uses its full-year estimate, rather than treating six-month profit as annual profit.
    // Partial boundary periods have no half-year prepayment.
    const prepayment = indexes.length === 12 ? tax / 2 : 0;
    const yearEnd = repaymentDay(year, yearEndMonth, 0);
    const prepaymentDate = repaymentDay(year, yearEndMonth - 4, 0);
    const settlementDate = yearEnd + 150;
    schedule(prepaymentDate, prepayment);
    schedule(settlementDate, tax - prepayment);
    result.years.push({yearEnd, months: indexes.length, profit, openingLoss, expired, used,
      closingLoss: losses.reduce((sum, loss) => sum + loss.amount, 0), taxable: previousTaxable,
      tax, prepayment, prepaymentDate, settlement: tax - prepayment, settlementDate});
  });
  return result;
}

/** Resolve tax/interest feedback in memory; never read last run's tax payments as an input. */
function calculateRepaymentsWithCompanyTax(model) {
  let taxPayments = model.months.map(() => 0);
  for (let attempt = 0; attempt < 100; attempt++) {
    const repayments = calculateIncomingRepayments({...model,
      expenses: model.expenses.map((expense, i) => expense + taxPayments[i])});
    const profits = model.tax.profitBeforeInterest.map((profit, i) => profit - repayments.interestAccrued[i]);
    const tax = calculateCompanyTax(model.months, profits, model.tax.rate, model.tax.yearEndMonth);
    const difference = Math.max(...tax.cashPayments.map((payment, i) => Math.abs(payment - taxPayments[i])));
    if (difference < 0.000001) return {...repayments, tax};
    taxPayments = tax.cashPayments;
  }
  throw new Error('Tax and repayments did not reconcile. Nothing was written.');
}

function readUnitDepreciation(inputs, months, timezone) {
  const grid = inputs['Units Const.'].map(row => row.slice(0, months.length + 1));
  checkRepaymentTimeline(grid[0].slice(1), months, timezone, 'Units Const.');
  const rows = inputs.Units;
  const assets = rows.filter(row => row[0] !== '').map(row => {
    const name = String(row[0]);
    const matches = grid.filter(item => item[0] === name.toLowerCase() + ' const. cost');
    if (matches.length !== 1) throw new Error('Expected one construction cost schedule for ' + name);
    return {name, lifeYears: repaymentNumber(row[12], name + ' useful life'),
      costs: matches[0].slice(1).map(value => repaymentNumber(value, name + ' cost'))};
  });
  if (!assets.length) throw new Error('No unit depreciation inputs found.');
  return calculateUnitDepreciation(assets, months.length);
}

function companyTaxWriteRequests(spreadsheet, model, result) {
  const cash = repaymentSheet(spreadsheet, 'Cash Flow');
  const labels = model.cashLabels;
  const outputs = {
    'company tax': result.tax.cashPayments,
    'construction depreciation': model.tax.depreciation,
    'loss carryforward': result.tax.lossCarryforward,
    'taxable profit': result.tax.taxableProfit,
    'monthly company tax': result.tax.taxExpense
  };
  return Object.keys(outputs).map(label => {
    if (labels.filter(value => value === label).length !== 1) throw new Error('Expected one Cash Flow row: ' + label);
    return repaymentRowWrite(cash, labels.indexOf(label) + 1, 3, outputs[label], model.months.length);
  });
}
