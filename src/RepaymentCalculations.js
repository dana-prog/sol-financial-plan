/** Calculates incoming repayments in memory so cash and debt never depend on each other in Sheets. */
function calculateIncomingRepayments(model) {
  const states = model.loans.map(() => ({principal: 0, interest: 0, capitalized: 0, interestPaid: 0}));
  const payments = model.loans.map(loan => ({row: loan.row, principal: [], interest: []}));
  const order = model.loans.map((loan, index) => ({index, priority: loan.priority, row: loan.row}))
    .sort((a, b) => a.priority - b.priority || a.row - b.row);
  const cash = [];
  const interestAccrued = [];
  let balance = 0;
  model.months.forEach((month, index) => {
    const accrued = model.loans.map((loan, i) => accrueIncomingLoanMonth(loan, states[i], month, loan.draws[index]));
    interestAccrued.push(accrued.reduce((sum, state, i) => sum + state.interest - states[i].interest, 0));
    const draws = model.loans.reduce((sum, loan) => sum + loan.draws[index], 0);
    const beforeRepayments = balance + model.receipts[index] + draws - model.expenses[index];
    let available = Math.max(0, beforeRepayments - model.reserve);
    // Pay eligible interest before principal, then use loan priority and row order.
    for (const kind of ['interest', 'principal']) {
      order.forEach(({index: i}) => {
        const start = kind === 'interest' ? model.loans[i].interestRepaymentStart : model.principalRepaymentStart;
        const payment = start !== null && month >= start
          ? Math.min(Math.max(0, accrued[i][kind]), available) : 0;
        payments[i][kind].push(payment);
        accrued[i][kind] -= payment;
        available -= payment;
      });
    }
    const paid = payments.reduce((sum, payment) => sum + payment.principal[index] + payment.interest[index], 0);
    balance = beforeRepayments - paid;
    cash.push(balance);
    accrued.forEach((state, i) => {
      states[i] = {...state, interestPaid: payments[i].interest[index]};
    });
  });
  if (![...cash, ...payments.flatMap(p => [...p.principal, ...p.interest])].every(Number.isFinite)) {
    throw new Error('Repayment calculation produced an invalid amount. Nothing was written.');
  }
  return {payments, cash, interestAccrued};
}

function accrueIncomingLoanMonth(loan, previous, month, draw) {
  const principal = previous.principal + draw;
  if (loan.date === null || month < repaymentMonthStart(loan.date)) {
    return {principal, interest: previous.interest, capitalized: 0};
  }
  const date = new Date((loan.date - 25569) * 86400000);
  const current = new Date((month - 25569) * 86400000);
  const end = repaymentDay(current.getUTCFullYear(), current.getUTCMonth() + 1, 0);
  const anniversary = month > repaymentMonthStart(loan.date) && current.getUTCMonth() === date.getUTCMonth();
  const anniversaryDay = repaymentDay(current.getUTCFullYear(), date.getUTCMonth(),
    Math.min(date.getUTCDate(), end - month + 1));
  const remainingCapitalized = Math.max(0, previous.capitalized - previous.interestPaid);
  const dailyRate = loan.rate / 365;
  const capitalized = anniversary
    ? previous.interest + (principal + remainingCapitalized) * dailyRate * Math.max(0, anniversaryDay - month)
    : remainingCapitalized;
  // Match the existing Actual/365 formulas, including interest capitalized on the loan anniversary.
  const generated = ((principal + remainingCapitalized) * (end - Math.max(month, loan.date) + 1)
    + (anniversary ? (capitalized - remainingCapitalized) * (end - anniversaryDay + 1) : 0)) * dailyRate;
  return {principal, interest: previous.interest + generated, capitalized};
}

function repaymentDay(year, month, day) {
  return Math.round(Date.UTC(year, month, day) / 86400000) + 25569;
}

function repaymentMonthStart(day) {
  const date = new Date((day - 25569) * 86400000);
  return repaymentDay(date.getUTCFullYear(), date.getUTCMonth(), 1);
}

function interestRepaymentStartMonth(date, rate, years, name) {
  if (rate === 0 || date === null) return null;
  if (typeof years !== 'number' || !Number.isInteger(years) || years < 0) {
    throw new Error('Enter a non-negative whole number of years in column G for ' + name + '. Nothing was written.');
  }
  const loanDate = new Date((date - 25569) * 86400000);
  const start = repaymentDay(loanDate.getUTCFullYear() + years, loanDate.getUTCMonth(), 1);
  if (!Number.isFinite(start)) throw new Error('Interest repayment start month is out of range for ' + name + '. Nothing was written.');
  return start;
}

function showRepaymentCalculation() {
  const template = HtmlService.createTemplateFromFile('LoanModeProgress');
  template.action = 'Calculating repayments and tax…';
  template.completionMessage = '';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(420).setHeight(70), 'Recalculate Loan Payments and Tax');
}

function recalculateIncomingLoanRepayments() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another repayment calculation is running. Please try again shortly.');
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const model = readIncomingRepaymentModel(spreadsheet);
    const result = calculateRepaymentsWithCompanyTax(model);
    writeCompanyTaxResults(spreadsheet, model, result);
    const sheet = spreadsheet.getSheetByName('In Loans');
    result.payments.forEach(payment => {
      sheet.getRange(payment.row + 6, 11, 1, model.months.length).setValues([payment.principal]);
      sheet.getRange(payment.row + 9, 11, 1, model.months.length).setValues([payment.interest]);
    });
    SpreadsheetApp.flush();
    const timestamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM HH:mm');
    sheet.getRange('D1').clearDataValidations().setValue('Last Repayment Calculation Time\n' + timestamp)
      .setNote('Last successful repayment and company-tax calculation. Run SOL → Recalculate Loan Payments and Tax after changing loans, cash flow, unit costs, useful lives, or tax inputs.');
    return 'Loan repayments and tax updated.';
  } finally {
    lock.releaseLock();
  }
}

function readIncomingRepaymentModel(spreadsheet) {
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const incoming = repaymentSheet(spreadsheet, 'In Loans');
  const grid = incoming.getRange(3, 1, incoming.getLastRow() - 2, incoming.getLastColumn()).getValues();
  const headers = grid[0];
  if (headers[9] !== 'Param') throw new Error('Expected the In Loans Param header in column J.');
  const months = [];
  for (let c = 10; headers[c] instanceof Date; c++) months.push(repaymentDate(headers[c], timezone, 'In Loans timeline'));
  if (!months.length) throw new Error('The In Loans monthly timeline is missing.');
  months.forEach((day, i) => {
    const previous = new Date(((months[i - 1] || day) - 25569) * 86400000);
    if (day !== repaymentMonthStart(day) || (i && day !== repaymentDay(previous.getUTCFullYear(), previous.getUTCMonth() + 1, 1))) {
      throw new Error('The loan timeline must contain consecutive month-start dates.');
    }
  });
  const params = repaymentSheet(spreadsheet, 'Params');
  const paramRows = params.getRange(2, 2, params.getLastRow() - 1, 2).getValues();
  const parameter = name => {
    const matches = paramRows.filter(row => row[0] === name);
    if (matches.length !== 1) throw new Error('Expected exactly one parameter: ' + name);
    return matches[0][1];
  };
  const principalRepaymentStart = repaymentMonthStart(repaymentDate(parameter('principal loan repayment start month'), timezone, 'principal loan repayment start month'));
  const reserve = repaymentNumber(parameter('minimum cash reserve'), 'minimum cash reserve');
  if (reserve < 0) throw new Error('Minimum cash reserve cannot be negative.');
  const loans = [];
  const labels = ['start balance', 'loan draw', 'interest generated', 'loan paid', 'end balance', 'principal owed',
    'principal paid', 'principal carried forward', 'interest owed', 'interest paid', 'interest carried forward',
    'interest capitalization month', 'capitalized interest', 'balance bearing interest'];
  grid.forEach((row, index) => {
    if (row[9] !== 'start balance') return;
    if (!row[0]) throw new Error('Missing loan name at In Loans row ' + (index + 3));
    labels.forEach((label, offset) => {
      if (grid[index + offset]?.[9] !== label) throw new Error('Unexpected schedule layout for ' + row[0] + ': ' + label);
    });
    const draws = grid[index + 1].slice(10, 10 + months.length).map(value => repaymentNumber(value, row[0] + ' loan draw'));
    const date = row[2] === '' ? null : repaymentDate(row[2], timezone, row[0] + ' date');
    if (date === null && draws.some(Boolean)) throw new Error('A loan with draw amounts must have a date: ' + row[0]);
    const rate = repaymentNumber(row[5], row[0] + ' interest');
    if (rate < 0 || draws.some(value => value < 0)) throw new Error('Negative loan draws or interest rates are unsupported: ' + row[0]);
    const interestRepaymentStart = interestRepaymentStartMonth(date, rate, row[6], row[0]);
    loans.push({row: index + 3, name: row[0], priority: repaymentNumber(row[1], row[0] + ' priority'),
      date, rate, draws, interestRepaymentStart});
  });
  if (!loans.length) throw new Error('No incoming loan schedules were found.');
  const cash = repaymentSheet(spreadsheet, 'Cash Flow').getRange(3, 2, 40, months.length + 1).getValues();
  checkRepaymentTimeline(cash[0].slice(1), months, timezone, 'Cash Flow');
  const cashRow = name => {
    const rows = cash.slice(1).filter(row => row[0] === name);
    if (rows.length !== 1) throw new Error('Expected exactly one Cash Flow input row: ' + name);
    return rows[0].slice(1).map(value => repaymentNumber(value, 'Cash Flow: ' + name));
  };
  const expenses = Array(months.length).fill(0);
  ['land expenses', 'pre const. expenses', 'units const. expenses', 'other const. expenses',
    'direct operating expenses', 'other operating expenses', 'outgoing loans - funds disbursed']
    .forEach(name => cashRow(name).forEach((value, i) => { expenses[i] += value; }));
  const outgoing = repaymentSheet(spreadsheet, 'Out Loans');
  const outgoingRows = outgoing.getRange(3, 8, outgoing.getLastRow() - 2, months.length + 1).getValues();
  checkRepaymentTimeline(outgoingRows[0].slice(1), months, timezone, 'Out Loans');
  const receiptsRows = outgoingRows.filter(row => row[0] === 'total loan paid');
  if (receiptsRows.length !== 1) throw new Error('Expected one total loan paid row in Out Loans.');
  const revenue = cashRow('revenue');
  const receipts = receiptsRows[0].slice(1).map((value, i) => repaymentNumber(value, 'Out Loans repayments') + revenue[i]);
  const depreciation = readUnitDepreciation(spreadsheet, months, timezone);
  const otherDepreciation = cashRow('setup depreciation');
  const operatingProfit = cashRow('operating profit');
  const interestIncome = months.map((_, i) => outgoingRows.filter(row => row[0] === 'interest generated')
    .reduce((sum, row) => sum + repaymentNumber(row[i + 1], 'Out Loans interest income'), 0));
  const tax = {rate: repaymentNumber(parameter('company tax rate'), 'company tax rate'),
    yearEndMonth: repaymentNumber(parameter('financial year-end month'), 'financial year-end month'), depreciation,
    profitBeforeInterest: operatingProfit.map((profit, i) => profit + interestIncome[i] - depreciation[i] - otherDepreciation[i])};
  return {months, loans, principalRepaymentStart, reserve, receipts, expenses, tax};
}

function repaymentSheet(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

function repaymentNumber(value, label) {
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error('Expected a numeric value for ' + label + '. Nothing was written.');
  return value;
}

function repaymentDate(value, timezone, label) {
  if (!(value instanceof Date) || !Number.isFinite(value.getTime())) throw new Error('Expected a date for ' + label + '. Nothing was written.');
  const [year, month, day] = Utilities.formatDate(value, timezone, 'yyyy-MM-dd').split('-').map(Number);
  return repaymentDay(year, month - 1, day);
}

function checkRepaymentTimeline(dates, months, timezone, sheet) {
  if (dates.length !== months.length || dates.some((date, i) => repaymentDate(date, timezone, sheet + ' timeline') !== months[i])) {
    throw new Error(sheet + ' and In Loans must use the same monthly timeline. Nothing was written.');
  }
}
