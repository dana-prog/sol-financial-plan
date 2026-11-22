/** Calculates in loan repayments in memory so cash and debt never depend on each other in Sheets. */
function calculateInLoanRepayments(model) {
  const states = model.loans.map(() => ({principal: 0, interest: 0, capitalized: 0, interestPaid: 0}));
  const payments = model.loans.map(loan => ({row: loan.row, principal: [], interest: []}));
  const order = model.loans.map((loan, index) => ({index, priority: loan.priority, row: loan.row}));
  const warnings = [];
  const cash = [];
  const interestAccrued = [];
  const interestByLoan = model.loans.map(() => []);
  let balance = 0;
  model.months.forEach((month, index) => {
    const accrued = model.loans.map((loan, i) => accrueInLoanMonth(loan, states[i], month, loan.draws[index]));
    accrued.forEach((state, i) => interestByLoan[i].push(state.interest - states[i].interest));
    interestAccrued.push(accrued.reduce((sum, state, i) => sum + state.interest - states[i].interest, 0));
    const draws = model.loans.reduce((sum, loan) => sum + loan.draws[index], 0);
    const beforeRepayments = balance + model.receipts[index] + draws - model.expenses[index];
    let available = Math.max(0, beforeRepayments - model.reserve);
    // Pay interest from its deadline month, prioritizing earlier deadlines.
    // Principal is ordered by maturity; priority and row break ties.
    for (const kind of ['interest', 'principal']) {
      const dueOrder = [...order].sort((a, b) => {
        const dueA = kind === 'interest' ? model.loans[a.index].interestRepaymentStart : model.loans[a.index].maturity;
        const dueB = kind === 'interest' ? model.loans[b.index].interestRepaymentStart : model.loans[b.index].maturity;
        return (dueA ?? Infinity) - (dueB ?? Infinity) || a.priority - b.priority || a.row - b.row;
      });
      dueOrder.forEach(({index: i}) => {
        const loan = model.loans[i];
        const start = kind === 'interest' ? loan.interestRepaymentStart : model.principalRepaymentStart;
        const requested = model.paymentPlan ? model.paymentPlan[i][kind][index] : available;
        if (model.paymentPlan && (!Number.isFinite(requested) || requested < -0.001)) throw new Error('Invalid repayment plan. Nothing was written.');
        const payment = start !== null && month >= start
          ? Math.min(Math.max(0, accrued[i][kind]), Math.max(0, requested)) : 0;
        payments[i][kind].push(payment);
        accrued[i][kind] -= payment;
        available -= payment;
      });
    }
    const paid = payments.reduce((sum, payment) => sum + payment.principal[index] + payment.interest[index], 0);
    model.loans.forEach((loan, i) => {
      if (loan.interestRepaymentStart === month && accrued[i].interest > 0.01 &&
          !payments[i].interest.some(value => value > 0.01)) {
        warnings.push({kind: 'interest', loanIndex: i, deadline: month});
      }
      if (loan.maturity === month && accrued[i].principal > 0.01) {
        warnings.push({kind: 'maturity', loanIndex: i, deadline: month,
          principal: accrued[i].principal, interest: accrued[i].interest});
      }
    });
    balance = beforeRepayments - paid;
    cash.push(balance);
    accrued.forEach((state, i) => {
      states[i] = {...state, interestPaid: payments[i].interest[index]};
    });
  });
  if (![...cash, ...payments.flatMap(p => [...p.principal, ...p.interest])].every(Number.isFinite)) {
    throw new Error('Repayment calculation produced an invalid amount. Nothing was written.');
  }
  const warningMessages = warnings.map(warning => {
    const loan = model.loans[warning.loanIndex];
    const label = loan.name + ': ';
    const amount = value => 'THB ' + value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    if (warning.kind === 'interest') {
      const firstPayment = payments[warning.loanIndex].interest.findIndex(value => value > 0.01);
      return label + 'interest payment has not started by its deadline\n'
        + '  Interest deadline month: ' + repaymentDisplayDate(warning.deadline, true) + '\n'
        + '  Projected first payment: ' + (firstPayment < 0 ? 'Not within the forecast' : repaymentDisplayDate(model.months[firstPayment], true));
    }
    return label + 'principal remains unpaid at maturity\n'
      + '  Maturity month: ' + repaymentDisplayDate(warning.deadline, true) + '\n'
      + '  Principal unpaid: ' + amount(warning.principal) + '\n'
      + '  Interest unpaid: ' + amount(warning.interest);
  });
  return {payments, cash, interestAccrued, interestByLoan, warnings: warningMessages, warningDetails: warnings};
}

function repaymentDisplayDate(day, monthOnly = false) {
  if (day == null) return 'Not available';
  const date = new Date((day - 25569) * 86400000);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return (monthOnly ? '' : String(date.getUTCDate()).padStart(2, '0') + ' ') + months[date.getUTCMonth()] + ' ' + date.getUTCFullYear();
}

function accrueInLoanMonth(loan, previous, month, draw) {
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
    throw new Error('Enter a non-negative whole number of years in the Interest repayment start deadline (years) field for ' + name + '. Nothing was written.');
  }
  const loanDate = new Date((date - 25569) * 86400000);
  const start = repaymentDay(loanDate.getUTCFullYear() + years, loanDate.getUTCMonth(), 1);
  if (!Number.isFinite(start)) throw new Error('Interest repayment start month is out of range for ' + name + '. Nothing was written.');
  return start;
}

/** Loan receipts are fixed inputs; report missed deadlines without changing them. */
function outgoingLoanWarnings(grid, paramIndex, months, timezone) {
  const header = grid[0];
  if (!grid.slice(1).some(row => row[paramIndex] === 'start balance')) return [];
  const col = name => repaymentHeaderIndex(header, name, 'Out Loans');
  const nameCol = col('Name'), dateCol = col('Date'), rateCol = col('Interest');
  const principalYearsCol = col('Full repayment within (years)');
  const interestYearsCol = col('Interest repayment start deadline (years)');
  const issues = [];
  const amount = value => 'THB ' + value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  grid.forEach((row, index) => {
    if (index === 0 || row[paramIndex] !== 'start balance') return;
    const name = String(row[nameCol] || 'Unnamed loan');
    if (!(row[dateCol] instanceof Date)) throw new Error('Expected a date for loan ' + name + '. Nothing was written.');
    const date = repaymentDate(row[dateCol], timezone, name + ' date');
    const rate = repaymentNumber(row[rateCol], name + ' interest');
    const years = row[principalYearsCol];
    if (!Number.isInteger(years) || years < 0) throw new Error('Expected a non-negative whole number of repayment years for loan ' + name + '. Nothing was written.');
    const principalRow = grid[index + 7], generatedRow = grid[index + 2], interestRow = grid[index + 9];
    if (!principalRow || principalRow[paramIndex] !== 'principal carried forward' || !generatedRow
      || generatedRow[paramIndex] !== 'interest generated' || !interestRow || interestRow[paramIndex] !== 'interest received') {
      throw new Error('Unexpected loan schedule layout for ' + name + '. Nothing was written.');
    }
    const d = new Date((date - 25569) * 86400000);
    const maturity = repaymentDay(d.getUTCFullYear() + years, d.getUTCMonth(), 1);
    const maturityIndex = months.indexOf(maturity);
    if (maturityIndex >= 0) {
      const outstanding = repaymentNumber(principalRow[paramIndex + 1 + maturityIndex], name + ' principal carried forward');
      if (outstanding > 0.01) issues.push('Loan ' + name + ': principal remains unpaid at maturity\n  Maturity month: '
        + repaymentDisplayDate(maturity, true) + '\n  Principal unpaid: ' + amount(outstanding));
    }
    const deadline = interestRepaymentStartMonth(date, rate, row[interestYearsCol], name);
    const deadlineIndex = months.indexOf(deadline);
    if (deadlineIndex >= 0) {
      const total = (r, end) => months.slice(0, end + 1).reduce((sum, _, i) =>
        sum + repaymentNumber(r[paramIndex + 1 + i], name + ' interest schedule'), 0);
      if (total(generatedRow, deadlineIndex) > 0.01 && total(interestRow, deadlineIndex) <= 0.01) {
        const first = months.findIndex((_, i) => repaymentNumber(interestRow[paramIndex + 1 + i], name + ' interest received') > 0.01);
        issues.push('Loan ' + name + ': interest receipt has not started by its deadline\n  Interest deadline month: '
          + repaymentDisplayDate(deadline, true) + '\n  Projected first receipt: '
          + (first < 0 ? 'Not within the forecast' : repaymentDisplayDate(months[first], true)));
      }
    }
  });
  return issues;
}

function showRepaymentCalculation() {
  const template = HtmlService.createTemplateFromFile('RepaymentProgress');
  template.action = 'Calculating repayments and tax…';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(460).setHeight(160), 'Recalculate In Loan Payments and Tax');
}

function repaymentResultsChanged(previous, next) {
  if (!Array.isArray(previous) || previous.length !== next.length) return true;
  return next.some((row, index) => !Array.isArray(previous[index]) || previous[index].length !== row.length
    || row.some((value, column) => typeof previous[index][column] !== 'number'
      || !Number.isFinite(previous[index][column]) || Math.abs(previous[index][column] - value) > 0.000001));
}

function recalculateInLoanRepayments() {
  const started = Date.now();
  const timings = {};
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another repayment calculation is running. Please try again shortly.');
  try {
    timings.lockMs = Date.now() - started;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let phase = Date.now();
    const model = readInLoanRepaymentModel(spreadsheet);
    timings.readMs = Date.now() - phase;
    phase = Date.now();
    const initial = calculateRepaymentsWithCompanyTax(model);
    const needsSearch = initial.warnings.length || initial.cash.some(value => value < -0.01);
    const search = needsSearch ? searchRepaymentSchedule(model, initial) : {result: initial, found: true};
    const result = search.result;
    timings.calculateMs = Date.now() - phase;
    phase = Date.now();
    const sheet = repaymentSheet(spreadsheet, 'In Loans');
    const requests = companyTaxWriteRequests(spreadsheet, model, result);
    result.payments.forEach(payment => {
      requests.push(repaymentRowWrite(sheet, payment.row + 6, model.timelineColumn, payment.principal, model.months.length));
      requests.push(repaymentRowWrite(sheet, payment.row + 9, model.timelineColumn, payment.interest, model.months.length));
    });
    const timestamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'dd MMM HH:mm');
    requests.push({updateCells: {
      range: {sheetId: sheet.getSheetId(), startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1},
      rows: [{values: [{userEnteredValue: {stringValue: 'Last Repayment Calculation Time\n' + timestamp},
        note: 'Last successful repayment and company-tax calculation. Run SOL → Recalculate In Loan Payments and Tax after changing loans, cash flow, unit costs, useful lives, or tax inputs.'}]}],
      fields: 'userEnteredValue,note'
    }});
    timings.prepareWriteMs = Date.now() - phase;
    phase = Date.now();
    // One Sheets batch prevents a failed repayment write from leaving new tax values beside old repayments.
    // Sheets checks permissions for every target as part of this atomic write.
    const response = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheet.getId() + ':batchUpdate', {
      method: 'post', contentType: 'application/json',
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
      payload: JSON.stringify({requests}), muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Could not save repayments and tax (HTTP ' + response.getResponseCode() + '). ' + response.getContentText());
    }
    timings.saveMs = Date.now() - phase;
    const issues = [...result.warnings, ...(model.outgoingWarnings || [])];
    result.cash.forEach((balance, index) => {
      if (balance < -0.01) issues.push('Negative cash closing: ' + repaymentDisplayDate(model.months[index], true)
        + '\n  Shortfall: THB ' + Math.abs(balance).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','));
    });
    const repaymentsChanged = repaymentResultsChanged(model.existingRepayments,
      result.payments.flatMap(payment => [payment.principal, payment.interest]));
    const taxChanged = repaymentResultsChanged(model.existingTaxResults, [result.tax.cashPayments,
      model.tax.depreciation, result.tax.lossCarryforward, result.tax.taxableProfit, result.tax.taxExpense]);
    const message = !repaymentsChanged && !taxChanged ? 'Repayment and tax results are unchanged.'
      : repaymentsChanged && taxChanged ? 'Repayment and tax results changed.'
      : repaymentsChanged ? 'Repayment results changed; tax results are unchanged.'
      : 'Tax results changed; repayment results are unchanged.';
    return (issues.length ? message + '\n\nCould not find a repayment schedule that resolves all issues:\n' + issues.join('\n\n')
      : message + '\n\nNo deadline or negative cash closing issues found within the forecast.');
  } finally {
    lock.releaseLock();
    console.log('Repayment calculation timings (ms): ' + JSON.stringify({...timings, totalMs: Date.now() - started}));
  }
}

function readInLoanRepaymentModel(spreadsheet) {
  const inputs = readRepaymentInputs(spreadsheet);
  // API date serials represent the spreadsheet's local calendar, not a UTC instant.
  const timezone = 'UTC';
  const grid = inputs['In Loans'];
  const headers = grid[0];
  const columns = Object.fromEntries(['Name', 'Priority', 'Date', 'Interest', 'Maturity date',
    'Interest repayment start deadline (years)'].map(name => [name, repaymentHeaderIndex(headers, name)]));
  const paramIndex = repaymentHeaderIndex(headers, 'Param');
  if (paramIndex < 9 || headers.lastIndexOf('Param') !== paramIndex) {
    throw new Error('Expected one In Loans Param header after the loan input columns.');
  }
  const timelineIndex = paramIndex + 1;
  const months = [];
  for (let c = timelineIndex; headers[c] instanceof Date; c++) months.push(repaymentDate(headers[c], timezone, 'In Loans timeline'));
  if (!months.length) throw new Error('The In Loans monthly timeline is missing.');
  months.forEach((day, i) => {
    const previous = new Date(((months[i - 1] || day) - 25569) * 86400000);
    if (day !== repaymentMonthStart(day) || (i && day !== repaymentDay(previous.getUTCFullYear(), previous.getUTCMonth() + 1, 1))) {
      throw new Error('The loan timeline must contain consecutive month-start dates.');
    }
  });
  const paramRows = inputs.Params;
  const parameter = name => {
    const matches = paramRows.filter(row => row[0] === name);
    if (matches.length !== 1) throw new Error('Expected exactly one parameter: ' + name);
    return matches[0][1];
  };
  const principalRepaymentStart = repaymentMonthStart(repaymentDate(parameter('in loan principal repayment start date'), timezone, 'in loan principal repayment start date'));
  const reserve = repaymentNumber(parameter('minimum cash reserve'), 'minimum cash reserve');
  if (reserve < 0) throw new Error('Minimum cash reserve cannot be negative.');
  const loans = [];
  const existingRepayments = [];
  const labels = ['start balance', 'in loans received', 'interest generated', 'in loans repaid', 'end balance', 'principal outstanding',
    'principal paid', 'principal carried forward', 'interest outstanding', 'interest paid', 'interest carried forward',
    'interest capitalization month', 'capitalized interest', 'balance bearing interest'];
  grid.forEach((row, index) => {
    if (row[paramIndex] !== 'start balance') return;
    if (!row[columns.Name]) throw new Error('Missing loan name at In Loans row ' + (index + 3));
    labels.forEach((label, offset) => {
      const actual = grid[index + offset]?.[paramIndex];
      const legacy = {'in loans received': 'loan draw', 'in loans repaid': 'loan paid', 'principal outstanding': 'principal owed', 'interest outstanding': 'interest owed'}[label];
      if (actual !== label && actual !== legacy) throw new Error('Unexpected schedule layout for ' + row[columns.Name] + ': ' + label);
    });
    const draws = grid[index + 1].slice(timelineIndex, timelineIndex + months.length).map(value => repaymentNumber(value, row[columns.Name] + ' in loan draw'));
    const date = row[columns.Date] === '' ? null : repaymentDate(row[columns.Date], timezone, row[columns.Name] + ' date');
    if (date === null && draws.some(Boolean)) throw new Error('A loan with draw amounts must have a date: ' + row[columns.Name]);
    const rate = repaymentNumber(row[columns.Interest], row[columns.Name] + ' interest');
    if (rate < 0 || draws.some(value => value < 0)) throw new Error('Negative in loan draws or interest rates are unsupported: ' + row[columns.Name]);
    const interestRepaymentStart = interestRepaymentStartMonth(date, rate, row[columns['Interest repayment start deadline (years)']], row[columns.Name]);
    const maturityValue = row[columns['Maturity date']];
    const maturity = maturityValue === '' || maturityValue == null ? null : repaymentMonthStart(repaymentDate(maturityValue, timezone, row[columns.Name] + ' maturity date'));
    loans.push({row: index + 3, name: row[columns.Name], priority: repaymentNumber(row[columns.Priority], row[columns.Name] + ' priority'),
      date, rate, draws, interestRepaymentStart, maturity});
    existingRepayments.push(grid[index + 6].slice(timelineIndex, timelineIndex + months.length),
      grid[index + 9].slice(timelineIndex, timelineIndex + months.length));
  });
  if (!loans.length) throw new Error('No in loan schedules were found.');
  const cash = inputs['Cash Flow'].map(row => row.slice(0, months.length + 1));
  checkRepaymentTimeline(cash[0].slice(1), months, timezone, 'Cash Flow');
  const cashRow = (name, legacyName) => {
    const rows = cash.slice(1).filter(row => row[0] === name || (legacyName && row[0] === legacyName));
    if (rows.length !== 1) throw new Error('Expected exactly one Cash Flow input row: ' + name);
    return rows[0].slice(1).map(value => repaymentNumber(value, 'Cash Flow: ' + name));
  };
  const expenses = Array(months.length).fill(0);
  ['land expenses', 'setup expenses', 'construction expenses', 'infra expenses',
    'direct operating expenses', 'other operating expenses', 'out loans disbursed']
    .forEach(name => cashRow(name, name === 'out loans disbursed' ? 'loans disbursed' : null).forEach((value, i) => { expenses[i] += value; }));
  const outgoingGrid = inputs['Out Loans'];
  const outgoingParam = repaymentHeaderIndex(outgoingGrid[0], 'Param', 'Out Loans');
  const outgoingRows = outgoingGrid.map(row => row.slice(outgoingParam, outgoingParam + months.length + 1));
  checkRepaymentTimeline(outgoingRows[0].slice(1), months, timezone, 'Out Loans');
  const receiptsRows = outgoingRows.filter(row => ['total out loan repayments received', 'total loan repayments received', 'total loan paid'].includes(row[0]));
  if (receiptsRows.length !== 1) throw new Error('Expected one total out loan repayments received row in Out Loans.');
  const revenue = cashRow('revenue');
  const receipts = receiptsRows[0].slice(1).map((value, i) => repaymentNumber(value, 'Out Loans repayments') + revenue[i]);
  const depreciation = readUnitDepreciation(inputs, months, timezone);
  const otherDepreciation = cashRow('setup depreciation');
  const operatingProfit = cashRow('operating profit');
  const interestIncome = months.map((_, i) => outgoingRows.filter(row => row[0] === 'interest generated')
    .reduce((sum, row) => sum + repaymentNumber(row[i + 1], 'Out Loans interest income'), 0));
  const tax = {rate: repaymentNumber(parameter('company tax rate'), 'company tax rate'),
    yearEndMonth: repaymentNumber(parameter('financial year-end month'), 'financial year-end month'), depreciation,
    profitBeforeInterest: operatingProfit.map((profit, i) => profit + interestIncome[i] - depreciation[i] - otherDepreciation[i])};
  const outgoingWarnings = outgoingLoanWarnings(outgoingGrid, outgoingParam, months, timezone);
  const existingTaxResults = ['company tax', 'construction depreciation', 'loss carryforward', 'taxable profit', 'monthly company tax']
    .map(label => cash.find(row => row[0] === label)?.slice(1));
  return {months, loans, principalRepaymentStart, reserve, receipts, expenses, tax, outgoingWarnings, existingRepayments, existingTaxResults, timelineColumn: timelineIndex + 1,
    cashLabels: ['', '', ...cash.map(row => row[0])]};
}

/** Read calculated inputs together; date formats preserve the existing date validation. */
function readRepaymentInputs(spreadsheet) {
  const ranges = [
    {name: 'In Loans', row: 3, column: 1},
    {name: 'Params', row: 2, column: 2, width: 2},
    {name: 'Cash Flow', row: 3, column: 2},
    {name: 'Out Loans', row: 3, column: 1},
    {name: 'Construction Exp (P)', row: 3, column: 1},
    {name: 'Units', row: 1, column: 4}
  ].map(input => {
    const sheet = repaymentSheet(spreadsheet, input.name);
    const height = sheet.getLastRow() - input.row + 1;
    const width = input.width || sheet.getLastColumn() - input.column + 1;
    if (height < 1 || width < 1) throw new Error('Missing repayment inputs in ' + input.name + '. Nothing was written.');
    const address = sheet.getRange(input.row, input.column, height, width).getA1Notation();
    return {...input, height, width, address: "'" + input.name.replace(/'/g, "''") + "'!" + address};
  });
  const fields = 'sheets(properties(title),data(rowData(values(effectiveValue,effectiveFormat(numberFormat(type))))))';
  const query = ranges.map(input => 'ranges=' + encodeURIComponent(input.address)).join('&');
  const response = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheet.getId()
    + '?' + query + '&fields=' + encodeURIComponent(fields), {
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Could not read repayment inputs (HTTP ' + response.getResponseCode() + '). Nothing was written. ' + response.getContentText());
  }
  const sheets = JSON.parse(response.getContentText()).sheets || [];
  return Object.fromEntries(ranges.map(input => {
    const sheet = sheets.find(item => item.properties.title === input.name);
    if (!sheet) throw new Error('Missing repayment inputs in ' + input.name + '. Nothing was written.');
    const rows = sheet.data?.[0]?.rowData || [];
    // The API omits trailing empty rows/cells; restore the rectangular getValues shape.
    return [input.name, Array.from({length: input.height}, (_, r) =>
      Array.from({length: input.width}, (_, c) => repaymentInputValue(rows[r]?.values?.[c])))];
  }));
}

function repaymentInputValue(cell) {
  const value = cell?.effectiveValue;
  if (!value) return '';
  if (value.errorValue) return '#' + value.errorValue.type;
  if (typeof value.numberValue === 'number') {
    const type = cell.effectiveFormat?.numberFormat?.type;
    return type === 'DATE' || type === 'DATE_TIME'
      ? new Date(Math.round((value.numberValue - 25569) * 86400000)) : value.numberValue;
  }
  return value.stringValue ?? value.boolValue ?? '';
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
  if (timezone === 'UTC') return repaymentDay(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate());
  const [year, month, day] = Utilities.formatDate(value, timezone, 'yyyy-MM-dd').split('-').map(Number);
  return repaymentDay(year, month - 1, day);
}

function checkRepaymentTimeline(dates, months, timezone, sheet) {
  if (dates.length !== months.length || dates.some((date, i) => repaymentDate(date, timezone, sheet + ' timeline') !== months[i])) {
    throw new Error(sheet + ' and In Loans must use the same monthly timeline. Nothing was written.');
  }
}


function repaymentHeaderIndex(headers, name, sheetName = 'In Loans') {
  const index = headers.indexOf(name);
  if (index < 0 || headers.lastIndexOf(name) !== index) {
    throw new Error('Expected exactly one ' + sheetName + ' header: ' + name + '. Nothing was written.');
  }
  return index;
}

function repaymentRowWrite(sheet, row, column, values, monthCount) {
  if (values.length !== monthCount || !values.every(Number.isFinite)) {
    throw new Error('Invalid result for ' + sheet.getName() + ' row ' + row + '. Nothing was written.');
  }
  return {updateCells: {
    range: {sheetId: sheet.getSheetId(), startRowIndex: row - 1, endRowIndex: row,
      startColumnIndex: column - 1, endColumnIndex: column - 1 + values.length},
    rows: [{values: values.map(numberValue => ({userEnteredValue: {numberValue}}))}],
    fields: 'userEnteredValue'
  }};
}
