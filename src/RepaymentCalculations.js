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
    throw new Error('Enter a non-negative whole number of years in the Interest repayment start deadline (years) field for ' + name + '. Nothing was written.');
  }
  const loanDate = new Date((date - 25569) * 86400000);
  const start = repaymentDay(loanDate.getUTCFullYear() + years, loanDate.getUTCMonth(), 1);
  if (!Number.isFinite(start)) throw new Error('Interest repayment start month is out of range for ' + name + '. Nothing was written.');
  return start;
}

function showRepaymentCalculation() {
  const template = HtmlService.createTemplateFromFile('RepaymentProgress');
  template.action = 'Calculating repayments and tax…';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(460).setHeight(160), 'Recalculate Loan Payments and Tax');
}

function recalculateIncomingLoanRepayments() {
  const started = Date.now();
  const timings = {};
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another repayment calculation is running. Please try again shortly.');
  try {
    timings.lockMs = Date.now() - started;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let phase = Date.now();
    const model = readIncomingRepaymentModel(spreadsheet);
    timings.readMs = Date.now() - phase;
    phase = Date.now();
    const result = calculateRepaymentsWithCompanyTax(model);
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
        note: 'Last successful repayment and company-tax calculation. Run SOL → Recalculate Loan Payments and Tax after changing loans, cash flow, unit costs, useful lives, or tax inputs.'}]}],
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
    return 'Loan repayments and tax updated in ' + ((Date.now() - started) / 1000).toFixed(1) + ' seconds.';
  } finally {
    lock.releaseLock();
    console.log('Repayment calculation timings (ms): ' + JSON.stringify({...timings, totalMs: Date.now() - started}));
  }
}

function readIncomingRepaymentModel(spreadsheet) {
  const inputs = readRepaymentInputs(spreadsheet);
  // API date serials represent the spreadsheet's local calendar, not a UTC instant.
  const timezone = 'UTC';
  const grid = inputs['In Loans'];
  const headers = grid[0];
  const columns = Object.fromEntries(['Name', 'Priority', 'Date', 'Interest',
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
  const principalRepaymentStart = repaymentMonthStart(repaymentDate(parameter('principal loan repayment start month'), timezone, 'principal loan repayment start month'));
  const reserve = repaymentNumber(parameter('minimum cash reserve'), 'minimum cash reserve');
  if (reserve < 0) throw new Error('Minimum cash reserve cannot be negative.');
  const loans = [];
  const labels = ['start balance', 'loan draw', 'interest generated', 'loan paid', 'end balance', 'principal owed',
    'principal paid', 'principal carried forward', 'interest owed', 'interest paid', 'interest carried forward',
    'interest capitalization month', 'capitalized interest', 'balance bearing interest'];
  grid.forEach((row, index) => {
    if (row[paramIndex] !== 'start balance') return;
    if (!row[columns.Name]) throw new Error('Missing loan name at In Loans row ' + (index + 3));
    labels.forEach((label, offset) => {
      if (grid[index + offset]?.[paramIndex] !== label) throw new Error('Unexpected schedule layout for ' + row[columns.Name] + ': ' + label);
    });
    const draws = grid[index + 1].slice(timelineIndex, timelineIndex + months.length).map(value => repaymentNumber(value, row[columns.Name] + ' loan draw'));
    const date = row[columns.Date] === '' ? null : repaymentDate(row[columns.Date], timezone, row[columns.Name] + ' date');
    if (date === null && draws.some(Boolean)) throw new Error('A loan with draw amounts must have a date: ' + row[columns.Name]);
    const rate = repaymentNumber(row[columns.Interest], row[columns.Name] + ' interest');
    if (rate < 0 || draws.some(value => value < 0)) throw new Error('Negative loan draws or interest rates are unsupported: ' + row[columns.Name]);
    const interestRepaymentStart = interestRepaymentStartMonth(date, rate, row[columns['Interest repayment start deadline (years)']], row[columns.Name]);
    loans.push({row: index + 3, name: row[columns.Name], priority: repaymentNumber(row[columns.Priority], row[columns.Name] + ' priority'),
      date, rate, draws, interestRepaymentStart});
  });
  if (!loans.length) throw new Error('No incoming loan schedules were found.');
  const cash = inputs['Cash Flow'].map(row => row.slice(0, months.length + 1));
  checkRepaymentTimeline(cash[0].slice(1), months, timezone, 'Cash Flow');
  const cashRow = name => {
    const rows = cash.slice(1).filter(row => row[0] === name);
    if (rows.length !== 1) throw new Error('Expected exactly one Cash Flow input row: ' + name);
    return rows[0].slice(1).map(value => repaymentNumber(value, 'Cash Flow: ' + name));
  };
  const expenses = Array(months.length).fill(0);
  ['land expenses', 'planning expenses', 'construction expenses', 'infra expenses',
    'direct operating expenses', 'other operating expenses', 'outgoing loans - funds disbursed']
    .forEach(name => cashRow(name).forEach((value, i) => { expenses[i] += value; }));
  const outgoingRows = inputs['Out Loans'].map(row => row.slice(0, months.length + 1));
  checkRepaymentTimeline(outgoingRows[0].slice(1), months, timezone, 'Out Loans');
  const receiptsRows = outgoingRows.filter(row => row[0] === 'total loan paid');
  if (receiptsRows.length !== 1) throw new Error('Expected one total loan paid row in Out Loans.');
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
  return {months, loans, principalRepaymentStart, reserve, receipts, expenses, tax, timelineColumn: timelineIndex + 1,
    cashLabels: ['', '', ...cash.map(row => row[0])]};
}

/** Read calculated inputs together; date formats preserve the existing date validation. */
function readRepaymentInputs(spreadsheet) {
  const ranges = [
    {name: 'In Loans', row: 3, column: 1},
    {name: 'Params', row: 2, column: 2, width: 2},
    {name: 'Cash Flow', row: 3, column: 2},
    {name: 'Out Loans', row: 3, column: 8},
    {name: 'Construction Exp (P)', row: 3, column: 2},
    {name: 'Units', row: 2, column: 4, width: 13}
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


function repaymentHeaderIndex(headers, name) {
  const index = headers.indexOf(name);
  if (index < 0 || headers.lastIndexOf(name) !== index) {
    throw new Error('Expected exactly one In Loans header: ' + name + '. Nothing was written.');
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
