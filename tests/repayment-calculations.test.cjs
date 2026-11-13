const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../src/RepaymentCalculations.js'), 'utf8');
const context = vm.createContext({Date});
vm.runInContext(source, context);
const day = (year, month, date = 1) => context.repaymentDay(year, month - 1, date);
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} != ${expected}`);
const loan = (overrides = {}) => ({row: 4, name: 'Loan', priority: 1, date: day(2026, 1), rate: 0, draws: [100], interestRepaymentStart: day(2026, 1), ...overrides});
const model = (overrides = {}) => ({months: [day(2026, 1)], loans: [loan()], receipts: [0], expenses: [0], reserve: 0, principalRepaymentStart: day(2026, 1), timelineColumn: 11, ...overrides});
const calc = value => context.calculateBorrowingRepayments(value);
const state = (overrides = {}) => ({principal: 0, interest: 0, capitalized: 0, interestPaid: 0, ...overrides});

test('preserves the reserve and caps principal payments at the debt owed', () => {
  const result = calc(model({reserve: 40}));
  near(result.payments[0].principal[0], 60);
  near(result.cash[0], 40);
  const surplus = calc(model({receipts: [50]}));
  near(surplus.payments[0].principal[0], 100);
  near(surplus.cash[0], 50);
});

test('starts principal repayments in the configured month and carries debt and cash forward', () => {
  const result = calc(model({months: [day(2026, 1), day(2026, 2)], loans: [loan({draws: [100, 0]})],
    receipts: [0, 0], expenses: [0, 0], principalRepaymentStart: day(2026, 2)}));
  assert.deepEqual(Array.from(result.payments[0].principal), [0, 100]);
  assert.deepEqual(Array.from(result.cash), [100, 0]);
});

test('pays all interest before principal and resolves equal priority by sheet row', () => {
  const result = calc(model({loans: [loan({row: 18, rate: 0.365}), loan({row: 4, rate: 0.365})], expenses: [190]}));
  near(result.payments[0].interest[0], 3.1);
  near(result.payments[1].interest[0], 3.1);
  near(result.payments[1].principal[0], 3.8);
  near(result.payments[0].principal[0], 0);
});

test('interest starts in each contract month independently of the principal start', () => {
  const months = [day(2026, 1), day(2026, 2), day(2026, 3)];
  const result = calc(model({months, loans: [
    loan({rate: 0.365, draws: [100, 0, 0], interestRepaymentStart: months[0]}),
    loan({row: 18, rate: 0.365, draws: [100, 0, 0], interestRepaymentStart: months[1]})
  ], receipts: [0, 0, 0], expenses: [0, 0, 0], principalRepaymentStart: months[2]}));
  near(result.payments[0].interest[0], 3.1);
  near(result.payments[1].interest[0], 0);
  near(result.payments[1].interest[1], 5.9);
  for (const payment of result.payments) {
    assert.deepEqual(Array.from(payment.principal).slice(0, 2), [0, 0]);
    assert.ok(payment.interest[2] > 0);
  }
  assert.ok(result.payments.some(payment => payment.principal[2] > 0));
});

test('moving the principal start does not delay contract interest', () => {
  const input = model({months: [day(2026, 1), day(2026, 2)],
    loans: [loan({rate: 0.365, draws: [100, 0]})], receipts: [0, 0], expenses: [0, 0],
    principalRepaymentStart: day(2027, 1)});
  assert.deepEqual(calc(input).payments[0].interest,
    calc({...input, principalRepaymentStart: day(2030, 1)}).payments[0].interest);
});

test('interest before the principal start still respects cash availability and the reserve', () => {
  const result = calc(model({loans: [loan({rate: 0.365})], reserve: 99,
    principalRepaymentStart: day(2027, 1)}));
  near(result.payments[0].interest[0], 1);
  near(result.payments[0].principal[0], 0);
  near(result.cash[0], 99);
});

test('a future interest start does not block otherwise eligible principal', () => {
  const result = calc(model({loans: [loan({rate: 0.365, interestRepaymentStart: day(2027, 1)})]}));
  near(result.payments[0].interest[0], 0);
  near(result.payments[0].principal[0], 100);
});

test('contract interest starts in the anniversary month, including leap-day loans and zero-year terms', () => {
  assert.equal(context.interestRepaymentStartMonth(day(2026, 4, 8), 0.07, 2, 'Loan'), day(2028, 4));
  assert.equal(context.interestRepaymentStartMonth(day(2024, 2, 29), 0.07, 1, 'Loan'), day(2025, 2));
  assert.equal(context.interestRepaymentStartMonth(day(2026, 4, 8), 0.07, 0, 'Loan'), day(2026, 4));
  assert.equal(context.interestRepaymentStartMonth(day(2026, 4, 8), 0, '', 'Loan'), null);
  assert.equal(context.interestRepaymentStartMonth(null, 0.07, '', 'Loan'), null);
});

test('missing or invalid interest terms fail instead of silently skipping payments', () => {
  for (const years of ['', '2', -1, 0.5, NaN, Infinity]) {
    assert.throws(() => context.interestRepaymentStartMonth(day(2026, 1), 0.07, years, 'Contract A'), /Interest repayment start deadline \(years\) field for Contract A/);
  }
});

test('interest uses loan priority when cash cannot cover all interest', () => {
  const result = calc(model({loans: [loan({row: 4, priority: 2, rate: 0.365}), loan({row: 18, priority: 1, rate: 0.365})], expenses: [195]}));
  near(result.payments[1].interest[0], 3.1);
  near(result.payments[0].interest[0], 1.9);
  near(result.payments[0].principal[0] + result.payments[1].principal[0], 0);
});

test('negative cash produces no repayments', () => {
  const result = calc(model({expenses: [150]}));
  near(result.cash[0], -50);
  near(result.payments[0].principal[0], 0);
});

test('a loan with zero draws produces no debt or repayments', () => {
  const result = calc(model({loans: [loan({rate: 0.07, draws: [0]})], receipts: [1000]}));
  near(result.payments[0].principal[0], 0);
  near(result.payments[0].interest[0], 0);
  near(result.cash[0], 1000);
});

test('prorates initial interest including the draw date', () => {
  const result = context.accrueBorrowingMonth(loan({date: day(2026, 1, 20), rate: 0.365}), state(), day(2026, 1), 1000);
  near(result.interest, 12);
});

test('capitalizes interest on the anniversary and compounds only the remaining days', () => {
  const result = context.accrueBorrowingMonth(loan({date: day(2025, 1, 20), rate: 0.365}),
    state({principal: 1000, interest: 100, capitalized: 80, interestPaid: 30}), day(2026, 1), 0);
  near(result.capitalized, 119.95);
  near(result.interest, 100 + 1050 * 31 / 1000 + (119.95 - 50) * 12 / 1000);
});

test('clamps leap-day anniversary to February 28 and uses Actual/365', () => {
  const result = context.accrueBorrowingMonth(loan({date: day(2024, 2, 29), rate: 0.365}),
    state({principal: 1000, interest: 10}), day(2025, 2), 0);
  near(result.capitalized, 37);
  near(result.interest, 10 + 28 + 0.037);
});

test('no interest is generated before the draw month or without a loan date', () => {
  near(context.accrueBorrowingMonth(loan({date: day(2026, 2), rate: 0.1}), state(), day(2026, 1), 0).interest, 0);
  near(context.accrueBorrowingMonth(loan({date: null}), state(), day(2026, 1), 0).interest, 0);
});

test('repeated calculations are deterministic and do not mutate inputs', () => {
  const input = model({loans: [loan({rate: 0.07})], reserve: 20});
  const original = JSON.stringify(input);
  assert.deepEqual(calc(input), calc(input));
  assert.equal(JSON.stringify(input), original);
});

test('repayments and cash reconcile over a multi-year forecast', () => {
  const months = Array.from({length: 96}, (_, i) => day(2026 + Math.floor(i / 12), i % 12 + 1));
  const input = model({months, loans: [loan({rate: 0.07, draws: months.map((_, i) => i ? 0 : 100000)})],
    receipts: months.map(() => 1000), expenses: months.map(() => 100), reserve: 15000, principalRepaymentStart: day(2027, 1)});
  const result = calc(input);
  let cash = 0;
  months.forEach((_, i) => {
    const paid = result.payments[0].principal[i] + result.payments[0].interest[i];
    cash += input.loans[0].draws[i] + input.receipts[i] - input.expenses[i] - paid;
    near(result.cash[i], cash);
    assert.ok(result.payments[0].principal[i] >= 0 && result.payments[0].interest[i] >= 0);
    assert.ok(!paid || cash >= input.reserve - 1e-7);
  });
});

test('invalid calculated amounts fail before any output is returned', () => {
  assert.throws(() => calc(model({receipts: [NaN]})), /invalid amount/);
});

test('server action saves payments, tax and timestamp in one batch and releases its lock', () => {
  const requests = [];
  let released = false;
  let fetches = 0;
  const incoming = {getSheetId: () => 123, getRange: () => ({canEdit: () => true})};
  const server = vm.createContext({Date,
    SpreadsheetApp: {getActiveSpreadsheet: () => ({getSheetByName: () => incoming,
      getId: () => 'test', getSpreadsheetTimeZone: () => 'Asia/Bangkok'})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock() {released = true;}})},
    Utilities: {formatDate: () => '11 Nov 15:00'},
    ScriptApp: {getOAuthToken: () => 'test-token'},
    UrlFetchApp: {fetch(url, options) {
      assert.equal(url, 'https://sheets.googleapis.com/v4/spreadsheets/test:batchUpdate');
      requests.push(...JSON.parse(options.payload).requests);
      fetches++;
      return {getResponseCode: () => 200};
    }}
  });
  vm.runInContext(source, server);
  server.readBorrowingRepaymentModel = () => model({timelineColumn: 14});
  server.calculateRepaymentsWithCompanyTax = value => server.calculateBorrowingRepayments(value);
  const taxRequest = {updateCells: {fields: 'userEnteredValue'}};
  server.companyTaxWriteRequests = () => [taxRequest];
  server.recalculateBorrowingRepayments();
  assert.equal(fetches, 1);
  assert.deepEqual(requests[0], taxRequest);
  assert.deepEqual(requests.slice(1, 3).map(x => x.updateCells.range.startRowIndex), [9, 12]);
  assert.deepEqual(requests.slice(1, 3).map(x => x.updateCells.range.startColumnIndex), [13, 13]);
  assert.ok(requests.at(-1).updateCells.rows[0].values[0].userEnteredValue.stringValue.startsWith('Last Repayment Calculation Time'));
  assert.equal(released, true);
  released = false;
  server.UrlFetchApp.fetch = () => ({getResponseCode: () => 403, getContentText: () => 'Protected range'});
  assert.throws(() => server.recalculateBorrowingRepayments(), /HTTP 403/);
  assert.equal(released, true);
});

test('input failure causes no writes and releases the lock', () => {
  let released = false;
  const server = vm.createContext({Date,
    SpreadsheetApp: {getActiveSpreadsheet: () => ({getSheetByName() {throw new Error('Must not write');}})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock() {released = true;}})}
  });
  vm.runInContext(source, server);
  server.readBorrowingRepaymentModel = () => {throw new Error('Bad input');};
  assert.throws(() => server.recalculateBorrowingRepayments(), /Bad input/);
  assert.equal(released, true);
});

test('display columns do not change loan inputs or repayment calculations', () => {
  const date = new Date('2026-01-01T00:00:00Z');
  const labels = ['start balance', 'loan draw', 'interest generated', 'loan paid', 'end balance',
    'principal owed', 'principal paid', 'principal carried forward', 'interest owed', 'interest paid',
    'interest carried forward', 'interest capitalization month', 'capitalized interest', 'balance bearing interest'];
  const cashLabels = ['land expenses', 'setup expenses', 'construction expenses', 'infra expenses',
    'direct operating expenses', 'other operating expenses', 'loans disbursed',
    'revenue', 'setup depreciation', 'operating profit'];
  function read(extraColumns) {
    const grid = [['Name', 'Priority', 'Date', 'Maturity date', 'Lender', 'Amount', 'Interest', 'Interest repayment start deadline (years)', 'Maturity', 'Status',
      ...Array(extraColumns).fill('Loan details'), 'Param', date],
    ...labels.map((label, i) => [...(i ? Array(10).fill('') : ['Loan', 1, date, date, 'Lender', 100, 0.07, 0, 5, 'Received']),
      ...Array(extraColumns).fill('display only'), label, i === 1 ? 100 : 0])];
    const sheets = {
      'Borrowings': grid,
      Params: [['principal loan repayment start month', date], ['minimum cash reserve', 0],
        ['company tax rate', 0.2], ['financial year-end month', 12]],
      'Cash Flow': [['Param', date], ...cashLabels.map(label => [label, label === 'setup expenses' ? 7 : label === 'loans disbursed' ? 3 : 0])],
      'Loans': [['Param', date], ['total loan paid', 0]]
    };
    const reader = vm.createContext({Date, Utilities: {formatDate: value => value.toISOString().slice(0, 10)}});
    vm.runInContext(source, reader);
    reader.readUnitDepreciation = () => [0];
    reader.readRepaymentInputs = () => sheets;
    return reader.readBorrowingRepaymentModel({});
  }
  const original = read(0);
  assert.equal(original.expenses[0], 10);
  assert.equal(original.timelineColumn, 12);
  const normalize = value => JSON.parse(JSON.stringify(value));
  for (const addedColumns of [2, 3]) {
    const expanded = read(addedColumns);
    assert.equal(expanded.timelineColumn, 12 + addedColumns);
    assert.deepEqual(normalize({...original, timelineColumn: 12 + addedColumns}), normalize(expanded));
    assert.deepEqual(normalize(calc(original)), normalize(calc(expanded)));
  }
});

test('missing or duplicate required headers fail before calculation', () => {
  assert.throws(() => context.repaymentHeaderIndex(['Name'], 'Interest'), /exactly one/);
  assert.throws(() => context.repaymentHeaderIndex(['Interest', 'Interest'], 'Interest'), /exactly one/);
});

test('invalid result rows are rejected before saving; valid writes use the atomic batch permission check', () => {
  const sheet = {getName: () => 'Borrowings', getSheetId: () => 123};
  assert.throws(() => context.repaymentRowWrite(sheet, 10, 14, [NaN], 1), /Invalid result/);
  assert.throws(() => context.repaymentRowWrite(sheet, 10, 14, [1], 2), /Invalid result/);
  assert.equal(context.repaymentRowWrite(sheet, 10, 14, [1], 1).updateCells.range.sheetId, 123);
});

test('batch inputs retain unrounded amounts, booleans, errors and local calendar dates', () => {
  const value = (effectiveValue, type) => context.repaymentInputValue({effectiveValue,
    effectiveFormat: {numberFormat: {type}}});
  assert.equal(value({numberValue: 58850}, 'NUMBER'), 58850);
  assert.equal(value({boolValue: false}), false);
  assert.equal(value({stringValue: '001'}), '001');
  assert.equal(value({errorValue: {type: 'REF'}}), '#REF');
  assert.equal(context.repaymentInputValue(undefined), '');
  const serial = day(2028, 2, 29);
  for (const type of ['DATE', 'DATE_TIME']) {
    const date = value({numberValue: serial + 0.75}, type);
    assert.equal(context.repaymentDate(date, 'UTC', 'Loan'), serial);
  }
  assert.throws(() => context.repaymentDate(value({numberValue: serial}, 'NUMBER'), 'UTC', 'Loan'), /Expected a date/);
});

test('batch reader maps sheets by name and restores omitted empty rows and cells', () => {
  const names = ['Borrowings', 'Params', 'Cash Flow', 'Loans', 'Construction Exp (P)', 'Units'];
  const book = {getId: () => 'test', getSheetByName: name => ({
    getLastRow: () => 4, getLastColumn: () => 16,
    getRange: () => ({getA1Notation: () => 'A1:P4'})
  })};
  let reads = 0;
  const response = {sheets: names.slice().reverse().map(name => ({properties: {title: name},
    data: [{rowData: [{values: [{effectiveValue: {stringValue: name}}]}]}]}))};
  const reader = vm.createContext({Date, ScriptApp: {getOAuthToken: () => 'token'},
    UrlFetchApp: {fetch(url, options) {
      reads++;
      assert.equal(new URL(url).searchParams.getAll('ranges').length, 6);
      assert.equal(options.headers.Authorization, 'Bearer token');
      return {getResponseCode: () => 200, getContentText: () => JSON.stringify(response)};
    }}});
  vm.runInContext(source, reader);
  const actual = reader.readRepaymentInputs(book);
  assert.equal(reads, 1);
  for (const name of names) assert.equal(actual[name][0][0], name);
  assert.equal(actual.Params.length, 3);
  assert.equal(actual.Params[2].length, 2);
  assert.deepEqual(Array.from(actual.Params[2]), ['', '']);
  assert.equal(actual['Loans'][0].length, 8);
  response.sheets.pop();
  assert.throws(() => reader.readRepaymentInputs(book), /Missing repayment inputs/);
  reader.UrlFetchApp.fetch = () => ({getResponseCode: () => 403, getContentText: () => 'Denied'});
  assert.throws(() => reader.readRepaymentInputs(book), /HTTP 403.*Nothing was written/);
});
