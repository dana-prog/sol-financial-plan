const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../src/RepaymentCalculations.js'), 'utf8');
const context = vm.createContext({Date});
vm.runInContext(source, context);
const day = (year, month, date = 1) => context.repaymentDay(year, month - 1, date);
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-7, `${actual} != ${expected}`);
const loan = (overrides = {}) => ({row: 4, name: 'Loan', priority: 1, date: day(2026, 1), rate: 0, draws: [100], ...overrides});
const model = (overrides = {}) => ({months: [day(2026, 1)], loans: [loan()], receipts: [0], expenses: [0], reserve: 0, repaymentStart: day(2026, 1), ...overrides});
const calc = value => context.calculateIncomingRepayments(value);
const state = (overrides = {}) => ({principal: 0, interest: 0, capitalized: 0, interestPaid: 0, ...overrides});

test('preserves the reserve and caps principal payments at the debt owed', () => {
  const result = calc(model({reserve: 40}));
  near(result.payments[0].principal[0], 60);
  near(result.cash[0], 40);
  const surplus = calc(model({receipts: [50]}));
  near(surplus.payments[0].principal[0], 100);
  near(surplus.cash[0], 50);
});

test('starts repayments in the configured month and carries debt and cash forward', () => {
  const result = calc(model({months: [day(2026, 1), day(2026, 2)], loans: [loan({draws: [100, 0]})],
    receipts: [0, 0], expenses: [0, 0], repaymentStart: day(2026, 2)}));
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

test('a disabled planned loan with zero draws produces no debt or repayments', () => {
  const result = calc(model({loans: [loan({rate: 0.07, draws: [0]})], receipts: [1000]}));
  near(result.payments[0].principal[0], 0);
  near(result.payments[0].interest[0], 0);
  near(result.cash[0], 1000);
});

test('prorates initial interest including the draw date', () => {
  const result = context.accrueIncomingLoanMonth(loan({date: day(2026, 1, 20), rate: 0.365}), state(), day(2026, 1), 1000);
  near(result.interest, 12);
});

test('capitalizes interest on the anniversary and compounds only the remaining days', () => {
  const result = context.accrueIncomingLoanMonth(loan({date: day(2025, 1, 20), rate: 0.365}),
    state({principal: 1000, interest: 100, capitalized: 80, interestPaid: 30}), day(2026, 1), 0);
  near(result.capitalized, 119.95);
  near(result.interest, 100 + 1050 * 31 / 1000 + (119.95 - 50) * 12 / 1000);
});

test('clamps leap-day anniversary to February 28 and uses Actual/365', () => {
  const result = context.accrueIncomingLoanMonth(loan({date: day(2024, 2, 29), rate: 0.365}),
    state({principal: 1000, interest: 10}), day(2025, 2), 0);
  near(result.capitalized, 37);
  near(result.interest, 10 + 28 + 0.037);
});

test('no interest is generated before the draw month or without a loan date', () => {
  near(context.accrueIncomingLoanMonth(loan({date: day(2026, 2), rate: 0.1}), state(), day(2026, 1), 0).interest, 0);
  near(context.accrueIncomingLoanMonth(loan({date: null}), state(), day(2026, 1), 0).interest, 0);
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
    receipts: months.map(() => 1000), expenses: months.map(() => 100), reserve: 15000, repaymentStart: day(2027, 1)});
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

test('server action writes only incoming payment rows and releases its lock', () => {
  const writes = [];
  let released = false;
  const incoming = {getRange(row, column, height, width) {
    if (row === 'A1') return {clearDataValidations() {return this;}, setValue(value) {writes.push({status: value}); return this;}, setNote() {return this;}};
    return {setValues(values) {writes.push({row, column, height, width, values});}};
  }};
  const server = vm.createContext({Date,
    SpreadsheetApp: {getActiveSpreadsheet: () => ({getSheetByName(name) {assert.equal(name, 'In Loans'); return incoming;}, getSpreadsheetTimeZone: () => 'Asia/Bangkok'}), flush() {}},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock() {released = true;}})},
    Utilities: {formatDate: () => '11 Nov 15:00'}
  });
  vm.runInContext(source, server);
  server.readIncomingRepaymentModel = () => model();
  server.recalculateIncomingLoanRepayments();
  assert.deepEqual(writes.filter(x => x.row).map(x => x.row), [10, 13]);
  assert.ok(writes.at(-1).status.startsWith('Repayments:'));
  assert.equal(released, true);
});

test('input failure causes no writes and releases the lock', () => {
  let released = false;
  const server = vm.createContext({Date,
    SpreadsheetApp: {getActiveSpreadsheet: () => ({getSheetByName() {throw new Error('Must not write');}})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock() {released = true;}})}
  });
  vm.runInContext(source, server);
  server.readIncomingRepaymentModel = () => {throw new Error('Bad input');};
  assert.throws(() => server.recalculateIncomingLoanRepayments(), /Bad input/);
  assert.equal(released, true);
});

test('planned-loan edits tell the user to recalculate and the former repayment dropdown is ignored', () => {
  const dialogs = [];
  const server = vm.createContext({IN_LOANS_SHEET_ID: 1346380865,
    HtmlService: {createTemplateFromFile() {return {evaluate() {dialogs.push({action: this.action, message: this.completionMessage}); return {setWidth() {return this;}, setHeight() {return this;}};}};}},
    SpreadsheetApp: {getUi: () => ({showModalDialog() {}, alert() {throw new Error('Unexpected validation failure');}})}
  });
  vm.runInContext(fs.readFileSync(require('node:path').join(__dirname, '../src/LoanModes.js'), 'utf8'), server);
  const event = (column, value) => ({value, range: {getRow: () => 1, getColumn: () => column, getGridId: () => 1346380865}});
  assert.equal(server.handleLoanModeEdit(event(1, 'Repayment Calculations Enabled')), false);
  assert.equal(server.handleLoanModeEdit(event(4, 'Planned Loans Disabled')), true);
  assert.match(dialogs[0].message, /excluded from cash flow/);
  assert.match(dialogs[0].message, /Recalculate Repayments/);
  assert.equal(server.handleLoanModeEdit(event(4, 'Planned Loans Enabled')), true);
  assert.match(dialogs[1].message, /included in cash flow/);
});
