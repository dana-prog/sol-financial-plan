const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const ctx = vm.createContext({});
for (const file of ['AllExpensesFormula.js', 'DetailsBorders.js', 'AllExpenses.js']) {
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/${file}`, 'utf8'), ctx);
}
const headers = vm.runInContext('ALL_EXPENSES_HEADERS', ctx);
function snapshot() {
  const data = ['Setup', 'Planning', 'Design', 'Actual', 2026, 'Q1', 46023, -100, 0, -100, 'Setup Exp (A)', 'Design', 4];
  const cells = values => values.map(value => ({effectiveValue: {[typeof value === 'number' ? 'numberValue' : 'stringValue']: value}, formattedValue: String(value)}));
  const rows = [{values: cells(headers)}, {values: cells(data)}];
  rows[0].values[0].userEnteredValue = {formulaValue: '=LIVE_EXPENSES()'};
  return {
    properties: {sheetId: 1, gridProperties: {rowCount: 100, columnCount: 13}},
    basicFilter: {range: {sheetId: 1}, sortSpecs: [{dimensionIndex: 1, sortOrder: 'ASCENDING'}], criteria: {'0': {hiddenValues: ['Land']}}},
    data: [{rowData: rows, rowMetadata: [{}, {hiddenByUser: true}]}],
  };
}

test('formatting refresh preserves the live formula, signed records and filter criteria', () => {
  const sheet = snapshot();
  const before = JSON.stringify(sheet);
  const requests = ctx._allExpensesRefreshRequests(sheet);
  assert.equal(JSON.stringify(sheet), before);
  for (const request of requests) {
    assert.ok(!request.updateCells && !request.updateDimensionProperties && !request.updateSheetProperties);
    if (request.repeatCell) assert.equal(request.repeatCell.fields, 'userEnteredFormat.numberFormat');
  }
  const filter = requests.find(r => r.setBasicFilter).setBasicFilter.filter;
  assert.deepEqual(filter.criteria, sheet.basicFilter.criteria);
  assert.equal(filter.sortSpecs, undefined);
  assert.equal(filter.range.endColumnIndex, 13);
  assert.equal(filter.range.endRowIndex, 100);
  assert.equal(requests.find(r => r.repeatCell?.range.startColumnIndex === 6).repeatCell.cell.userEnteredFormat.numberFormat.pattern, 'mmmm');
});

test('invalid or static datasets fail before any write', () => {
  const sheet = snapshot();
  delete sheet.data[0].rowData[0].values[0].userEnteredValue;
  assert.throws(() => ctx._allExpensesRefreshRequests(sheet), /live formula/);
  const error = snapshot();
  error.data[0].rowData[1].values[7].effectiveValue = {errorValue: {message: 'Source amount is text'}};
  assert.throws(() => ctx._allExpensesRefreshRequests(error), /row 2: Source amount is text/);
  error.data[0].rowData[1].values[7].effectiveValue = {errorValue: {type: 'N_A'}};
  assert.throws(() => ctx._allExpensesRefreshRequests(error), /row 2: N_A/);
  const anchorError = snapshot();
  anchorError.data[0].rowData[0].values[0].effectiveValue = {errorValue: {type: 'N_A'}};
  assert.throws(() => ctx._allExpensesRefreshRequests(anchorError), /row 1: N_A/);
  const wrongHeader = snapshot();
  wrongHeader.data[0].rowData[0].values[5].effectiveValue.stringValue = 'Total';
  assert.throws(() => ctx._allExpensesRefreshRequests(wrongHeader), /headers/);
});

test('empty live dataset can be formatted without adding static totals', () => {
  const sheet = snapshot();
  sheet.data[0].rowData.splice(1);
  assert.ok(ctx._allExpensesRefreshRequests(sheet).every(request => !request.updateCells));
});

test('refresh releases the lock and never reads or rewrites raw expense sources', () => {
  let released = false;
  let saved;
  ctx.SpreadsheetApp = {getActiveSpreadsheet: () => ({getId: () => 'book'}), flush() {}};
  ctx.LockService = {getDocumentLock: () => ({tryLock: () => true, releaseLock: () => {released = true;}})};
  ctx.Sheets = {Spreadsheets: {
    get(id, options) {assert.equal(id, 'book'); assert.equal(options.ranges[0], "'All Expenses'"); return {sheets: [snapshot()]};},
    batchUpdate(body) {saved = body;},
  }};
  assert.match(ctx.refreshAllExpenses(), /updates automatically/);
  assert.ok(saved.requests.length && released);
  released = false;
  ctx.Sheets.Spreadsheets.get = () => {throw new Error('Read failed');};
  assert.throws(() => ctx.refreshAllExpenses(), /Read failed/);
  assert.ok(released);
});
