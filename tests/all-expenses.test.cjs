const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');
const ctx = vm.createContext({});
vm.runInContext(fs.readFileSync(`${__dirname}/../src/AllExpenses.js`, 'utf8'), ctx);
vm.runInContext(fs.readFileSync(`${__dirname}/../src/DetailsBorders.js`, 'utf8'), ctx);
const plain = value => JSON.parse(JSON.stringify(value));
const item = (type, year, quarter, month, cost) => [type, year, quarter, month, 'Item (detail)', cost];

test('details follow the defined type order with only one grand total', () => {
  const layout = plain(ctx._allExpensesLayout([
    item('Land Expenses', 2027, 'Q1 2027', 46500, 50),
    item('Borrowing Repayments', 2028, 'Q1 2028', 46800, 20),
    item('Borrowing Repayments', 2027, 'Q4 2027', 46700, 10),
  ]));
  assert.deepEqual([...new Set(layout.rows.map(row => row[0]))], ['Borrowing Repayments', 'Land Expenses', 'Total']);
  assert.equal(layout.rows[0][5], 10);
  assert.equal(layout.rows.length, 4);
  assert.deepEqual(layout.rows.slice(0, -1).map(row => row[5]), [10, 20, 50]);
  assert.equal(layout.rows[3][5], '=SUBTOTAL(9,F2:F4)');
  assert.equal(layout.rows.filter(row => typeof row[5] === 'string').length, 1);
});

test('empty plans produce a zero total; invalid calculation results are rejected', () => {
  assert.deepEqual(plain(ctx._allExpensesLayout([])), {
    rows: [['Total', '', '', '', '', '=0']], totalRow: 2,
  });
  assert.throws(() => ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1', 1, '#VALUE!')]), /invalid rows/);
  assert.throws(() => ctx._allExpensesLayout([item('Unknown', 2027, 'Q1', 1, 1)]), /invalid rows/);
});

test('parenthesis font spans include delimiters, nested details and multiple pairs', () => {
  const text = 'Clinic (large (east)) and room (west)';
  const spans = plain(ctx._parenthesisSpans(text));
  assert.deepEqual(spans.map(([start,end]) => text.slice(start,end)), ['(large (east))', '(west)']);
  assert.deepEqual(plain(ctx._parenthesisSpans('Plain text')), []);
  assert.deepEqual(plain(ctx._parenthesisSpans('Unmatched (detail')), []);
});

test('rich text switches to 10pt only inside parentheses and has unique ordered offsets', () => {
  assert.deepEqual(plain(ctx._allExpensesTextRuns('Item (detail) rest')), [
    {startIndex: 0, format: {fontSize: 11}}, {startIndex: 5, format: {fontSize: 10}},
    {startIndex: 13, format: {fontSize: 11}},
  ]);
  const adjacent = plain(ctx._allExpensesTextRuns('(a)(b)'));
  assert.deepEqual(adjacent.map(run => run.startIndex), [0, 3]);
  assert.ok(adjacent.every(run => run.format.fontSize === 10));
});

test('atomic display write preserves filters and repeats type, year and quarter labels without merges', () => {
  const sheet = {properties: {sheetId: 12, gridProperties: {rowCount: 20}},
    merges: [{sheetId: 12, startRowIndex: 1, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 1}],
    basicFilter: {filterSpecs: [{columnIndex: 1, filterCriteria: {hiddenValues: ['', '2028']}}]}};
  const layout = ctx._allExpensesLayout([
    item('Land Expenses', 2027, 'Q1 2027', 46400, 50),
    item('Land Expenses', 2027, 'Q1 2027', 46401, 60),
  ]);
  const requests = plain(ctx._allExpensesWriteRequests(sheet, layout));
  const write = requests.find(request => request.updateCells).updateCells;
  assert.equal(write.rows[0].values[5].userEnteredValue.numberValue, 50);
  assert.equal(write.rows[2].values[5].userEnteredValue.formulaValue, '=SUBTOTAL(9,F2:F3)');
  assert.deepEqual(write.rows[0].values[4].textFormatRuns, [
    {startIndex: 0, format: {fontSize: 11}}, {startIndex: 5, format: {fontSize: 10}},
  ]);
  const filter = requests.find(request => request.setBasicFilter).setBasicFilter.filter;
  assert.deepEqual(filter.filterSpecs[0].filterCriteria.hiddenValues, ['', '2028']);
  assert.equal(filter.range.endRowIndex, 3);
  assert.equal(requests.filter(request => request.addDimensionGroup).length, 0);
  assert.equal(requests.filter(request => request.mergeCells).length, 0);
  assert.deepEqual(requests.filter(request => request.unmergeCells).map(request => request.unmergeCells.range), sheet.merges);
  for (const row of write.rows.slice(0, -1)) {
    assert.deepEqual(row.values.slice(0, 3).map(cell => cell.userEnteredValue), [
      {stringValue: 'Land Expenses'}, {numberValue: 2027}, {stringValue: 'Q1 2027'},
    ]);
  }
  assert.equal(requests.at(-1).updateSheetProperties.properties.gridProperties.rowCount, 4);
  assert.deepEqual(sheet.basicFilter.filterSpecs[0].filterCriteria.hiddenValues, ['', '2028']);
});

function sheetSnapshot(layout) {
  const sheet = {properties: {sheetId: 12, title: 'All Expenses', gridProperties: {rowCount: layout.totalRow}},
    basicFilter: {range: {sheetId: 12, startRowIndex: 0, endRowIndex: layout.totalRow - 1, startColumnIndex: 0, endColumnIndex: 6},
      filterSpecs: [{columnIndex: 1, filterCriteria: {hiddenValues: ['2028']}}]}};
  const write = ctx._allExpensesWriteRequests(sheet, layout).find(request => request.updateCells).updateCells;
  sheet.data = [{rowData: [{values: [{userEnteredValue: {stringValue: 'Expense Type'}}]}, ...plain(write.rows)]}];
  return sheet;
}

test('unchanged labels retain formatting while refreshed costs and rich text are written', () => {
  const layout = ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1 2027', 46400, 50)]);
  const sheet = sheetSnapshot(layout);
  const next = ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1 2027', 46400, 75)]);
  const requests = plain(ctx._allExpensesRefreshRequests(sheet, next, 20000));
  assert.equal(requests.some(request => request.updateBorders), false);
  assert.equal(requests.some(request => request.repeatCell?.fields.includes('userEnteredFormat')), false);
  const write = requests.find(request => request.updateCells).updateCells;
  assert.equal(write.rows[0].values[5].userEnteredValue.numberValue, 75);
  assert.equal(write.rows[0].values[4].textFormatRuns[1].format.fontSize, 10);
  assert.equal(write.rows[1].values[5].userEnteredValue.formulaValue, '=SUBTOTAL(9,F2:F2)');
  assert.deepEqual(requests.find(request => request.setBasicFilter).setBasicFilter.filter.filterSpecs, sheet.basicFilter.filterSpecs);
  assert.equal(requests.at(-1).updateSheetProperties.properties.gridProperties.rowCount, 3);
});

test('changed labels rebuild the formatting and remove unused calculation rows', () => {
  const previous = ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1 2027', 46400, 50)]);
  const sheet = sheetSnapshot(previous);
  const next = ctx._allExpensesLayout([]);
  const requests = plain(ctx._allExpensesRefreshRequests(sheet, next, 20000));
  assert.ok(requests.some(request => request.repeatCell?.fields === 'userEnteredFormat'));
  assert.equal(requests.at(-1).updateSheetProperties.properties.gridProperties.rowCount, 2);
});

test('calculation uses the existing table without adding sheets or columns', () => {
  vm.runInContext('var ALL_EXPENSES_FORMULA = "=NAMED_FUNCTION()";', ctx);
  const layout = ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1 2027', 46400, 50)]);
  const sheet = sheetSnapshot(layout);
  const requests = plain(ctx._allExpensesCalculationRequests(sheet, 20000));
  assert.equal(requests.some(request => request.addSheet || request.deleteSheet), false);
  assert.equal(requests.find(request => request.appendDimension).appendDimension.dimension, 'ROWS');
  assert.equal(requests.at(-1).updateCells.start.rowIndex, 1);
  assert.equal(requests.at(-1).updateCells.start.columnIndex, 0);
  assert.equal(requests.at(-1).updateCells.rows[0].values[0].userEnteredValue.formulaValue, '=NAMED_FUNCTION()');
});

test('ready calculation responses need no extra API reads and ignore trailing formatted blanks', () => {
  const expense = item('Land Expenses', 2027, 'Q1 2027', 46400, 50);
  const grid = [{values: expense.map(value => ({effectiveValue: typeof value === 'number' ? {numberValue: value} : {stringValue: value}}))}, {values: [{}, {}]}];
  const timings = {};
  const result = {updatedSpreadsheet: {sheets: [{data: [{rowData: grid}]}]}};
  assert.deepEqual(plain(ctx._readCalculatedExpenses('id', 'range', result, timings)), [expense]);
  assert.equal(timings.polls, 0);
  const empty = {updatedSpreadsheet: {sheets: [{data: [{rowData: [{values: [{effectiveValue: {stringValue: '__NO_EXPENSES__'}}]}]}]}]}};
  assert.deepEqual(plain(ctx._readCalculatedExpenses('id', 'range', empty, {})), []);
});

test('formula errors fail immediately instead of waiting for the polling timeout', () => {
  const result = {updatedSpreadsheet: {sheets: [{data: [{rowData: [{values: [{effectiveValue: {
    errorValue: {type: 'REF', message: 'Formula cannot expand'},
  }}]}]}]}]}};
  assert.throws(() => ctx._readCalculatedExpenses('id', 'range', result, {}), /Formula cannot expand/);
});

test('a failed calculation restores original values, total formula, rich text, filter and grid size', () => {
  const layout = ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1 2027', 46400, 50)]);
  const snapshot = sheetSnapshot(layout);
  const calls = [];
  const runtime = vm.createContext({
    console: {log() {}},
    SpreadsheetApp: {getActiveSpreadsheet: () => ({getId: () => 'copy-id', getSheetByName: () => ({getName: () => 'All Expenses', getMaxRows: () => 3})})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock() {}})},
    ALL_EXPENSES_FORMULA: '=NAMED_FUNCTION()',
    Sheets: {Spreadsheets: {get: () => ({sheets: [snapshot]}), batchUpdate: request => {
      calls.push(plain(request));
      if (calls.length === 1) return {updatedSpreadsheet: {sheets: [{data: [{rowData: [{values: [{effectiveValue: {
        errorValue: {type: 'ERROR', message: 'Broken calculation'},
      }}]}]}]}]}};
      return {};
    }}},
  });
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/AllExpenses.js`, 'utf8'), runtime);
  assert.throws(() => runtime.refreshAllExpenses(), /Broken calculation/);
  assert.equal(calls.length, 2);
  const restore = calls[1].requests;
  assert.deepEqual(restore[1].updateCells.rows, snapshot.data[0].rowData.slice(1));
  assert.equal(restore[2].updateSheetProperties.properties.gridProperties.rowCount, 3);
  assert.deepEqual(restore.at(-1).setBasicFilter.filter, snapshot.basicFilter);
});
