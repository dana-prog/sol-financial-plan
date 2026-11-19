const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');
const ctx = vm.createContext({});
vm.runInContext(fs.readFileSync(`${__dirname}/../src/AllExpenses.js`, 'utf8'), ctx);
vm.runInContext(fs.readFileSync(`${__dirname}/../src/DetailsBorders.js`, 'utf8'), ctx);
const plain = value => JSON.parse(JSON.stringify(value));
const item = (type, year, quarter, month, cost) => [type, year, quarter, month, 'Item (detail)', cost];

test('type subtotals follow the defined order and sum only detail rows', () => {
  const layout = plain(ctx._allExpensesLayout([
    item('Land Expenses', 2027, 'Q1 2027', 46500, 50),
    item('Borrowing Repayments', 2028, 'Q1 2028', 46800, 20),
    item('Borrowing Repayments', 2027, 'Q4 2027', 46700, 10),
  ]));
  assert.deepEqual([...new Set(layout.rows.map(row => row[0]))], ['Borrowing Repayments', 'Land Expenses', 'Total']);
  assert.equal(layout.rows[0][5], 10);
  assert.deepEqual(layout.subtotals, [3, 4, 6, 7, 8, 10, 11, 12]);
  assert.equal(layout.rows[6][5], '=SUBTOTAL(9,F2:F7)');
  assert.equal(layout.rows[10][5], '=SUBTOTAL(9,F9:F11)');
  // Function 9 ignores filtered rows and nested SUBTOTALs, but retains collapsed groups.
  assert.equal(layout.rows[11][5], '=SUBTOTAL(9,F2:F12)');
});

test('nested sections identify the type, year and quarter merge boundaries', () => {
  const layout = plain(ctx._allExpensesLayout([
    item('Land Expenses', 2027, 'Q1 2027', 46400, 10),
    item('Land Expenses', 2027, 'Q2 2027', 46500, 10),
    item('Land Expenses', 2028, 'Q1 2028', 46800, 20),
  ]));
  assert.deepEqual(layout.sections, [
    {startRow: 2, endRow: 2, depth: 3},
    {startRow: 4, endRow: 4, depth: 3},
    {startRow: 2, endRow: 5, depth: 2},
    {startRow: 7, endRow: 7, depth: 3},
    {startRow: 7, endRow: 8, depth: 2},
    {startRow: 2, endRow: 9, depth: 1},
  ]);
  assert.equal(layout.subtotals.at(-1), 10);
  // Summary rows separate sibling groups, otherwise Sheets merges adjacent groups.
  for (const depth of [1, 2, 3]) {
    const groups = layout.sections.filter(group => group.depth === depth);
    for (let i=1;i<groups.length;i++) assert.ok(groups[i].startRow > groups[i-1].endRow + 1);
  }
});

test('empty plans produce a zero total; invalid calculation results are rejected', () => {
  assert.deepEqual(plain(ctx._allExpensesLayout([])), {
    rows: [['Total', '', '', '', '', '=0']], subtotals: [], sections: [], totalRow: 2,
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

test('atomic display write preserves filters and merges type, year and quarter labels', () => {
  const sheet = {properties: {sheetId: 12, gridProperties: {rowCount: 20}},
    basicFilter: {filterSpecs: [{columnIndex: 1, filterCriteria: {hiddenValues: ['', '2028']}}]}};
  const layout = ctx._allExpensesLayout([item('Land Expenses', 2027, 'Q1 2027', 46400, 50)]);
  const requests = plain(ctx._allExpensesWriteRequests(sheet, layout));
  const write = requests.find(request => request.updateCells).updateCells;
  assert.equal(write.rows[0].values[5].userEnteredValue.numberValue, 50);
  assert.equal(write.rows[1].values[5].userEnteredValue.formulaValue, '=SUBTOTAL(9,F2:F2)');
  assert.deepEqual(write.rows[0].values[4].textFormatRuns, [
    {startIndex: 0, format: {fontSize: 11}}, {startIndex: 5, format: {fontSize: 10}},
  ]);
  const filter = requests.find(request => request.setBasicFilter).setBasicFilter.filter;
  assert.deepEqual(filter.filterSpecs[0].filterCriteria.hiddenValues, ['2028']);
  assert.equal(filter.range.endRowIndex, 5);
  assert.equal(requests.filter(request => request.addDimensionGroup).length, 0);
  assert.deepEqual(requests.filter(request => request.mergeCells).map(request => request.mergeCells.range), [
    {sheetId: 12, startRowIndex: 1, endRowIndex: 3, startColumnIndex: 2, endColumnIndex: 3},
    {sheetId: 12, startRowIndex: 1, endRowIndex: 4, startColumnIndex: 1, endColumnIndex: 2},
    {sheetId: 12, startRowIndex: 1, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 1},
  ]);
  assert.equal(requests.at(-1).updateSheetProperties.properties.gridProperties.rowCount, 6);
  assert.deepEqual(sheet.basicFilter.filterSpecs[0].filterCriteria.hiddenValues, ['', '2028']);
});
