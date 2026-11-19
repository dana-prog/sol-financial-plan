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
