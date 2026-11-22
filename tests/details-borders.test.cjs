const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');

function context(globals = {}) {
  const ctx = vm.createContext({ BORDER_COLOR: '#666666', ...globals });
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/DetailsBorders.js`, 'utf8'), ctx);
  return ctx;
}

test('type boundaries are thick, year boundaries thin, and quarter changes have no divider', () => {
  const ctx = context();
  const values = [
    ['Type', 'Year', 'Quarter'],
    ['In Loan Repayments', '2028', 'Q1 2028'],
    ['In Loan Repayments', '2028', 'Q1 2028'],
    ['In Loan Repayments', '2028', 'Q2 2028'],
    ['In Loan Repayments', '2029', 'Q1 2029'],
    ['Out Loan Disbursements', '2027', 'Q4 2027'],
    ['', '', ''],
  ];
  assert.deepEqual(JSON.parse(JSON.stringify(ctx._detailsBorderRows(values))), [
    { row: 4, style: 'SOLID' },
    { row: 5, style: 'SOLID_THICK' },
    { row: 6, style: 'SOLID_THICK' },
  ]);
});

test('group columns are identified by header, and missing headers fail before writes', () => {
  const ctx = context();
  assert.deepEqual(JSON.parse(JSON.stringify(ctx._detailsBorderRows([
    ['Quarter', ' TYPE ', 'Year'], ['Q1 2028', 'Land Expenses', '2028'],
  ]))), [{ row: 2, style: 'SOLID_THICK' }]);
  let writes = 0;
  const api = {get: () => sheetData([['type', 'quarter', 'month'], ['Land', 'Q2 2028', 'May']]),
    batchUpdate: () => writes++};
  const command = context({Sheets: {Spreadsheets: {get: () => ({sheets: [api.get()]}), batchUpdate: api.batchUpdate}}});
  assert.throws(() => command.setDetailsSheetBorders(sheetHandle), /No borders were changed/);
  assert.equal(writes, 0);
});

test('expense headers and subtotal rows retain year dividers without a spurious subtotal divider', () => {
  const values = [
    ['Expense Type', 'Year'], ['Land Expenses', '2027'], ['Land Expenses', '2028'],
    ['Land Expenses', ''], ['Total', ''],
  ];
  assert.deepEqual(JSON.parse(JSON.stringify(context()._detailsBorderRows(values))), [
    {row: 2, style: 'SOLID'}, {row: 4, style: 'SOLID_THICK'}, {row: 5, style: 'SOLID_THICK'},
  ]);
});

const plain = value => JSON.parse(JSON.stringify(value));
const sheetHandle = {getName: () => 'All Expenses', getParent: () => ({getId: () => 'copy-id'})};
function sheetData(values, merges = [], rowMetadata = []) {
  return {properties: {sheetId: 12}, merges, data: [{
    rowData: values.map(row => ({values: row.map(value => ({formattedValue: value}))})), rowMetadata,
  }]};
}
const merged = (startRowIndex, endRowIndex, column) => ({
  sheetId: 12, startRowIndex, endRowIndex, startColumnIndex: column, endColumnIndex: column + 1,
});

test('border command uses one read and one atomic horizontal border update', () => {
  const calls = [];
  const snapshot = sheetData([
    ['Expense Type', 'Year', 'Quarter', 'Month', 'Item', 'Cost'],
    ['Land', '2028', 'Q1 2028', '', 'Item', '10'],
    ['Land', '2029', 'Q1 2029', '', 'Item', '20'],
  ]);
  const ctx = context({Sheets: {Spreadsheets: {
    get: (id, options) => {calls.push({id, options}); return {sheets: [snapshot]};},
    batchUpdate: (body, id) => calls.push({id, body: plain(body)}),
  }}});
  ctx.setDetailsSheetBorders(sheetHandle);
  assert.equal(calls.length, 2);
  assert.equal(calls[0].id, 'copy-id');
  assert.deepEqual(Array.from(calls[0].options.ranges), ["'All Expenses'"]);
  const requests = calls[1].body.requests;
  assert.ok(requests.every(request => Object.keys(request).join() === 'updateBorders'));
  assert.ok(requests.every(({updateBorders: border}) => !border.left && !border.right && !border.innerVertical));
  assert.deepEqual(requests[0].updateBorders, {
    range: {sheetId: 12, startRowIndex: 1, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 6},
    top: {style: 'NONE'}, bottom: {style: 'NONE'}, innerHorizontal: {style: 'NONE'},
  });
  assert.equal(requests[1].updateBorders.top.style, 'SOLID_THICK');
  assert.equal(requests[2].updateBorders.bottom.style, 'SOLID');
  assert.equal(requests[3].updateBorders.top.style, 'SOLID');
  assert.equal(requests[4].updateBorders.bottom.style, 'SOLID_THICK');
  for (const {updateBorders: border} of requests.slice(1)) {
    assert.deepEqual((border.top || border.bottom).color, {red: 0.4, green: 0.4, blue: 0.4});
  }
});

test('category edges cover whole merged labels; year edges do not cut through a category', () => {
  const sheet = sheetData([
    ['Expense Type', 'Year', 'Quarter', 'Month', 'Item', 'Cost'],
    ['Land', '2027', 'Q1', '', 'Item', '10'],
    ['', '', '', '', 'Item', '20'],
    ['', '2028', 'Q1', '', 'Item', '30'],
    ['Total', '', '', '', '', '60'],
  ], [merged(1, 4, 0), merged(1, 3, 1), merged(1, 3, 2)]);
  const before = JSON.stringify(sheet);
  const borders = plain(context()._detailsBorderRequests(sheet)).map(request => request.updateBorders);
  assert.equal(JSON.stringify(sheet), before);
  assert.ok(borders.some(border => border.bottom?.style === 'SOLID_THICK' &&
    JSON.stringify(border.range) === JSON.stringify(merged(1, 4, 0))));
  assert.ok(borders.some(border => border.bottom?.style === 'SOLID' &&
    border.range.startRowIndex === 1 && border.range.endRowIndex === 3 && border.range.startColumnIndex === 1));
  assert.ok(borders.filter(border => border.top?.style === 'SOLID' || border.bottom?.style === 'SOLID')
    .every(border => border.range.startColumnIndex >= 1));
});

test('filtered category and year boundaries are placed on their visible rows', () => {
  const sheet = sheetData([
    ['Expense Type', 'Year', 'Quarter', 'Month', 'Item', 'Cost'],
    ['Land', '2026', 'Q1', '', 'Item', '10'],
    ['', '2027', 'Q1', '', 'Item', '20'],
    ['', '2028', 'Q1', '', 'Item', '30'],
    ['Setup', '2026', 'Q1', '', 'Item', '40'],
    ['', '2027', 'Q1', '', 'Item', '50'],
    ['Total', '', '', '', '', '150'],
  ], [merged(1, 4, 0), merged(4, 6, 0)], [
    {}, {hiddenByFilter: true}, {}, {hiddenByFilter: true}, {hiddenByUser: true}, {}, {},
  ]);
  const borders = plain(context()._detailsBorderRequests(sheet)).map(request => request.updateBorders);
  assert.ok(borders.some(border => border.bottom?.style === 'SOLID_THICK' &&
    border.range.startRowIndex === 2 && border.range.endRowIndex === 3 && border.range.endColumnIndex === 6));
  assert.ok(borders.some(border => border.top?.style === 'SOLID_THICK' &&
    border.range.startRowIndex === 5 && border.range.endRowIndex === 6 && border.range.endColumnIndex === 6));
  assert.ok(!borders.some(border => border.top?.style === 'SOLID' || border.bottom?.style === 'SOLID'));
});

test('quarter changes within one year receive no horizontal divider', () => {
  const sheet = sheetData([
    ['Expense Type', 'Year', 'Quarter', 'Month', 'Item', 'Cost'],
    ['Land', '2027', 'Q1', '', 'Item', '10'],
    ['', '', 'Q2', '', 'Item', '20'],
  ], [merged(1, 3, 0), merged(1, 3, 1)]);
  const borders = plain(context()._detailsBorderRequests(sheet)).slice(1).map(request => request.updateBorders);
  assert.ok(!borders.some(border => border.bottom && border.range.endRowIndex === 2));
  assert.ok(!borders.some(border => border.top && border.range.startRowIndex === 2));
});

test('the menu supports both expense sheet names, including renamed workbook copies', () => {
  let activeName = 'Details';
  const calls = [];
  const sheet = { getName: () => activeName };
  const ctx = context({
    SOLLibrary: { deleteTmpExportResources() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getActiveSheet: () => sheet }) },
    setTimelineSheetBorders: () => calls.push('timeline'),
  });
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/Main.js`, 'utf8'), ctx);
  ctx.setDetailsSheetBorders = target => { assert.equal(target, sheet); calls.push('details'); };
  ctx._onSetTimelineBorders();
  activeName = 'All Expenses';
  ctx._onSetTimelineBorders();
  activeName = 'Operational Exp (P)';
  ctx._onSetTimelineBorders();
  assert.deepEqual(calls, ['details', 'details', 'timeline']);
});

test('the reordered expense dataset groups by Budget Category and finds Year after Basis', () => {
  const values = [
    ['Budget Category', 'Budget Line', 'Param', 'Basis', 'Year', 'Quarter', 'Month'],
    ['Setup', 'Planning', 'Design', 'Actual', '2026', 'Q1', 'January'],
    ['Setup', 'Legal', 'Permit', 'Actual', '2026', 'Q2', 'April'],
    ['Setup', 'Legal', 'Permit', 'Planned', '2027', 'Q1', 'January'],
    ['Infra', 'Water', 'Tank', 'Planned', '2028', 'Q1', 'February'],
  ];
  assert.deepEqual(plain(context()._detailsBorderRows(values)), [
    {row: 3, style: 'SOLID'}, {row: 4, style: 'SOLID_THICK'}, {row: 5, style: 'SOLID_THICK'},
  ]);
  assert.deepEqual(plain(context()._detailsBorderRows(values, new Set([4]))), [
    {row: 3, style: 'SOLID_THICK'}, {row: 5, style: 'SOLID_THICK'},
  ]);
});
