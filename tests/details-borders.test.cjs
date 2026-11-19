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
    ['Borrowing Repayments', '2028', 'Q1 2028'],
    ['Borrowing Repayments', '2028', 'Q1 2028'],
    ['Borrowing Repayments', '2028', 'Q2 2028'],
    ['Borrowing Repayments', '2029', 'Q1 2029'],
    ['Loan Disbursements', '2027', 'Q4 2027'],
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
  const sheet = {
    getLastRow: () => 2, getLastColumn: () => 3,
    getRange: () => ({ getMergedRanges: () => [], getDisplayValues: () => [['type', 'quarter', 'month'], ['Land', 'Q2 2028', 'May']],
      setBorder: () => writes++ }),
  };
  assert.throws(() => ctx.setDetailsSheetBorders(sheet), /No borders were changed/);
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

test('border command only changes horizontal borders across the populated width', () => {
  const writes = [];
  const sheet = {
    getLastRow: () => 3, getLastColumn: () => 6,
    getRange: (...range) => ({
      getMergedRanges: () => [],
      getDisplayValues: () => [['type', 'year', 'quarter'], ['Land', '2028', 'Q1 2028'], ['Land', '2029', 'Q1 2029']],
      getA1Notation: () => 'F1',
      setBorder: (...args) => writes.push({ range, args }),
    }),
    getRangeList: ranges => ({ setBorder: (...args) => writes.push({ ranges: Array.from(ranges), args }) }),
  };
  const ctx = context({ SpreadsheetApp: { BorderStyle: {
    SOLID: 'thin', SOLID_THICK: 'thick',
  } } });
  ctx.setDetailsSheetBorders(sheet);
  assert.deepEqual(writes, [
    { range: [2, 1, 2, 6], args: [null, null, false, null, null, false] },
    { ranges: ['A3:F3'], args: [null, null, true, null, null, null, '#666666', 'thick'] },
    { ranges: ['A2:F2'], args: [null, null, true, null, null, null, '#666666', 'thin'] },
  ]);
});

test('merged labels retain the category end and year boundaries when borders are rebuilt', () => {
  const writes = [];
  const sheet = {
    getLastRow: () => 5, getLastColumn: () => 6,
    getRange: () => ({
      getDisplayValues: () => [['Expense Type', 'Year'], ['Land', '2027'], ['', '2028'], ['', ''], ['Total', '']],
      getMergedRanges: () => [{getRow: () => 2, getColumn: () => 1, getNumRows: () => 3, getNumColumns: () => 1}],
      getA1Notation: () => 'F1', setBorder() {},
    }),
    getRangeList: ranges => ({setBorder: () => writes.push(Array.from(ranges))}),
  };
  const ctx = context({SpreadsheetApp: {BorderStyle: {SOLID: 'thin', SOLID_THICK: 'thick'}}});
  ctx.setDetailsSheetBorders(sheet);
  assert.deepEqual(writes, [['A4:F4', 'A5:F5'], ['A2:F2']]);
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
