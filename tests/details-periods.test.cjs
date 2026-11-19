const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');

function context(globals = {}) {
  const ctx = vm.createContext({ Date, ...globals });
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/Details.js`, 'utf8'), ctx);
  return ctx;
}

function fixture() {
  const values = { A2: 'Quarter', C2: 'Dec 2026' };
  const formulas = {};
  const validations = {};
  const sheet = {
    getSheetId: () => 2109063306,
    getMaxRows: () => 4,
    getParent: () => ({ getRangeByName: () => ({ getValues: () => [[new Date(2026, 10, 1), new Date(2028, 4, 1)]] }) }),
    getRange: cell => ({
      getValue: () => values[cell],
      getFormula: () => formulas[cell] || '',
      setFormula(formula) { formulas[cell] = formula; return this; },
      setValue(value) { values[cell] = value; return this; },
      setNumberFormat() { return this; },
      setDataValidation(value) { validations[cell] = value; return this; },

    }),
    activate() { values.activated = true; },
  };
  const ctx = context({
    SOLLibrary: { capitalize: value => value },
    SpreadsheetApp: {
      getActive: () => ({ getSheetById: () => sheet }),
      getUi: () => ({ alert: message => { throw new Error(message); } }),
      flush() {},
      newDataValidation: () => ({
        requireValueInList(options) { this.options = options; return this; },
        setAllowInvalid(value) { this.allowInvalid = value; return this; },
        build() { return this; },
      }),
    },
  });
  ctx._refreshDetailsView = () => {};
  return { ctx, sheet, values, formulas, validations };
}

test('dropdown choices are remembered and the native period formula is restored', () => {
  const { ctx, sheet, values, formulas, validations } = fixture();
  const range = { getSheet: () => sheet, getRow: () => 2, getLastRow: () => 2,
    getColumn: () => 3, getLastColumn: () => 3 };
  values.C2 = 'Q3 2028';
  ctx.updateDetailsPeriod({ range });
  assert.ok(formulas.C2.includes('TO_TEXT("Q3 2028")'));
  assert.ok(formulas.C2.startsWith('=LET('));
  values.A2 = 'Year';
  range.getColumn = () => 1;
  range.getLastColumn = () => 1;
  ctx.updateDetailsPeriod({ range });
  assert.ok(formulas.C2.includes('TO_TEXT("Q3 2028")'));
  assert.equal(values.C1, undefined);
});

function resizeFixture({ currentRows = 100, count = 20, total = 'Total', blankTail = true } = {}) {
  const changes = [];
  const sheet = {
    getMaxRows: () => currentRows,
    getMaxColumns: () => 4,
    deleteRows: (...args) => changes.push(['delete', ...args]),
    getRange: (row, column, rows, columns) => {
      if (row === 5 && columns === 1) return { getValues: () => Array.from({ length: count },
        (_, index) => [index === count - 1 ? total : 'Setup Expenses']) };
      return { isBlank: () => blankTail };
    },
  };
  const ctx = context({ SpreadsheetApp: { flush() {} }, Utilities: { sleep() {} } });
  return { ctx, sheet, changes };
}

test('Details trims only blank rows after Total, including an empty result', () => {
  for (const count of [20, 1]) {
    const { ctx, sheet, changes } = resizeFixture({ count });
    ctx._resizeDetailsRows(sheet);
    assert.deepEqual(changes, [['delete', count + 5, 100 - count - 4]]);
  }
});

test('Details trims automatically expanded sheets without deleting result rows', () => {
  const { ctx, sheet, changes } = resizeFixture({ currentRows: 304, count: 200 });
  ctx._resizeDetailsRows(sheet);
  assert.deepEqual(changes, [['delete', 205, 100]]);
});

test('Details leaves exact-sized tables alone and refuses to delete content or incomplete results', () => {
  const exact = resizeFixture({ currentRows: 24 });
  exact.ctx._resizeDetailsRows(exact.sheet);
  assert.deepEqual(exact.changes, []);
  for (const options of [{ count: 0 }, { total: '#REF!' }, { blankTail: false }]) {
    const { ctx, sheet, changes } = resizeFixture(options);
    assert.throws(() => ctx._resizeDetailsRows(sheet), /Details/);
    assert.deepEqual(changes, []);
  }
});

test('construction formats match source items in formula order without helper identifiers', () => {
  const ctx = context();
  const month = new Date(2028, 0, 1);
  const rows = [
    ['Construction Expenses', month, 1, 10],
    ['Construction Expenses', month, 6, 60],
    ['Construction Expenses', month, 1, 10],
    ['Setup Expenses', month, 'Design', 20], ['Total'], [],
  ];
  const source = [
    ['_category', 'param', month],
    ['Construction Plan', 'clinics const. count', 1],
    ['Construction Plan', 'complex room const. count', 6],
    ['Construction Plan', 'lobby const. count', 1],
    ['Construction Expenses', 'lobby const. cost', 10],
    ['Construction Expenses', 'clinics const. cost', 10],
    ['Construction Expenses', 'complex room const. cost', 60],
    ['Totals', 'all const. cost', 80],
  ];
  const formats = [['General'], ['[=1]0 "clinic";0 "clinics"'], ['[=1]0 "complex room";0 "complex rooms"'], ['[=1]0 "lobby";0 "lobbies"'], ['General'], ['General'], ['General'], ['General']];
  assert.deepEqual(JSON.parse(JSON.stringify(ctx._detailsUnitFormats(rows, source, formats))), [
    formats[1], ['[=1]0 "complex home";0 "complex homes"'], formats[3], ['General'], ['General'], ['General'],
  ]);
  assert.throws(() => ctx._detailsUnitFormats([['Construction Expenses', month, 99, 10]], source, formats), /could not be matched/);
  assert.throws(() => ctx._detailsUnitFormats([['Construction Expenses', new Date(2030, 0, 1), 1, 10]], source, formats), /outside the source/);
});

test('period options start with All and are deduplicated from the source timeline', () => {
  const ctx = context();
  const dates = [new Date(2028, 0, 1), new Date(2028, 1, 1), new Date(2028, 3, 1)];
  assert.deepEqual(Array.from(ctx._detailsPeriodOptions(dates, 'Quarter')), ['All Quarters', 'Q1 2028', 'Q2 2028']);
  assert.deepEqual(Array.from(ctx._detailsPeriodOptions(dates, 'Year')), ['All Years', '2028']);
  assert.equal(ctx._detailsPeriodOptions(dates, 'Month')[0], 'All Month');
  assert.throws(() => ctx._detailsPeriodOptions([], 'Year'), /no months/);
});

test('Details retries transient spreadsheet service failures and releases its refresh lock', () => {
  let released = 0;
  let attempts = 0;
  let formatted = 0;
  const sheet = { getMaxRows: () => 20, getRange: () => ({ breakApart() {} }) };
  const ctx = context({
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() { released++; } }) },
    Utilities: { sleep() {} },
  });
  ctx._resizeDetailsRows = () => {
    if (++attempts === 1) throw new Error('Service Spreadsheets failed while accessing document');
  };
  ctx._refreshDetailsUnitFormats = () => { formatted++; };
  ctx._mergeDetailsTypes = () => {};
  ctx._refreshDetailsView(sheet);
  assert.equal(attempts, 2);
  assert.equal(formatted, 1);
  assert.equal(released, 1);
  ctx._resizeDetailsRows = () => { throw new Error('Content below Total'); };
  assert.throws(() => ctx._refreshDetailsView(sheet), /Content below Total/);
  assert.equal(released, 2);
});

test('type merges group only consecutive matching expense rows, excluding blanks and Total', () => {
  const ctx = context();
  const groups = types => JSON.parse(JSON.stringify(ctx._detailsTypeGroups(types)));
  assert.deepEqual(groups([
    'Construction Expenses', 'Construction Expenses', 'Land Expenses',
    'Setup Expenses', 'Setup Expenses', '', '',
    'Construction Expenses', 'Construction Expenses', 'Total', 'Total',
  ]), [{ row: 5, count: 2 }, { row: 8, count: 2 }, { row: 12, count: 2 }]);
  assert.deepEqual(groups(['Total']), []);
  assert.deepEqual(groups(['Setup Expenses']), []);
  assert.deepEqual(groups([]), []);
});

test('unit labels refresh when filters or source data change', () => {
  const ctx = context();
  const refreshed = [];
  ctx._refreshDetailsView = sheet => refreshed.push(sheet);
  ctx._updateDetailsPeriodOptions = () => {};
  const details = { getSheetId: () => 2109063306 };
  const edit = (row, column, sheet = details) => ({ range: {
    getSheet: () => sheet, getRow: () => row, getLastRow: () => row,
    getColumn: () => column, getLastColumn: () => column,
  } });
  for (const column of [1, 3, 4]) ctx.updateDetailsView(edit(2, column));
  ctx.updateDetailsView(edit(2, 2));
  ctx.updateDetailsView(edit(5, 4));
  ctx.updateDetailsView();
  assert.equal(refreshed.length, 3);
  const source = { getSheetId: () => 123, getParent: () => ({ getSheetById: () => details }) };
  ctx.updateDetailsView(edit(10, 5, source));
  assert.equal(refreshed.length, 4);
  assert.ok(refreshed.every(sheet => sheet === details));
});

test('unrelated edits do not change the period selection', () => {
  const { ctx, sheet, values } = fixture();
  ctx.updateDetailsPeriod({ range: { getSheet: () => sheet, getRow: () => 3 } });
  ctx.updateDetailsPeriod({ range: { getSheet: () => ({ getSheetId: () => 1 }) } });
  ctx.updateDetailsPeriod({ range: { getSheet: () => sheet, getRow: () => 1, getLastRow: () => 1 } });
  ctx.updateDetailsPeriod({ range: { getSheet: () => sheet, getRow: () => 2, getLastRow: () => 2, getColumn: () => 2, getLastColumn: () => 2 } });
  ctx.updateDetailsPeriod({});
  assert.deepEqual(values, { A2: 'Quarter', C2: 'Dec 2026' });
});

test('menu navigation opens the selected month and retains land and planning views', () => {
  for (const [param, expected] of [['planning expenses', 'setup expenses'],
    ['construction expenses', 'construction expenses'], ['land expenses', 'land expenses'],
    ['direct operating expenses', 'direct operating expenses'],
    ['other operating expenses', 'other operating expenses'],
    ['borrowings repaid', 'borrowing repayments'], ['borrowing repayments', 'borrowing repayments'],
    ['loan disbursed', 'loan disbursements'], ['loan disbursements', 'loan disbursements'],
    ['infra expenses', 'infrastructure expenses'], ['infrastructure expenses', 'infrastructure expenses']]) {
    const { ctx, values, formulas, validations } = fixture();
    ctx.showDetails(param, new Date(2027, 4, 1));
    assert.equal(values.A2, 'Month');
    assert.ok(formulas.C2.includes('TO_TEXT("May 2027")'));
    assert.equal(values.D2, expected);
    assert.equal(values.activated, true);
    assert.ok(formulas.C2.startsWith('=LET('));
  }
});
