const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');

function context(globals = {}) {
  const ctx = vm.createContext({ Date, ...globals });
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/Details.js`, 'utf8'), ctx);
  return ctx;
}

const months = Array.from({ length: 14 }, (_, index) => new Date(2026, index + 10, 1));

test('period options follow timeline bounds and deduplicate quarters and years', () => {
  const ctx = context();
  const selection = (period, value) => JSON.parse(JSON.stringify(
    ctx._detailsPeriodSelection([...months, '', null], period, value)));
  assert.equal(selection('Month', 'Dec 2026').options.length, 15);
  assert.equal(selection('Month', 'Dec 2026').options[0], 'All Month');
  assert.equal(selection('Month', 'All Years').value, 'All Month');
  assert.deepEqual(selection('Quarter', 'Dec 2026'), {
    options: ['All Quarters', 'Q4 2026', 'Q1 2027', 'Q2 2027', 'Q3 2027', 'Q4 2027'], value: 'Q4 2026',
  });
  assert.deepEqual(selection('Year', 'Q4 2027'), { options: ['All Years', '2026', '2027'], value: '2027' });
  assert.equal(selection('Month', 'Q4 2026').value, 'Nov 2026');
  assert.equal(selection('Month', '2027').value, 'Jan 2027');
  assert.equal(selection('Quarter', 'stale selection').value, 'Q4 2026');
  assert.throws(() => ctx._detailsPeriodSelection([], 'Month', ''), /no months/);
  assert.throws(() => ctx._detailsPeriodSelection(months, 'Week', ''), /Unsupported/);
});

function fixture() {
  const values = { A2: 'Quarter', D2: 'Dec 2026' };
  const validations = {};
  const sheet = {
    getSheetId: () => 2109063306,
    getMaxRows: () => 4,
    getParent: () => ({ getRangeByName: () => ({ getValues: () => [months] }) }),
    getRange: cell => ({
      getValue: () => values[cell],
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
      newDataValidation: () => ({
        requireValueInList(options) { this.options = options; return this; },
        setAllowInvalid(value) { this.allowInvalid = value; return this; },
        build() { return this; },
      }),
    },
  });
  ctx._refreshDetailsView = () => {};
  return { ctx, sheet, values, validations };
}

test('editing period replaces dependent options and keeps the matching period', () => {
  const { ctx, sheet, values, validations } = fixture();
  const range = { getSheet: () => sheet, getRow: () => 2, getLastRow: () => 2, getColumn: () => 1 };
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.D2, 'Q4 2026');
  assert.equal(values.D1, 'Select Quarter');
  assert.equal(validations.D2.allowInvalid, false);
  assert.equal(validations.D2.options.length, 6);
  values.A2 = 'Year';
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.D2, '2026');
  assert.equal(values.D1, 'Select Year');
  values.A2 = 'Month';
  values.D2 = 'All Years';
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.D2, 'All Month');
  assert.equal(values.D1, 'Select Month');
  assert.equal(validations.D2.options[0], 'All Month');
  values.A2 = 'Quarter';
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.D2, 'All Quarters');
  assert.equal(values.D1, 'Select Quarter');
  assert.equal(validations.D2.options[0], 'All Quarters');
});

function resizeFixture({ currentRows = 100, count = 20, total = 'Total', blankTail = true } = {}) {
  const changes = [];
  let counterFormula = '';
  const sheet = {
    getMaxRows: () => currentRows,
    getMaxColumns: () => 8,
    getRowHeight: () => 24,
    getConditionalFormatRules: () => [],
    setConditionalFormatRules: () => {},
    setRowHeights: (...args) => changes.push(['height', ...args]),
    insertRowsAfter: (after, added) => { changes.push(['insert', after, added]); currentRows += added; },
    deleteRows: (...args) => changes.push(['delete', ...args]),
    getRange: (row, column, rows, columns) => {
      if (row === 'A5') return { getFormula: () => '=EXPENSE_DETAILS(A2,D2,F2)' };
      if (row === 'H1') return {
        getFormula: () => counterFormula,
        setFormula: formula => { counterFormula = formula; },
        getValue: () => typeof count === 'function' ? count() : count,
      };
      if (rows === undefined) return { getValue: () => total };
      if (row === 5 && rows === 1 && columns === 8) return {
        copyFormatToRange: (target, ...args) => changes.push(['format', ...args]),
      };
      return { isBlank: () => blankTail };
    },
  };
  const ctx = context({ SpreadsheetApp: { flush() {} }, Utilities: { sleep() {} } });
  return { ctx, sheet, changes, countFormula: () => counterFormula };
}

test('Details removes only empty rows after Total and keeps the header and total-only result', () => {
  for (const count of [20, 1]) {
    const { ctx, sheet, changes, countFormula } = resizeFixture({ count });
    ctx._resizeDetailsRows(sheet);
    assert.deepEqual(changes, [['delete', count + 5, 100 - count - 4]]);
    assert.equal(countFormula(), '=ROWS(EXPENSE_DETAILS(A2,D2,F2))');
  }
});

test('Details grows a trimmed sheet and restores row formatting', () => {
  const { ctx, sheet, changes } = resizeFixture({ currentRows: 5, count: 200 });
  ctx._resizeDetailsRows(sheet);
  assert.deepEqual(changes, [
    ['insert', 5, 199], ['format', 1, 8, 6, 204], ['height', 6, 199, 24],
  ]);
});

test('Details leaves exact-sized tables alone and refuses to delete content or incomplete results', () => {
  const exact = resizeFixture({ currentRows: 24 });
  exact.ctx._resizeDetailsRows(exact.sheet);
  assert.deepEqual(exact.changes, []);
  for (const options of [{ count: '#REF!' }, { count: 0 }, { total: '#REF!' }, { blankTail: false }]) {
    const { ctx, sheet, changes } = resizeFixture(options);
    assert.throws(() => ctx._resizeDetailsRows(sheet), /Details/);
    assert.deepEqual(changes, []);
  }
});

test('Details rechecks the result after a temporary row count changes during recalculation', () => {
  const counts = [300, 20, 20, 20];
  const { ctx, sheet, changes } = resizeFixture({ count: () => counts.shift() ?? 20 });
  ctx._resizeDetailsRows(sheet);
  assert.deepEqual(changes, [
    ['insert', 100, 204], ['format', 1, 8, 101, 304], ['height', 101, 204, 24],
    ['delete', 25, 280],
  ]);
});

test('construction details reuse numeric unit formats and reset other detail rows', () => {
  const ctx = context();
  const rows = [
    ['Construction Expenses', '', '', 1, '', '', 221850, 'clinics const. cost'],
    ['Construction Expenses', '', '', 6, '', '', 5246000, 'complex room const. cost'],
    ['Construction Expenses', '', '', 1, '', '', 1296000, 'lobby const. cost'],
    ['Setup Expenses', '', '', 'Design'],
    ['Total'],
    [],
  ];
  const params = [['clinics const. count'], ['complex room const. count'], ['lobby const. count']];
  const formats = [['[=1]0 "clinic";0 "clinics"'], ['[=1]0 "complex room";0 "complex rooms"'], ['[=1]0 "lobby";0 "lobbies"']];
  assert.deepEqual(JSON.parse(JSON.stringify(ctx._detailsUnitFormats(rows, params, formats))), [
    formats[0], ['[=1]0 "complex home";0 "complex homes"'], formats[2], ['General'], ['General'], ['General'],
  ]);
  assert.throws(() => ctx._detailsUnitFormats([rows[0]], [], []), /Missing planned-unit/);
});

test('Details retries transient spreadsheet service failures and releases its refresh lock', () => {
  let released = 0;
  let attempts = 0;
  let formatted = 0;
  const ctx = context({
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() { released++; } }) },
    Utilities: { sleep() {} },
  });
  ctx._resizeDetailsRows = () => {
    if (++attempts === 1) throw new Error('Service Spreadsheets failed while accessing document');
  };
  ctx._refreshDetailsUnitFormats = () => { formatted++; };
  ctx._refreshDetailsView({});
  assert.equal(attempts, 2);
  assert.equal(formatted, 1);
  assert.equal(released, 1);
  ctx._resizeDetailsRows = () => { throw new Error('Content below Total'); };
  assert.throws(() => ctx._refreshDetailsView({}), /Content below Total/);
  assert.equal(released, 2);
});

test('unit labels refresh when filters or source data change', () => {
  const ctx = context();
  const refreshed = [];
  ctx._refreshDetailsView = sheet => refreshed.push(sheet);
  const details = { getSheetId: () => 2109063306 };
  const edit = (row, column, sheet = details) => ({ range: {
    getSheet: () => sheet, getRow: () => row, getLastRow: () => row,
    getColumn: () => column, getLastColumn: () => column,
  } });
  for (const column of [1, 4, 6]) ctx.updateDetailsView(edit(2, column));
  ctx.updateDetailsView(edit(2, 3));
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
  ctx.updateDetailsPeriod({ range: { getSheet: () => sheet, getRow: () => 2, getLastRow: () => 2, getColumn: () => 2 } });
  ctx.updateDetailsPeriod({});
  assert.deepEqual(values, { A2: 'Quarter', D2: 'Dec 2026' });
});

test('menu navigation opens the selected month and retains land and planning views', () => {
  for (const [param, expected] of [['planning expenses', 'setup expenses'],
    ['construction expenses', 'construction expenses'], ['land expenses', 'land expenses'],
    ['direct operating expenses', 'direct operating expenses'],
    ['other operating expenses', 'other operating expenses'],
    ['borrowings repaid', 'borrowing repayments'], ['borrowing repayments', 'borrowing repayments'],
    ['loan disbursed', 'loan disbursements'], ['loan disbursements', 'loan disbursements'],
    ['infra expenses', 'infrastructure expenses'], ['infrastructure expenses', 'infrastructure expenses']]) {
    const { ctx, values, validations } = fixture();
    ctx.showDetails(param, new Date(2027, 4, 1));
    assert.equal(values.A2, 'Month');
    assert.equal(values.D2, 'May 2027');
    assert.equal(values.F2, expected);
    assert.equal(values.activated, true);
    assert.ok(validations.D2.options.includes('May 2027'));
  }
});
