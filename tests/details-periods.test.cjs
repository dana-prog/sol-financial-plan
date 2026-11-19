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

test('unit labels refresh when filters or source data change', () => {
  const ctx = context();
  const refreshed = [];
  ctx._refreshDetailsUnitFormats = sheet => refreshed.push(sheet);
  const details = { getSheetId: () => 2109063306 };
  const edit = (row, column, sheet = details) => ({ range: {
    getSheet: () => sheet, getRow: () => row, getLastRow: () => row,
    getColumn: () => column, getLastColumn: () => column,
  } });
  for (const column of [1, 4, 6]) ctx.updateDetailsUnitFormats(edit(2, column));
  ctx.updateDetailsUnitFormats(edit(2, 3));
  ctx.updateDetailsUnitFormats(edit(5, 4));
  ctx.updateDetailsUnitFormats();
  assert.equal(refreshed.length, 3);
  const source = { getSheetId: () => 123, getParent: () => ({ getSheetById: () => details }) };
  ctx.updateDetailsUnitFormats(edit(10, 5, source));
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
