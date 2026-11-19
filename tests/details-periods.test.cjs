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
  const values = { B1: 'Quarter', B2: 'Dec 2026' };
  const validations = {};
  const sheet = {
    getSheetId: () => 2109063306,
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
  const range = { getSheet: () => sheet, getRow: () => 1, getColumn: () => 2, getLastColumn: () => 2 };
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.B2, 'Q4 2026');
  assert.equal(values.A2, 'Select Quarter');
  assert.equal(validations.B2.allowInvalid, false);
  assert.equal(validations.B2.options.length, 6);
  values.B1 = 'Year';
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.B2, '2026');
  assert.equal(values.A2, 'Select Year');
  values.B1 = 'Month';
  values.B2 = 'All Years';
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.B2, 'All Month');
  assert.equal(values.A2, 'Select Month');
  assert.equal(validations.B2.options[0], 'All Month');
  values.B1 = 'Quarter';
  ctx.updateDetailsPeriod({ range });
  assert.equal(values.B2, 'All Quarters');
  assert.equal(values.A2, 'Select Quarter');
  assert.equal(validations.B2.options[0], 'All Quarters');
});

test('unrelated edits do not change the period selection', () => {
  const { ctx, sheet, values } = fixture();
  ctx.updateDetailsPeriod({ range: { getSheet: () => sheet, getRow: () => 2 } });
  ctx.updateDetailsPeriod({ range: { getSheet: () => ({ getSheetId: () => 1 }) } });
  ctx.updateDetailsPeriod({});
  assert.deepEqual(values, { B1: 'Quarter', B2: 'Dec 2026' });
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
    assert.equal(values.B1, 'Month');
    assert.equal(values.B2, 'May 2027');
    assert.equal(values.B3, expected);
    assert.equal(values.activated, true);
    assert.ok(validations.B2.options.includes('May 2027'));
  }
});
