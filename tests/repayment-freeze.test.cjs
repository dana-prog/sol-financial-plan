const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const zlib = require('node:zlib');
const path = require('node:path');

function fixture() {
  const cells = Array.from({length: 8}, () => Array.from({length: 13}, () => ({value: '', formula: ''})));
  cells[0][0].value = 'Live Loans Repayments Calculations';
  cells[2][9].value = 'Param';
  cells[2][10].value = new Date('2026-01-01T00:00:00Z');
  cells[2][11].value = new Date('2026-02-01T00:00:00Z');
  cells[2][12].value = '_name';
  ['start balance', 'principal paid', 'interest paid', 'loan draw', 'total loan paid'].forEach((label, i) => {
    cells[i + 3][9].value = label;
    cells[i + 3][0].value = i === 0 ? 'Loan A' : '';
    [10, 11].forEach(col => { cells[i + 3][col] = {value: i * 10 + col, formula: '=SUM(1,2)'}; });
  });
  const state = {cells, properties: {}, writes: [], messages: [], failWrite: 0, locked: false, busy: false};
  const props = {
    getProperty: key => state.properties[key] ?? null,
    setProperty: (key, value) => { state.properties[key] = value; },
    deleteProperty: key => { delete state.properties[key]; },
    getKeys: () => Object.keys(state.properties)
  };
  function range(row, column, height, width) {
    if (row === 'A1') { row = 1; column = 1; height = 1; width = 1; }
    const read = fn => Array.from({length: height}, (_, r) =>
      Array.from({length: width}, (_, c) => fn(cells[row - 1 + r][column - 1 + c])));
    const write = (data, formula) => {
      state.writes.push({row, column, height, width});
      if (state.writes.length === state.failWrite) throw new Error('Simulated write failure');
      data.forEach((line, r) => line.forEach((value, c) => {
        const cell = cells[row - 1 + r][column - 1 + c];
        if (formula) cell.formula = value;
        else { cell.value = value; cell.formula = ''; }
      }));
    };
    return {
      getRow: () => row,
      getColumn: () => column,
      getSheet: () => sheet,
      getValue: () => cells[row - 1][column - 1].value,
      setValue: value => {
        cells[row - 1][column - 1].value = value;
        return {setNote: note => { state.note = note; }};
      },
      setNote: note => { state.note = note; },
      getValues: () => read(cell => cell.value),
      getDisplayValues: () => read(cell => String(cell.value)),
      getFormulas: () => read(cell => cell.formula),
      setValues: data => write(data, false),
      setFormulas: data => write(data, true)
    };
  }
  const sheet = {getName: () => 'In Loans', getRange: range, getLastRow: () => 8, getLastColumn: () => 13, getSheetId: () => 42};
  const spreadsheet = {getSheetByName: () => sheet, getId: () => 'spreadsheet',
    toast(message) { state.messages.push({message, writes: state.writes.length}); }};
  const blob = (data, contentType) => ({contentType, getBytes: () => Buffer.from(data), getDataAsString: () => Buffer.from(data).toString()});
  const context = vm.createContext({
    Date,
    SpreadsheetApp: {getActiveSpreadsheet: () => spreadsheet, flush() {}},
    PropertiesService: {getDocumentProperties: () => props},
    LockService: {getDocumentLock: () => ({
      tryLock: () => { if (state.busy) return false; state.locked = true; return true; },
      releaseLock: () => { state.locked = false; }
    })},
    Utilities: {
      newBlob: blob,
      gzip: data => {
        assert.equal(data.contentType, 'application/json');
        return blob(zlib.gzipSync(data.getBytes()), 'application/gzip');
      },
      ungzip: data => {
        assert.equal(data.contentType, 'application/gzip');
        return blob(zlib.gunzipSync(data.getBytes()), 'application/json');
      },
      base64Encode: data => Buffer.from(data).toString('base64'),
      base64Decode: data => Buffer.from(data, 'base64')
    }
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/RepaymentFreeze.js'), 'utf8'), context);
  return {...state, state, context, range, sheet};
}

test('freezes only principal/interest formulas and restores every formula', () => {
  const {state, context} = fixture();
  const before = structuredClone(state.cells);
  context.freezeLoanRepayments();
  assert.deepEqual(state.writes.map(write => write.row), [5, 6]);
  assert.equal(state.cells[4][10].formula, '');
  assert.equal(state.cells[4][10].value, before[4][10].value);
  assert.equal(state.cells[6][10].formula, before[6][10].formula);
  context.unfreezeLoanRepayments();
  assert.deepEqual(state.cells, before);
  assert.deepEqual(state.properties, {});
  assert.equal(state.locked, false);
});

test('repeated freeze preserves the original backup', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  const backup = {...state.properties};
  context.freezeLoanRepayments();
  assert.deepEqual(state.properties, backup);
  context.unfreezeLoanRepayments();
  assert.equal(state.cells[4][10].formula, '=SUM(1,2)');
});

test('manual values are never replaced, including edits while frozen', () => {
  const {state, context} = fixture();
  state.cells[4][10] = {formula: '', value: 5};
  context.freezeLoanRepayments();
  state.cells[4][10].value = 99;
  context.unfreezeLoanRepayments();
  assert.deepEqual(state.cells[4][10], {formula: '', value: 99});
});

test('formula errors prevent all writes', () => {
  const {state, context} = fixture();
  state.cells[5][11].value = '#REF!';
  assert.throws(() => context.freezeLoanRepayments(), /numeric result/);
  assert.equal(state.writes.length, 0);
  assert.deepEqual(state.properties, {});
});

test('changed timeline prevents restoration and retains backup', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  state.cells[2][10].value = new Date('2027-01-01T00:00:00Z');
  assert.throws(() => context.unfreezeLoanRepayments(), /layout or timeline changed/);
  assert.equal(state.writes.length, 2);
  assert.ok(Object.keys(state.properties).length);
});

test('partial freeze can be recovered using Unfreeze', () => {
  const {state, context} = fixture();
  const before = structuredClone(state.cells);
  state.failWrite = 2;
  assert.throws(() => context.freezeLoanRepayments(), /Simulated/);
  assert.ok(Object.keys(state.properties).length);
  context.unfreezeLoanRepayments();
  assert.deepEqual(state.cells, before);
});

test('partial restoration retains backup and can be retried', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  state.failWrite = 4;
  assert.throws(() => context.unfreezeLoanRepayments(), /Simulated/);
  assert.ok(Object.keys(state.properties).length);
  context.unfreezeLoanRepayments();
  assert.deepEqual(state.properties, {});
});

test('missing backup chunk prevents all restoration writes', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  delete state.properties.inLoanRepaymentSnapshotV1_0;
  assert.throws(() => context.unfreezeLoanRepayments(), /Incomplete/);
  assert.equal(state.writes.length, 2);
});

test('no backup and lock contention cannot change cells', () => {
  const {state, context} = fixture();
  assert.throws(() => context.unfreezeLoanRepayments(), /No saved/);
  state.busy = true;
  assert.throws(() => context.freezeLoanRepayments(), /Another repayment/);
  assert.equal(state.writes.length, 0);
});

test('A1 dropdown freezes and restores formulas through the existing edit handler', () => {
  const {state, context, range} = fixture();
  context.SOLLibrary = {};
  context.updateUnits = () => { throw new Error('Unrelated handler called'); };
  context.updateTimelineParamNames = () => { throw new Error('Unrelated handler called'); };
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Main.js'), 'utf8'), context);
  const event = {range: range('A1')};
  state.cells[0][0].value = 'Frozen Loans Repayments Calculations';
  context.onInstallableEdit(event);
  assert.match(state.messages[0].message, /Freezing repayments/);
  assert.equal(state.messages[0].writes, 0);
  assert.match(state.messages[1].message, /are frozen/);
  assert.equal(state.cells[4][10].formula, '');
  assert.equal(state.cells[0][0].value, 'Frozen Loans Repayments Calculations');
  state.cells[0][0].value = 'Live Loans Repayments Calculations';
  context.onInstallableEdit(event);
  assert.match(state.messages[2].message, /Restoring repayment formulas/);
  assert.equal(state.messages[2].writes, 2);
  assert.match(state.messages[3].message, /formulas restored/);
  assert.equal(state.cells[4][10].formula, '=SUM(1,2)');
  assert.equal(state.note, '');
});

test('direct actions keep A1 synchronized', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  assert.equal(state.cells[0][0].value, 'Frozen Loans Repayments Calculations');
  context.unfreezeLoanRepayments();
  assert.equal(state.cells[0][0].value, 'Live Loans Repayments Calculations');
});

test('edits outside In Loans A1 never toggle repayments', () => {
  const {state, context, range, sheet} = fixture();
  assert.equal(context.handleRepaymentModeEdit(), false);
  assert.equal(context.handleRepaymentModeEdit({range: range(1, 2, 1, 1)}), false);
  assert.equal(context.handleRepaymentModeEdit({range: range(2, 1, 1, 1)}), false);
  sheet.getName = () => 'Cash Flow';
  assert.equal(context.handleRepaymentModeEdit({range: range('A1')}), false);
  assert.equal(state.writes.length, 0);
});

test('failed dropdown action leaves an explanatory note and keeps the backup', () => {
  const {state, context, range} = fixture();
  state.cells[0][0].value = 'Frozen Loans Repayments Calculations';
  state.failWrite = 2;
  assert.throws(() => context.handleRepaymentModeEdit({range: range('A1')}), /Simulated/);
  assert.match(state.note, /not applied successfully/);
  assert.ok(state.properties.inLoanRepaymentSnapshotV1);
  state.cells[0][0].value = 'Live Loans Repayments Calculations';
  context.handleRepaymentModeEdit({range: range('A1')});
  assert.equal(state.note, '');
  assert.equal(state.cells[4][10].formula, '=SUM(1,2)');
});

test('backup left before the first write is safely replaced on retry', () => {
  const {state, context} = fixture();
  state.failWrite = 1;
  assert.throws(() => context.freezeLoanRepayments(), /Simulated/);
  assert.ok(state.properties.inLoanRepaymentSnapshotV1);
  state.cells[4][10].value = 123;
  context.freezeLoanRepayments();
  assert.equal(state.cells[4][10].value, 123);
  assert.equal(state.cells[4][10].formula, '');
  context.unfreezeLoanRepayments();
  assert.equal(state.cells[4][10].formula, '=SUM(1,2)');
});

test('a partially frozen backup cannot be overwritten by another Freeze', () => {
  const {state, context} = fixture();
  state.failWrite = 2;
  assert.throws(() => context.freezeLoanRepayments(), /Simulated/);
  const backup = {...state.properties};
  assert.throws(() => context.freezeLoanRepayments(), /interrupted/);
  assert.deepEqual(state.properties, backup);
});

test('invalid dropdown value cannot modify repayment cells', () => {
  const {state, context, range} = fixture();
  state.cells[0][0].value = 'invalid';
  assert.throws(() => context.handleRepaymentModeEdit({range: range('A1')}), /Select Live/);
  assert.equal(state.writes.length, 0);
});
