const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const zlib = require('node:zlib');
const path = require('node:path');

function fixture(loanCount = 1, monthCount = 2) {
  const columnCount = 11 + monthCount;
  const cells = Array.from({length: 3 + 5 * loanCount}, () => Array.from({length: columnCount}, () => ({value: '', formula: ''})));
  cells[0][0].value = 'Repayment Calculations Enabled';
  cells[2][9].value = 'Param';
  const monthColumns = Array.from({length: monthCount}, (_, month) => 10 + month);
  monthColumns.forEach((column, month) => { cells[2][column].value = new Date(Date.UTC(2026, month, 1)); });
  cells[2][10 + monthCount].value = '_name';
  const labels = Array.from({length: loanCount}, () =>
    ['start balance', 'principal paid', 'interest paid', 'loan draw', 'total loan paid']).flat();
  labels.forEach((label, i) => {
    cells[i + 3][9].value = label;
    cells[i + 3][0].value = i === 0 ? 'Loan A' : '';
    monthColumns.forEach(col => { cells[i + 3][col] = {value: i * 10 + col, formula: '=SUM(1,2)'}; });
  });
  const state = {cells, properties: {}, writes: [], messages: [], dialogs: [], failWrite: 0, locked: false, busy: false,
    reads: {formulas: 0, values: 0, display: 0}, flushes: 0, corruptWrite: false, corruptFormula: false};
  const props = {
    getProperties: () => ({...state.properties}),
    setProperties: entries => { Object.assign(state.properties, entries); },
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
        if (formula) cell.formula = state.corruptFormula ? '=0' : value;
        else { cell.value = state.corruptWrite ? value + 1 : value; cell.formula = ''; }
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
      getValues: () => { state.reads.values++; return read(cell => cell.value); },
      getDisplayValues: () => { state.reads.display++; return read(cell => String(cell.value)); },
      getFormulas: () => { state.reads.formulas++; return read(cell => cell.formula); },
      setValues: data => write(data, false),
      setFormulas: data => write(data, true)
    };
  }
  const sheet = {getName: () => 'In Loans', getRange: range, getLastRow: () => cells.length, getLastColumn: () => columnCount, getSheetId: () => 42};
  const spreadsheet = {getSheetByName: () => sheet, getId: () => 'spreadsheet',
    toast(message, title, timeout) { state.messages.push({message, title, timeout, writes: state.writes.length}); }};
  const blob = (data, contentType) => ({contentType, getBytes: () => Buffer.from(data), getDataAsString: () => Buffer.from(data).toString()});
  const ui = {
    ButtonSet: {OK: 'OK'},
    alert(title, message, buttons) {
      assert.equal(state.locked, false, 'Dialogs must not suspend an operation holding the lock');
      state.dialogs.push({title, message, buttons, writes: state.writes.length, flushes: state.flushes});
    }
  };
  const context = vm.createContext({
    Date,
    SpreadsheetApp: {getActiveSpreadsheet: () => spreadsheet, getUi: () => ui, flush() { state.flushes++; }},
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
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/LoanModes.js'), 'utf8'), context);
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
  state.cells[0][0].value = 'Repayment Calculations Disabled';
  context.onInstallableEdit(event);
  assert.match(state.messages[0].message, /Freezing repayments/);
  assert.equal(state.messages[0].writes, 0);
  assert.match(state.dialogs[0].message, /are frozen/);
  assert.equal(state.dialogs[0].buttons, 'OK');
  assert.equal(state.messages[0].timeout, -1);
  assert.equal(state.messages[1].message, '');
  assert.equal(state.cells[4][10].formula, '');
  assert.equal(state.cells[0][0].value, 'Repayment Calculations Disabled');
  state.cells[0][0].value = 'Repayment Calculations Enabled';
  context.onInstallableEdit(event);
  assert.match(state.messages[2].message, /Restoring repayment formulas/);
  assert.equal(state.messages[2].writes, 2);
  assert.match(state.dialogs[1].message, /formulas restored/);
  assert.equal(state.messages[3].message, '');
  assert.equal(state.cells[4][10].formula, '=SUM(1,2)');
  assert.equal(state.note, '');
});

test('direct actions keep A1 synchronized', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  assert.equal(state.cells[0][0].value, 'Repayment Calculations Disabled');
  context.unfreezeLoanRepayments();
  assert.equal(state.cells[0][0].value, 'Repayment Calculations Enabled');
});

test('edits outside In Loans A1 never toggle repayments', () => {
  const {state, context, range, sheet} = fixture();
  assert.equal(context.handleLoanModeEdit(), false);
  assert.equal(context.handleLoanModeEdit({range: range(1, 2, 1, 1)}), false);
  assert.equal(context.handleLoanModeEdit({range: range(2, 1, 1, 1)}), false);
  sheet.getName = () => 'Cash Flow';
  assert.equal(context.handleLoanModeEdit({range: range('A1')}), false);
  assert.equal(state.writes.length, 0);
});

test('failed dropdown action leaves an explanatory note and keeps the backup', () => {
  const {state, context, range} = fixture();
  state.cells[0][0].value = 'Repayment Calculations Disabled';
  state.failWrite = 2;
  assert.throws(() => context.handleLoanModeEdit({range: range('A1')}), /Simulated/);
  assert.match(state.note, /not applied successfully/);
  assert.equal(state.dialogs.length, 1);
  assert.match(state.dialogs[0].title, /Failed/);
  assert.match(state.dialogs[0].message, /Simulated write failure/);
  assert.ok(state.properties.inLoanRepaymentSnapshotV1);
  state.cells[0][0].value = 'Repayment Calculations Enabled';
  context.handleLoanModeEdit({range: range('A1')});
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
  assert.throws(() => context.handleLoanModeEdit({range: range('A1')}), /Select Repayment Calculations Enabled/);
  assert.equal(state.writes.length, 0);
});

test('seven loans over 96 months use constant bulk reads and preserve manual gaps', () => {
  const {state, context} = fixture(7, 96);
  state.cells.forEach((row, r) => row.forEach((cell, c) => {
    if (cell.formula) cell.formula = '=' + r + '+' + c;
  }));
  state.cells[4][35] = {formula: '', value: 7};
  state.cells[9][105] = {formula: '', value: 9};
  const before = structuredClone(state.cells);
  context.freezeLoanRepayments();
  assert.equal(state.reads.formulas + state.reads.values + state.reads.display, 6);
  assert.equal(state.flushes, 1);
  state.cells.slice(3).forEach((row, index) => {
    const repaymentRow = ['principal paid', 'interest paid'].includes(row[9].value);
    row.slice(10, 106).forEach((cell, column) => {
      assert.equal(cell.formula, repaymentRow ? '' : before[index + 3][column + 10].formula);
    });
  });
  state.reads = {formulas: 0, values: 0, display: 0};
  state.flushes = 0;
  context.unfreezeLoanRepayments();
  assert.equal(state.reads.formulas + state.reads.values + state.reads.display, 3);
  assert.equal(state.flushes, 1);
  assert.deepEqual(state.cells, before);
});

test('bulk verification detects incorrect frozen values and retains recovery backup', () => {
  const {state, context} = fixture();
  state.corruptWrite = true;
  assert.throws(() => context.freezeLoanRepayments(), /Freeze verification failed/);
  assert.ok(state.properties.inLoanRepaymentSnapshotV1);
});

test('bulk verification detects incorrect restored formulas and retains recovery backup', () => {
  const {state, context} = fixture();
  context.freezeLoanRepayments();
  state.corruptFormula = true;
  assert.throws(() => context.unfreezeLoanRepayments(), /Formula restoration failed/);
  assert.ok(state.properties.inLoanRepaymentSnapshotV1);
  state.corruptFormula = false;
  context.unfreezeLoanRepayments();
  assert.equal(state.cells[4][10].formula, '=SUM(1,2)');
});

for (const mode of ['Enabled', 'Disabled']) {
  test('D1 ' + mode + ' shows progress and a completion dialog without changing formulas', () => {
    const {state, context, range} = fixture();
    state.cells[0][3].value = 'Planned Loans ' + mode;
    const before = structuredClone(state.cells);
    context.SOLLibrary = {};
    context.updateUnits = () => { throw new Error('Unrelated handler called'); };
    context.updateTimelineParamNames = () => { throw new Error('Unrelated handler called'); };
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Main.js'), 'utf8'), context);
    context.onInstallableEdit({range: range(1, 4)});
    assert.match(state.messages[0].message, /Applying Planned Loans/);
    assert.equal(state.messages[0].timeout, -1);
    assert.equal(state.messages[1].message, '');
    assert.equal(state.dialogs.length, 1);
    assert.match(state.dialogs[0].title, /Done/);
    assert.match(state.dialogs[0].message, mode === 'Disabled' ? /draws are set to zero/ : /original calculations/);
    assert.equal(state.dialogs[0].flushes, 1);
    assert.deepEqual(state.cells, before);
    assert.deepEqual(state.properties, {});
    assert.equal(state.writes.length, 0);
  });
}

test('D1 failure reports an error dialog without reporting success', () => {
  const {state, context, range} = fixture();
  state.cells[0][3].value = 'Planned Loans Disabled';
  context.SpreadsheetApp.flush = () => { throw new Error('Recalculation failed'); };
  assert.throws(() => context.handleLoanModeEdit({range: range(1, 4)}), /Recalculation failed/);
  assert.equal(state.dialogs.length, 1);
  assert.match(state.dialogs[0].title, /Failed/);
  assert.match(state.note, /Recalculation failed/);
  assert.equal(state.messages.at(-1).message, '');
});

test('a repayment label in D1 is rejected without touching the repayment backup', () => {
  const {state, context, range} = fixture();
  state.cells[0][3].value = 'Repayment Calculations Disabled';
  assert.throws(() => context.handleLoanModeEdit({range: range(1, 4)}), /Select Planned Loans/);
  assert.equal(state.writes.length, 0);
  assert.deepEqual(state.properties, {});
});
