const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const context = vm.createContext({console});
vm.runInContext(fs.readFileSync('src/perf/PerfExport.js', 'utf8'), context);
vm.runInContext(fs.readFileSync('src/InLoansCalculations.js', 'utf8'), context);
const plain = value => JSON.parse(JSON.stringify(value));
const number = numberValue => ({numberValue});
const string = stringValue => ({stringValue});
const formula = formulaValue => ({formulaValue});

function grid() {
  return {properties: {sheetId: 42, gridProperties: {rowCount: 3, columnCount: 4}},
    data: [{rowData: [
      {values: [
        {userEnteredValue: number(5), effectiveValue: number(5)},
        {userEnteredValue: formula('=SEQUENCE(2,2)'), effectiveValue: number(1)},
        {effectiveValue: number(2)}, {}]},
      {values: [{}, {effectiveValue: number(3)}, {effectiveValue: number(4)}, {}]},
      {values: [{userEnteredValue: formula('="=literal"'), effectiveValue: string('=literal')},
        {userEnteredValue: formula('=""'), effectiveValue: string('')},
        {userEnteredValue: formula('=FALSE'), effectiveValue: {boolValue: false}},
        {userEnteredValue: number(0), effectiveValue: number(0)}]}
    ]}]};
}

function apply(sheet, requests) {
  for (const {updateCells} of requests) {
    const {rowIndex, columnIndex} = updateCells.start;
    updateCells.rows.forEach((row, r) => row.values.forEach((cell, c) => {
      sheet.data[0].rowData[rowIndex + r].values[columnIndex + c] = plain(cell);
    }));
  }
}

test('freeze and unfreeze arrays, blanks, booleans and formula-like strings; retain input edits', () => {
  const live = grid();
  const snapshot = context._buildInLoansSnapshot(live);
  const frozen = structuredClone(live);
  apply(frozen, context._inLoansCellWrites(snapshot, 'freeze'));
  assert.deepEqual(frozen.data[0].rowData[0].values[2].userEnteredValue, number(2));
  assert.deepEqual(frozen.data[0].rowData[2].values[0].userEnteredValue, string('=literal'));
  frozen.data[0].rowData[0].values[0].userEnteredValue = number(99);
  context._validateInLoansRestore(snapshot, frozen);
  apply(frozen, context._inLoansCellWrites(snapshot, 'restore'));
  assert.deepEqual(frozen.data[0].rowData[0].values[0].userEnteredValue, number(99));
  assert.deepEqual(frozen.data[0].rowData[0].values[1].userEnteredValue, formula('=SEQUENCE(2,2)'));
  assert.deepEqual(frozen.data[0].rowData[1].values[2], {});
  assert.deepEqual(frozen.data[0].rowData[2].values[3].userEnteredValue, number(0));
});

test('refuse freezing errors and sheets without formulas; allow omitted empty results', () => {
  const data = grid();
  data.data[0].rowData[0].values[1].effectiveValue = {errorValue: {type: 'LOADING'}};
  assert.throws(() => context._buildInLoansSnapshot(data), /still loading/);
  delete data.data[0].rowData[0].values[1].effectiveValue;
  assert.equal(context._buildInLoansSnapshot(data).formulaCount, 4);
  data.data = [];
  assert.throws(() => context._buildInLoansSnapshot(data), /no formulas/);
});

test('refuse changed frozen results, new values in blank cells, and structural changes', () => {
  const data = grid();
  const snapshot = context._buildInLoansSnapshot(data);
  apply(data, context._inLoansCellWrites(snapshot, 'freeze'));
  data.data[0].rowData[0].values[2].userEnteredValue = number(999);
  assert.throws(() => context._validateInLoansRestore(snapshot, data), /C1 was changed/);
  data.data[0].rowData[0].values[2].userEnteredValue = number(2);
  data.data[0].rowData[0].values[3] = {userEnteredValue: string('new input')};
  assert.throws(() => context._validateInLoansRestore(snapshot, data), /D1 was changed/);
  data.properties.gridProperties.rowCount++;
  assert.throws(() => context._validateInLoansRestore(snapshot, data), /structure changed/);
});

test('backup content is written as literal text with the narrow field mask', () => {
  const request = plain(context._inLoansWrite(1, 0, 0, [[string('=text')]]));
  assert.equal(request.updateCells.fields, 'userEnteredValue');
  assert.deepEqual(request.updateCells.rows[0].values[0], {userEnteredValue: string('=text')});
});

test('repeated freeze never reads current formulas or replaces the saved backup', () => {
  const readBackup = context._readInLoansBackup;
  context._readInLoansBackup = () => ({state: 'frozen'});
  try {
    assert.match(context._freezeInLoans({}), /already frozen/);
  } finally {
    context._readInLoansBackup = readBackup;
  }
});

test('dropdown routes only In Loans A1 and preserves queued edit choices', () => {
  const saved = {...context};
  let selection = 'Freeze';
  let state = 'frozen';
  const messages = [];
  const range = {getSheet: () => ({getName: () => 'In Loans'}), getRow: () => 1,
    getColumn: () => 1, getNumRows: () => 1, getNumColumns: () => 1};
  const source = {toast: message => messages.push(message)};
  Object.assign(context, {
    _getInLoansCalculationToggle: () => range,
    _freezeInLoans: () => { state = 'frozen'; return 'Frozen'; },
    _unfreezeInLoans: () => { state = 'live'; return 'Live'; },
    _syncInLoansCalculationToggle: () => { selection = state === 'frozen' ? 'Unfreeze' : 'Freeze'; },
    LockService: {getDocumentLock: () => ({waitLock() {}, releaseLock() {}})}
  });
  try {
    assert.equal(context.updateInLoansCalculationMode({range, source, value: 'Unfreeze'}), true);
    assert.equal(state, 'live');
    assert.equal(selection, 'Freeze');
    context.updateInLoansCalculationMode({range, source, value: 'Freeze'});
    assert.equal(state, 'frozen');
    assert.deepEqual(messages, ['Live', 'Frozen']);
    assert.equal(context.updateInLoansCalculationMode({range: {...range, getColumn: () => 2}, source}), false);
    assert.equal(context.updateInLoansCalculationMode({range: {...range,
      getSheet: () => ({getName: () => 'Cash Flow'})}, source}), false);
  } finally { Object.assign(context, saved); }
});

test('failed freeze and multi-cell paste restore dropdown to actual state', () => {
  const saved = {...context};
  let selection = 'Freeze';
  let message;
  let released = 0;
  const range = {getSheet: () => ({getName: () => 'In Loans'}), getRow: () => 1,
    getColumn: () => 1, getNumRows: () => 1, getNumColumns: () => 1};
  const source = {toast: text => {message = text;}};
  Object.assign(context, {
    console: {error() {}},
    _getInLoansCalculationToggle: () => range,
    _freezeInLoans: () => { throw new Error('Loading result'); },
    _syncInLoansCalculationToggle: () => { selection = 'Freeze'; },
    LockService: {getDocumentLock: () => ({waitLock() {}, releaseLock() {released++;}})}
  });
  try {
    context.updateInLoansCalculationMode({range, source, value: 'Freeze'});
    assert.equal(selection, 'Freeze');
    assert.match(message, /Loading result/);
    context.updateInLoansCalculationMode({range: {...range, getNumColumns: () => 2}, source});
    assert.match(message, /multi-cell paste/);
    assert.equal(released, 2);
  } finally { Object.assign(context, saved); }
});

test('dropdown offers Freeze while live and Unfreeze while frozen', () => {
  const saved = {...context};
  let backup = null;
  let selection;
  Object.assign(context, {
    _getInLoansSheet: () => ({}),
    _getInLoansCalculationToggle: () => ({setValue: value => {selection = value;}}),
    _readInLoansBackup: () => backup,
    SpreadsheetApp: {flush() {}}
  });
  try {
    context._syncInLoansCalculationToggle({});
    assert.equal(selection, 'Freeze');
    backup = {state: 'frozen'};
    context._syncInLoansCalculationToggle({});
    assert.equal(selection, 'Unfreeze');
    backup = {state: 'live'};
    context._syncInLoansCalculationToggle({});
    assert.equal(selection, 'Freeze');
  } finally { Object.assign(context, saved); }
});
