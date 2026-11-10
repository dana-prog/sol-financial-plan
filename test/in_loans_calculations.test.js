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
  let selection = 'Live Values';
  let state = 'frozen';
  const messages = [];
  const range = {getSheet: () => ({getSheetId: () => 42}), getRow: () => 1,
    getColumn: () => 1, getNumRows: () => 1, getNumColumns: () => 1};
  const source = {toast: message => messages.push(message)};
  Object.assign(context, {
    _getInLoansCalculationToggle: () => ({sheetId: 42}),
    _freezeInLoans: () => { state = 'frozen'; return 'Frozen Values'; },
    _unfreezeInLoans: () => { state = 'live'; return 'Live'; },
    _syncInLoansCalculationToggle: () => { selection = state === 'frozen' ? 'Frozen Values' : 'Live Values'; },
    LockService: {getDocumentLock: () => ({waitLock() {}, releaseLock() {}})}
  });
  try {
    assert.equal(context.updateInLoansCalculationMode({range, source, value: 'Live Values'}), true);
    assert.equal(state, 'live');
    assert.equal(selection, 'Live Values');
    context.updateInLoansCalculationMode({range, source, value: 'Frozen Values'});
    assert.equal(state, 'frozen');
    assert.deepEqual(messages, ['Live', 'Frozen Values']);
    assert.equal(context.updateInLoansCalculationMode({range: {...range, getColumn: () => 2}, source}), false);
    assert.equal(context.updateInLoansCalculationMode({range: {...range,
      getSheet: () => ({getSheetId: () => 99})}, source}), false);
  } finally { Object.assign(context, saved); }
});

test('failed freeze and multi-cell paste restore dropdown to actual state', () => {
  const saved = {...context};
  let selection = 'Live Values';
  let message;
  let released = 0;
  const range = {getSheet: () => ({getSheetId: () => 42}), getRow: () => 1,
    getColumn: () => 1, getNumRows: () => 1, getNumColumns: () => 1};
  const source = {toast: text => {message = text;}};
  Object.assign(context, {
    console: {error() {}},
    _getInLoansCalculationToggle: () => ({sheetId: 42}),
    _freezeInLoans: () => { throw new Error('Loading result'); },
    _syncInLoansCalculationToggle: () => { selection = 'Live Values'; },
    LockService: {getDocumentLock: () => ({waitLock() {}, releaseLock() {released++;}})}
  });
  try {
    context.updateInLoansCalculationMode({range, source, value: 'Frozen Values'});
    assert.equal(selection, 'Live Values');
    assert.match(message, /Loading result/);
    context.updateInLoansCalculationMode({range: {...range, getNumColumns: () => 2}, source});
    assert.match(message, /multi-cell paste/);
    assert.equal(released, 2);
  } finally { Object.assign(context, saved); }
});

test('dropdown displays saved status and avoids rewriting an already correct status', () => {
  const saved = {...context};
  let backup = null;
  let selection = 'Frozen Values';
  let writes = 0;
  Object.assign(context, {
    _getInLoansCalculationToggle: () => ({sheetId: 42, value: selection, isStatus: true}),
    _readInLoansBackup: () => backup,
    Sheets: {Spreadsheets: {batchUpdate: (body, id) => {
      assert.equal(id, 'test-id');
      const update = body.requests[0].updateCells;
      assert.equal(update.start.sheetId, 42);
      assert.equal(update.fields, 'userEnteredValue,dataValidation');
      selection = update.rows[0].values[0].userEnteredValue.stringValue;
      writes++;
    }}}
  });
  const spreadsheet = {getId: () => 'test-id'};
  try {
    context._syncInLoansCalculationToggle(spreadsheet);
    assert.equal(selection, 'Live Values');
    assert.equal(writes, 1);
    context._syncInLoansCalculationToggle(spreadsheet);
    assert.equal(writes, 1);
    backup = {state: 'frozen'};
    context._syncInLoansCalculationToggle(spreadsheet);
    assert.equal(selection, 'Frozen Values');
    assert.equal(writes, 2);
    context._syncInLoansCalculationToggle(spreadsheet);
    assert.equal(writes, 2);
    backup = {state: 'live'};
    context._syncInLoansCalculationToggle(spreadsheet);
    assert.equal(selection, 'Live Values');
  } finally { Object.assign(context, saved); }
});

test('control metadata reads do not require calculation-dependent SpreadsheetApp lookups', () => {
  const saved = {...context};
  const spreadsheet = {getId: () => 'test-id'};
  const properties = {sheetId: 42, title: 'In Loans',
    gridProperties: {rowCount: 3, columnCount: 4}};
  let choices = ['Frozen Values', 'Live Values'];
  Object.assign(context, {Sheets: {Spreadsheets: {get: (id, options) => {
    assert.equal(id, 'test-id');
    assert.ok(!options.fields.includes('effectiveValue'));
    return {sheets: [{properties, data: [{rowData: [{values: [{
      userEnteredValue: string('Live Values'),
      dataValidation: {condition: {type: 'ONE_OF_LIST',
        values: choices.map(userEnteredValue => ({userEnteredValue}))}}
    }]}]}]}]};
  }}}});
  try {
    assert.deepEqual(plain(context._getInLoansCalculationToggle(spreadsheet)),
      {sheetId: 42, value: 'Live Values', isStatus: true});
    assert.deepEqual(plain(context._getInLoansSheet(spreadsheet)), properties);
    assert.equal(context._newInLoansBackupId(spreadsheet), 1);
    const data = context._readInLoansCells(spreadsheet, properties, 'userEnteredValue');
    assert.equal(data.properties.sheetId, 42);
    choices = ['Freeze', 'Unfreeze'];
    assert.equal(context._getInLoansCalculationToggle(spreadsheet).isStatus, false);
    choices = ['Unrelated', 'Dropdown'];
    assert.equal(context._getInLoansCalculationToggle(spreadsheet), null);
  } finally { Object.assign(context, saved); }
});
