const IN_LOANS_FREEZE_BACKUP = '_in_loans_freeze_backup';
const IN_LOANS_FREEZE_VERSION = 1;
const IN_LOANS_BACKUP_CHUNK_SIZE = 30000;
const IN_LOANS_CALCULATION_TOGGLE = 'A1';

function updateInLoansCalculationMode(event) {
  const range = event && event.range;
  if (!range || range.getRow() !== 1 || range.getColumn() !== 1) return false;
  const toggle = _getInLoansCalculationToggle(event.source);
  if (!toggle || range.getSheet().getSheetId() !== toggle.sheetId) return false;

  const message = _applyInLoansCalculationAction(event.source, spreadsheet => {
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
      throw new Error('Change the A1 dropdown on its own, outside a multi-cell paste.');
    }
    const selection = event.value;
    if (selection === 'Frozen Values' || selection === 'Freeze') return _freezeInLoans(spreadsheet);
    if (selection === 'Live Values' || selection === 'Unfreeze') return _unfreezeInLoans(spreadsheet);
    throw new Error('Choose Frozen Values or Live Values from the A1 dropdown.');
  });
  event.source.toast(message, 'In Loans calculations', 10);
  return true;
}

function _getInLoansCalculationToggle(spreadsheet) {
  const result = Sheets.Spreadsheets.get(spreadsheet.getId(), {
    ranges: ["'In Loans'!" + IN_LOANS_CALCULATION_TOGGLE],
    fields: 'sheets(properties(sheetId),data(rowData(values(userEnteredValue,dataValidation))))'
  });
  const sheet = result.sheets[0];
  const cell = sheet.data?.[0]?.rowData?.[0]?.values?.[0] || {};
  const condition = cell.dataValidation?.condition;
  if (condition?.type !== 'ONE_OF_LIST') return null;
  const choices = condition.values.map(value => value.userEnteredValue);
  const isStatus = choices.includes('Frozen Values') && choices.includes('Live Values');
  const isLegacy = choices.includes('Freeze') && choices.includes('Unfreeze');
  return choices.length === 2 && (isStatus || isLegacy)
    ? {sheetId: sheet.properties.sheetId, value: cell.userEnteredValue?.stringValue, isStatus} : null;
}

function _syncInLoansCalculationToggle(spreadsheet) {
  const toggle = _getInLoansCalculationToggle(spreadsheet);
  if (!toggle) return;
  const backup = _readInLoansBackup(spreadsheet);
  const value = backup && backup.state === 'frozen' ? 'Frozen Values' : 'Live Values';
  // Successful edits already display the correct state; avoid another recalculation.
  if (toggle.isStatus && toggle.value === value) return;
  Sheets.Spreadsheets.batchUpdate({requests: [{updateCells: {
    start: {sheetId: toggle.sheetId, rowIndex: 0, columnIndex: 0},
    rows: [{values: [{userEnteredValue: {stringValue: value}, dataValidation: {
      condition: {type: 'ONE_OF_LIST', values: [
        {userEnteredValue: 'Frozen Values'}, {userEnteredValue: 'Live Values'}]},
      strict: true, showCustomUi: true
    }}]}], fields: 'userEnteredValue,dataValidation'
  }}]}, spreadsheet.getId());
}

function freezeInLoans() {
  _runInLoansCalculationAction(_freezeInLoans);
}

function unfreezeInLoans() {
  _runInLoansCalculationAction(_unfreezeInLoans);
}

function showInLoansCalculationStatus() {
  _runInLoansCalculationAction((spreadsheet) => {
    const backup = _readInLoansBackup(spreadsheet);
    return backup && backup.state === 'frozen'
      ? 'Frozen since ' + backup.frozenAt + '. Existing input cells remain editable. ' +
        'Do not sort, insert, delete, or move rows/columns until you unfreeze.'
      : 'Live: In Loans contains its formulas.';
  });
}

function _runInLoansCalculationAction(action) {
  const message = _applyInLoansCalculationAction(SpreadsheetApp.getActive(), action);
  SpreadsheetApp.getUi().alert('In Loans calculations', message,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function _applyInLoansCalculationAction(spreadsheet, action) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(10000);
    let message;
    try {
      message = action(spreadsheet);
    } catch (error) {
      console.error(error);
      message = 'Could not complete the operation: ' + error.message;
    }
    try {
      // A status-write failure must not hide the original operation's error.
      _syncInLoansCalculationToggle(spreadsheet);
    } catch (error) {
      console.error(error);
      message += ' A1 status could not be updated: ' + error.message;
    }
    return message;
  } catch (error) {
    console.error(error);
    return 'Could not complete the operation: ' + error.message;
  } finally {
    lock.releaseLock();
  }
}

function _freezeInLoans(spreadsheet) {
  const previous = _readInLoansBackup(spreadsheet);
  if (previous && previous.state === 'frozen') {
    return 'In Loans is already frozen. The original formulas are still saved.';
  }
  const sheet = _getInLoansSheet(spreadsheet);
  const data = _readInLoansCells(spreadsheet, sheet);
  const snapshot = _buildInLoansSnapshot(data);
  const serialized = JSON.stringify(snapshot);
  const chunks = [];
  for (let offset = 0; offset < serialized.length; offset += IN_LOANS_BACKUP_CHUNK_SIZE) {
    chunks.push(serialized.slice(offset, offset + IN_LOANS_BACKUP_CHUNK_SIZE));
  }
  const backupId = previous ? previous.sheetId : _newInLoansBackupId(spreadsheet);
  const header = {
    version: IN_LOANS_FREEZE_VERSION, state: 'frozen',
    frozenAt: new Date().toISOString(),
    chunks: chunks.length, checksum: _inLoansChecksum(serialized)
  };
  const requests = [];
  if (!previous) {
    requests.push({addSheet: {properties: {
      sheetId: backupId, title: IN_LOANS_FREEZE_BACKUP, hidden: true,
      gridProperties: {rowCount: chunks.length + 1, columnCount: 1}
    }}});
  } else if (previous.rowCount < chunks.length + 1) {
    requests.push({appendDimension: {
      sheetId: backupId, dimension: 'ROWS',
      length: chunks.length + 1 - previous.rowCount
    }});
  }
  requests.push(_inLoansWrite(backupId, 0, 0,
    [JSON.stringify(header), ...chunks].map(text => [{stringValue: text}])));
  requests.push(..._inLoansCellWrites(snapshot, 'freeze'));
  // Store the backup and freeze together: a failed batch cannot leave half a freeze.
  Sheets.Spreadsheets.batchUpdate({requests}, spreadsheet.getId());
  return 'In Loans is frozen. Its formulas are safely saved as text. ' +
    'Existing inputs remain editable; other tabs use the frozen results. ' +
    'Use SOL → In Loans calculations → Unfreeze to resume. ' +
    'Do not sort or change the row/column structure while frozen.';
}

function _unfreezeInLoans(spreadsheet) {
  const backup = _readInLoansBackup(spreadsheet);
  if (!backup || backup.state !== 'frozen') return 'In Loans is already live.';
  const sheet = _getInLoansSheet(spreadsheet);
  const current = _readInLoansCells(spreadsheet, sheet, 'userEnteredValue');
  const toggle = _getInLoansCalculationToggle(spreadsheet);
  _validateInLoansRestore(backup.snapshot, current,
    toggle ? {row: 0, column: 0} : null);
  const requests = _inLoansCellWrites(backup.snapshot, 'restore');
  const header = {version: IN_LOANS_FREEZE_VERSION, state: 'live'};
  requests.push(_inLoansWrite(backup.sheetId, 0, 0, [[{stringValue: JSON.stringify(header)}]]));
  Sheets.Spreadsheets.batchUpdate({requests}, spreadsheet.getId());
  return 'In Loans formulas have been restored. Calculations can resume. ' +
    'Edits to existing input cells were preserved.';
}

function _getInLoansSheet(spreadsheet) {
  const metadata = Sheets.Spreadsheets.get(spreadsheet.getId(), {fields: 'sheets(properties)'});
  const sheet = metadata.sheets.find(sheet => sheet.properties.title === 'In Loans');
  if (!sheet) throw new Error('The In Loans tab was not found.');
  return sheet.properties;
}

function _readInLoansCells(spreadsheet, sheet, cellFields) {
  const fields = 'sheets(properties,data(rowData(values(' +
    (cellFields || 'userEnteredValue,effectiveValue') + '))))';
  const range = "'" + sheet.title.replace(/'/g, "''") + "'!A1:" +
    _a1(sheet.gridProperties.rowCount, sheet.gridProperties.columnCount);
  const result = Sheets.Spreadsheets.get(spreadsheet.getId(), {ranges: [range], fields});
  return result.sheets[0];
}

function _buildInLoansSnapshot(sheetData) {
  const grid = sheetData.properties.gridProperties;
  if (grid.rowCount * grid.columnCount > 100000) {
    throw new Error('In Loans exceeds the supported 100,000-cell snapshot size.');
  }
  const rows = sheetData.data?.[0]?.rowData || [];
  const cells = [];
  let formulaCount = 0;
  for (let row = 0; row < grid.rowCount; row++) {
    const values = (rows[row] && rows[row].values) || [];
    for (let column = 0; column < grid.columnCount; column++) {
      const cell = values[column] || {};
      const original = cell.userEnteredValue || null;
      const frozen = cell.effectiveValue || null;
      if (frozen && frozen.errorValue) {
        throw new Error(_a1(row + 1, column + 1) + ' has a calculation error or is still loading. ' +
          'Wait for valid results before freezing.');
      }
      if (original && original.formulaValue) {
        formulaCount++;
      }
      if (!original || original.formulaValue) cells.push([row, column, original, frozen]);
    }
  }
  if (!formulaCount) throw new Error('In Loans has no formulas to freeze.');
  return {sheetId: sheetData.properties.sheetId, rowCount: grid.rowCount,
    columnCount: grid.columnCount, cells, formulaCount};
}

function _validateInLoansRestore(snapshot, current, controlCell) {
  const properties = current.properties;
  if (properties.sheetId !== snapshot.sheetId ||
      properties.gridProperties.rowCount !== snapshot.rowCount ||
      properties.gridProperties.columnCount !== snapshot.columnCount) {
    throw new Error('The In Loans row/column structure changed. Restore its structure before unfreezing.');
  }
  const rows = current.data?.[0]?.rowData || [];
  for (const [row, column, original, frozen] of snapshot.cells) {
    // Permit installing the control into a previously blank cell while frozen.
    if (!original && !frozen && controlCell?.row === row && controlCell.column === column) continue;
    const values = (rows[row] && rows[row].values) || [];
    const entered = (values[column] || {}).userEnteredValue || null;
    if (!_inLoansValuesEqual(entered, frozen)) {
      throw new Error(_a1(row + 1, column + 1) + ' was changed while frozen. ' +
        'Undo that edit before unfreezing; saved formulas have been retained.');
    }
  }
}

function _inLoansValuesEqual(left, right) {
  const normalize = value => !value || value.stringValue === '' ? null : value;
  return JSON.stringify(normalize(left)) === JSON.stringify(normalize(right));
}

function _inLoansCellWrites(snapshot, action) {
  const runs = [];
  let run = null;
  for (const [row, column, original, frozen] of snapshot.cells) {
    // Blank cells need validation, but writing them is unnecessary.
    if (!original && !frozen) continue;
    const value = action === 'freeze' ? frozen : original;
    if (!run || run.row !== row || run.column + run.values.length !== column) {
      run = {row, column, values: []};
      runs.push(run);
    }
    run.values.push(value);
  }
  return runs.map(run => _inLoansWrite(snapshot.sheetId, run.row, run.column, [run.values]));
}

function _inLoansWrite(sheetId, row, column, values) {
  return {updateCells: {start: {sheetId, rowIndex: row, columnIndex: column},
    rows: values.map(rowValues => ({values: rowValues.map(value =>
      value ? {userEnteredValue: value} : {})})), fields: 'userEnteredValue'}};
}

function _readInLoansBackup(spreadsheet) {
  // Read through the same API used for writes; SpreadsheetApp caches can be stale.
  const metadata = Sheets.Spreadsheets.get(spreadsheet.getId(), {fields: 'sheets(properties)'});
  const sheet = metadata.sheets.find(sheet => sheet.properties.title === IN_LOANS_FREEZE_BACKUP);
  if (!sheet) return null;
  const result = Sheets.Spreadsheets.get(spreadsheet.getId(), {
    ranges: ["'" + IN_LOANS_FREEZE_BACKUP + "'!A1:A" + sheet.properties.gridProperties.rowCount],
    fields: 'sheets(data(rowData(values(userEnteredValue))))'
  });
  const rows = result.sheets[0].data?.[0]?.rowData || [];
  const text = rows.map(row => row.values?.[0]?.userEnteredValue?.stringValue || '');
  const header = JSON.parse(text[0]);
  if (header.version !== IN_LOANS_FREEZE_VERSION || !['live', 'frozen'].includes(header.state)) {
    throw new Error('The freeze backup is not recognized. It has not been overwritten.');
  }
  const backup = {...header, sheetId: sheet.properties.sheetId,
    rowCount: sheet.properties.gridProperties.rowCount};
  if (header.state === 'live') return backup;
  if (!Number.isInteger(header.chunks) || header.chunks < 1 || header.chunks >= text.length) {
    throw new Error('The freeze backup is incomplete.');
  }
  const serialized = text.slice(1, header.chunks + 1).join('');
  if (_inLoansChecksum(serialized) !== header.checksum) throw new Error('The freeze backup was modified.');
  return {...backup, snapshot: JSON.parse(serialized)};
}

function _inLoansChecksum(text) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text));
}

function _newInLoansBackupId(spreadsheet) {
  const metadata = Sheets.Spreadsheets.get(spreadsheet.getId(), {fields: 'sheets(properties(sheetId))'});
  const ids = new Set(metadata.sheets.map(sheet => sheet.properties.sheetId));
  let id = 1;
  while (ids.has(id)) id++;
  return id;
}
