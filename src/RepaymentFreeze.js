const REPAYMENT_FREEZE = Object.freeze({
  sheetName: 'In Loans',
  toggleCell: 'A1',
  liveValue: 'Live Loans Repayments Calculations',
  frozenValue: 'Frozen Loans Repayments Calculations',
  headerRow: 3,
  paramColumn: 10,
  firstMonthColumn: 11,
  backupKey: 'inLoanRepaymentSnapshotV1',
  chunkSize: 8000,
  params: ['principal paid', 'interest paid']
});

function handleRepaymentModeEdit(e) {
  if (!e || !e.range || e.range.getSheet().getName() !== REPAYMENT_FREEZE.sheetName ||
      e.range.getRow() !== 1 || e.range.getColumn() !== 1) return false;
  const cell = e.range.getSheet().getRange(REPAYMENT_FREEZE.toggleCell);
  try {
    const mode = cell.getValue();
    if (mode !== REPAYMENT_FREEZE.frozenValue && mode !== REPAYMENT_FREEZE.liveValue) {
      throw new Error('Select Live Loans Repayments Calculations or Frozen Loans Repayments Calculations.');
    }
    const message = mode === REPAYMENT_FREEZE.frozenValue
      ? 'Freezing repayments. Please wait for confirmation.'
      : 'Restoring repayment formulas. Please wait for confirmation.';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Loan repayments', 30);
    if (mode === REPAYMENT_FREEZE.frozenValue) freezeLoanRepayments();
    else unfreezeLoanRepayments();
  } catch (error) {
    cell.setNote('The selected mode was not applied successfully. ' + error.message);
    throw error;
  }
  return true;
}

function setRepaymentMode_(sheet, mode) {
  sheet.getRange(REPAYMENT_FREEZE.toggleCell).setValue(mode).setNote('');
}

function freezeLoanRepayments() {
  return withRepaymentLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    SpreadsheetApp.flush();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRepaymentSheet_(spreadsheet);
    const layout = getRepaymentLayout_(sheet);
    const previous = loadRepaymentSnapshot_(props);
    if (previous) {
      validateRepaymentSnapshot_(previous, spreadsheet, sheet, layout);
      const current = previous.entries.map(entry => repaymentRange_(sheet, entry).getFormulas()[0]);
      if (current.every(row => row.every(formula => !formula))) {
        setRepaymentMode_(sheet, REPAYMENT_FREEZE.frozenValue);
        spreadsheet.toast('Repayments are already frozen.', 'Loan repayments', 6);
        return;
      }
      if (!current.every((row, index) => JSON.stringify(row) === JSON.stringify(previous.entries[index].formulas))) {
        throw new Error('A repayment operation was interrupted. Select Live to restore the saved formulas before freezing again.');
      }
      // The old MIME-type error could leave a backup before any cell was frozen.
      // Replace it only after verifying that every original formula is still present.
      props.deleteProperty(REPAYMENT_FREEZE.backupKey);
    }
    const range = sheet.getRange(4, layout.firstColumn, layout.labels.length, layout.width);
    const entries = collectRepaymentEntries_(layout, range.getFormulas(), range.getValues());
    if (!entries.length) throw new Error('No repayment formulas were found. Nothing was changed.');

    const snapshot = {spreadsheetId: spreadsheet.getId(), sheetId: sheet.getSheetId(), layout, entries};
    saveRepaymentSnapshot_(props, snapshot);
    // Capture every result before the first write: repayments depend on one another.
    entries.forEach(entry => repaymentRange_(sheet, entry).setValues([entry.values]));
    SpreadsheetApp.flush();
    entries.forEach(entry => {
      const target = repaymentRange_(sheet, entry);
      if (target.getFormulas()[0].some(Boolean) ||
          JSON.stringify(target.getValues()[0]) !== JSON.stringify(entry.values)) {
        throw new Error('Freeze verification failed. The backup is retained; run Unfreeze to recover.');
      }
    });
    setRepaymentMode_(sheet, REPAYMENT_FREEZE.frozenValue);
    spreadsheet.toast('Principal and interest repayments are frozen.', 'Loan repayments', 6);
  });
}

function unfreezeLoanRepayments() {
  return withRepaymentLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    const snapshot = loadRepaymentSnapshot_(props);
    if (!snapshot) throw new Error('No saved repayment formulas were found. Nothing was changed.');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getRepaymentSheet_(spreadsheet);
    validateRepaymentSnapshot_(snapshot, spreadsheet, sheet, getRepaymentLayout_(sheet));
    snapshot.entries.forEach(entry => repaymentRange_(sheet, entry).setFormulas([entry.formulas]));
    SpreadsheetApp.flush();
    snapshot.entries.forEach(entry => {
      if (JSON.stringify(repaymentRange_(sheet, entry).getFormulas()[0]) !== JSON.stringify(entry.formulas)) {
        throw new Error('Formula restoration failed. The backup is retained; retry Unfreeze.');
      }
    });
    // Remove the active marker only after every formula has been restored and checked.
    props.deleteProperty(REPAYMENT_FREEZE.backupKey);
    deleteRepaymentChunks_(props);
    setRepaymentMode_(sheet, REPAYMENT_FREEZE.liveValue);
    spreadsheet.toast('Repayment formulas restored. Calculations are live.', 'Loan repayments', 6);
  });
}

function validateRepaymentSnapshot_(snapshot, spreadsheet, sheet, layout) {
  if (snapshot.spreadsheetId !== spreadsheet.getId() || snapshot.sheetId !== sheet.getSheetId() ||
      JSON.stringify(snapshot.layout) !== JSON.stringify(layout)) {
    throw new Error('The loan layout or timeline changed. Restore its structure before unfreezing; the backup is retained.');
  }
}

function withRepaymentLock_(action) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another repayment operation is running. Try again shortly.');
  try {
    return action();
  } catch (error) {
    SpreadsheetApp.getActiveSpreadsheet().toast(error.message, 'Loan repayments', 10);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function getRepaymentSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(REPAYMENT_FREEZE.sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + REPAYMENT_FREEZE.sheetName);
  return sheet;
}

function getRepaymentLayout_(sheet) {
  const config = REPAYMENT_FREEZE;
  const headers = sheet.getRange(config.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (String(headers[config.paramColumn - 1]).toLowerCase() !== 'param') {
    throw new Error('Expected the Param header in In Loans!J3. Nothing was changed.');
  }
  let width = 0;
  while (headers[config.firstMonthColumn - 1 + width] instanceof Date) width++;
  const rowCount = sheet.getLastRow() - config.headerRow;
  if (!width || rowCount < 1) throw new Error('The loan timeline or loan rows are missing.');
  const labels = sheet.getRange(4, 1, rowCount, config.paramColumn).getDisplayValues()
    .map(row => [row[0], row[config.paramColumn - 1]]);
  return {
    firstColumn: config.firstMonthColumn,
    width,
    months: headers.slice(config.firstMonthColumn - 1, config.firstMonthColumn - 1 + width)
      .map(date => date.toISOString()),
    labels
  };
}

function collectRepaymentEntries_(layout, formulas, values) {
  const entries = [];
  layout.labels.forEach((labels, index) => {
    if (!REPAYMENT_FREEZE.params.includes(labels[1].trim().toLowerCase())) return;
    let column = 0;
    while (column < layout.width) {
      if (!formulas[index][column]) { column++; continue; }
      const start = column;
      while (column < layout.width && formulas[index][column]) {
        if (typeof values[index][column] !== 'number' || !Number.isFinite(values[index][column])) {
          throw new Error('Repayment formula has no numeric result at row ' + (index + 4) + '. Nothing was changed.');
        }
        column++;
      }
      entries.push({
        row: index + 4,
        column: layout.firstColumn + start,
        formulas: formulas[index].slice(start, column),
        values: values[index].slice(start, column)
      });
    }
  });
  return entries;
}

function repaymentRange_(sheet, entry) {
  return sheet.getRange(entry.row, entry.column, 1, entry.formulas.length);
}

function saveRepaymentSnapshot_(props, snapshot) {
  const config = REPAYMENT_FREEZE;
  const json = JSON.stringify(snapshot);
  // Compression keeps the full monthly formulas within document property storage limits.
  const encoded = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
  deleteRepaymentChunks_(props);
  const count = Math.ceil(encoded.length / config.chunkSize);
  for (let index = 0; index < count; index++) {
    props.setProperty(config.backupKey + '_' + index,
      encoded.slice(index * config.chunkSize, (index + 1) * config.chunkSize));
  }
  props.setProperty(config.backupKey, String(count));
  if (JSON.stringify(loadRepaymentSnapshot_(props)) !== json) {
    throw new Error('The repayment backup could not be verified. Nothing was changed.');
  }
}

function loadRepaymentSnapshot_(props) {
  const key = REPAYMENT_FREEZE.backupKey;
  const marker = props.getProperty(key);
  if (!marker) return null;
  const count = Number(marker);
  if (!Number.isInteger(count) || count < 1) throw new Error('Invalid repayment backup. Nothing was changed.');
  let encoded = '';
  for (let index = 0; index < count; index++) {
    const chunk = props.getProperty(key + '_' + index);
    if (!chunk) throw new Error('Incomplete repayment backup. Nothing was changed.');
    encoded += chunk;
  }
  const compressed = Utilities.newBlob(Utilities.base64Decode(encoded), 'application/gzip');
  return JSON.parse(Utilities.ungzip(compressed).getDataAsString('UTF-8'));
}

function deleteRepaymentChunks_(props) {
  const prefix = REPAYMENT_FREEZE.backupKey + '_';
  props.getKeys().filter(key => key.startsWith(prefix)).forEach(key => props.deleteProperty(key));
}
