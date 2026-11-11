const REPAYMENT_FREEZE = Object.freeze({
  sheetName: 'In Loans',
  toggleCell: 'A1',
  liveValue: 'Repayment Calculations Enabled',
  frozenValue: 'Repayment Calculations Disabled',
  headerRow: 3,
  paramColumn: 10,
  firstMonthColumn: 11,
  backupKey: 'inLoanRepaymentSnapshotV1',
  chunkSize: 8000,
  params: ['principal paid', 'interest paid']
});

function setRepaymentMode_(sheet, mode) {
  sheet.getRange(REPAYMENT_FREEZE.toggleCell).setValue(mode).setNote('');
}

const REPAYMENT_SHEETS = [
  REPAYMENT_FREEZE,
  {sheetName: 'Out Loans', paramColumn: 8, firstMonthColumn: 9, backupKey: 'outLoanRepaymentSnapshotV1'}
];

function freezeLoanRepayments() {
  return withRepaymentLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // Capture both sheets before any write: their calculations feed the same cash flow.
    const plans = REPAYMENT_SHEETS.map(config => {
      const sheet = getRepaymentSheet_(spreadsheet, config);
      const layout = getRepaymentLayout_(sheet, config);
      const range = repaymentTimelineRange_(sheet, layout);
      const formulas = range.getFormulas();
      const previous = loadRepaymentSnapshot_(props, config.backupKey);
      if (previous) {
        validateRepaymentSnapshot_(previous, spreadsheet, sheet, layout);
        const current = previous.entries.map(entry => repaymentEntryValues_(formulas, layout, entry));
        if (current.every(row => row.every(formula => !formula))) {
          return {config, sheet, range, snapshot: previous, entries: []};
        }
        if (!current.every((row, index) => JSON.stringify(row) === JSON.stringify(previous.entries[index].formulas))) {
          throw new Error('A repayment operation was interrupted. Select Repayment Calculations Enabled to restore the saved formulas before freezing again.');
        }
      }
      const entries = collectRepaymentEntries_(layout, formulas, range.getValues());
      const snapshot = {spreadsheetId: spreadsheet.getId(), sheetId: sheet.getSheetId(), layout, entries};
      return {config, sheet, range, snapshot, entries};
    });
    if (!plans.some(plan => plan.snapshot.entries.length)) {
      throw new Error('No repayment formulas were found. Nothing was changed.');
    }
    const changes = plans.filter(plan => plan.entries.length);
    changes.forEach(plan => saveRepaymentSnapshot_(props, plan.snapshot, plan.config.backupKey));
    changes.forEach(plan => plan.entries.forEach(entry => repaymentRange_(plan.sheet, entry).setValues([entry.values])));
    SpreadsheetApp.flush();
    changes.forEach(plan => {
      const frozenFormulas = plan.range.getFormulas();
      const frozenValues = plan.range.getValues();
      plan.entries.forEach(entry => {
        if (repaymentEntryValues_(frozenFormulas, plan.snapshot.layout, entry).some(Boolean) ||
            JSON.stringify(repaymentEntryValues_(frozenValues, plan.snapshot.layout, entry)) !== JSON.stringify(entry.values)) {
          throw new Error('Freeze verification failed. The backup is retained; run Unfreeze to recover.');
        }
      });
    });
    setRepaymentMode_(plans[0].sheet, REPAYMENT_FREEZE.frozenValue);
    return changes.length
      ? 'Principal and interest repayments in In Loans and Out Loans are frozen.'
      : 'Repayments are already frozen.';
  });
}

function unfreezeLoanRepayments() {
  return withRepaymentLock_(() => {
    const props = PropertiesService.getDocumentProperties();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const plans = REPAYMENT_SHEETS.map(config => {
      const snapshot = loadRepaymentSnapshot_(props, config.backupKey);
      // Older freezes saved only In Loans; Out Loans remains live until first frozen.
      if (!snapshot) return null;
      const sheet = getRepaymentSheet_(spreadsheet, config);
      validateRepaymentSnapshot_(snapshot, spreadsheet, sheet, getRepaymentLayout_(sheet, config));
      return {config, snapshot, sheet};
    }).filter(Boolean);
    if (!plans.length) throw new Error('No saved repayment formulas were found. Nothing was changed.');
    plans.forEach(plan => plan.snapshot.entries.forEach(entry => repaymentRange_(plan.sheet, entry).setFormulas([entry.formulas])));
    SpreadsheetApp.flush();
    plans.forEach(plan => {
      const restored = repaymentTimelineRange_(plan.sheet, plan.snapshot.layout).getFormulas();
      plan.snapshot.entries.forEach(entry => {
        if (JSON.stringify(repaymentEntryValues_(restored, plan.snapshot.layout, entry)) !== JSON.stringify(entry.formulas)) {
          throw new Error('Formula restoration failed. The backup is retained; retry Unfreeze.');
        }
      });
    });
    // Keep all recovery backups until both sheets have been restored and checked.
    plans.forEach(plan => {
      props.deleteProperty(plan.config.backupKey);
      deleteRepaymentChunks_(props, plan.config.backupKey);
    });
    setRepaymentMode_(getRepaymentSheet_(spreadsheet), REPAYMENT_FREEZE.liveValue);
    return 'Principal and interest repayment formulas restored in In Loans and Out Loans. Calculations are live.';
  });
}

function validateRepaymentSnapshot_(snapshot, spreadsheet, sheet, layout) {
  const changes = [];
  if (snapshot.sheetId !== sheet.getSheetId()) changes.push('sheet ID');
  if (snapshot.spreadsheetId !== spreadsheet.getId()) changes.push('spreadsheet ID');
  for (const field of ['firstColumn', 'width', 'months']) {
    if (JSON.stringify(snapshot.layout[field]) !== JSON.stringify(layout[field])) changes.push(field);
  }
  // Loan names are editable; row positions, parameter labels, and dates determine formula placement.
  const savedLabels = snapshot.layout.labels;
  const changedRow = savedLabels.findIndex((row, index) => row[1] !== layout.labels[index]?.[1]);
  if (changedRow !== -1) changes.push('parameter at row ' + (changedRow + 4));
  if (savedLabels.length !== layout.labels.length) changes.push('row count');
  if (changes.length) {
    throw new Error('The loan layout or timeline changed in ' + sheet.getName() + ' (' + changes.join('; ') + '). Restore its structure before unfreezing; the backup is retained.');
  }
}

function withRepaymentLock_(action) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('Another repayment operation is running. Try again shortly.');
  try {
    return action();
  } finally {
    lock.releaseLock();
  }
}

function getRepaymentSheet_(spreadsheet, config = REPAYMENT_FREEZE) {
  const sheet = spreadsheet.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + config.sheetName);
  return sheet;
}

function getRepaymentLayout_(sheet, config = REPAYMENT_FREEZE) {
  const headers = sheet.getRange(REPAYMENT_FREEZE.headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (String(headers[config.paramColumn - 1]).toLowerCase() !== 'param') {
    throw new Error('Expected the Param header in ' + config.sheetName + ' column ' + config.paramColumn + '. Nothing was changed.');
  }
  let width = 0;
  while (headers[config.firstMonthColumn - 1 + width] instanceof Date) width++;
  const rowCount = sheet.getLastRow() - REPAYMENT_FREEZE.headerRow;
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

function repaymentTimelineRange_(sheet, layout) {
  return sheet.getRange(4, layout.firstColumn, layout.labels.length, layout.width);
}

function repaymentEntryValues_(grid, layout, entry) {
  const startColumn = entry.column - layout.firstColumn;
  return grid[entry.row - 4].slice(startColumn, startColumn + entry.formulas.length);
}

function saveRepaymentSnapshot_(props, snapshot, backupKey = REPAYMENT_FREEZE.backupKey) {
  const config = REPAYMENT_FREEZE;
  const json = JSON.stringify(snapshot);
  // Compression keeps the full monthly formulas within document property storage limits.
  const encoded = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
  deleteRepaymentChunks_(props, backupKey);
  const count = Math.ceil(encoded.length / config.chunkSize);
  const chunks = {};
  for (let index = 0; index < count; index++) {
    chunks[backupKey + '_' + index] = encoded.slice(index * config.chunkSize, (index + 1) * config.chunkSize);
  }
  props.setProperties(chunks);
  props.setProperty(backupKey, String(count));
  if (JSON.stringify(loadRepaymentSnapshot_(props, backupKey)) !== json) {
    throw new Error('The repayment backup could not be verified. Nothing was changed.');
  }
}

function loadRepaymentSnapshot_(props, key = REPAYMENT_FREEZE.backupKey) {
  const properties = props.getProperties();
  const marker = properties[key];
  if (!marker) return null;
  const count = Number(marker);
  if (!Number.isInteger(count) || count < 1) throw new Error('Invalid repayment backup. Nothing was changed.');
  let encoded = '';
  for (let index = 0; index < count; index++) {
    const chunk = properties[key + '_' + index];
    if (!chunk) throw new Error('Incomplete repayment backup. Nothing was changed.');
    encoded += chunk;
  }
  const compressed = Utilities.newBlob(Utilities.base64Decode(encoded), 'application/gzip');
  return JSON.parse(Utilities.ungzip(compressed).getDataAsString('UTF-8'));
}

function deleteRepaymentChunks_(props, backupKey = REPAYMENT_FREEZE.backupKey) {
  const prefix = backupKey + '_';
  props.getKeys().filter(key => key.startsWith(prefix)).forEach(key => props.deleteProperty(key));
}
