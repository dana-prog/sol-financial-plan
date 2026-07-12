// using an installable trigger since simple triggers have limited permissions (for example cannot write to log sheet)
// noinspection JSUnusedGlobalSymbols
function onInstallableOpen() {
  SOLLibrary.debugDuration('onInstallableOpen', () => {
    _createMenu();
    persistSheetsColumnsMap();
  });
}

// using an installable trigger since:
// simple triggers have limited permissions (for example cannot write to log sheet)
// simple triggers have a 30 seconds timeout
// noinspection JSUnusedGlobalSymbols
function onInstallableEdit(e) {
  SOLLibrary.debugDuration('onInstallableEdit', () => {
    updateUnits(e);
    updateTimelineParamNames(e);
  });

  // const sheet = e.range.getSheet();
  // if (sheet.getName() !== CF_SHEET_NAME) return;
  // if (e.range.getA1Notation() !== CF_TOGGLE) return;
  //
  // const value = String(e.value || '').trim();
  // if (value === 'Static Values (fast)') freezeCashFlow();
  // else if (value === 'Auto Calculated Formulas (slow)') unfreezeCashFlow();
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  SpreadsheetApp.getUi().createMenu('SOL')
    .addItem('Show Details', '_onShowDetails')
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    // .addItem('Export Named Functions', '_onExportNamedFunctions')
    // .addItem('Export Formulas', '_onExportFormulas')
    .addToUi();
}

function _onShowDetails() {
  const activeSheet = SpreadsheetApp.getActive().getActiveSheet();
  const activeCell = activeSheet.getActiveCell();

  const row = activeCell.getRow();
  const col = activeCell.getColumn();

  // Assumes:
  // - Parameters are in column B.
  // - Month headers are in row 3.
  const param = activeSheet.getRange(row, 2).getValue();
  const month = activeSheet.getRange(3, col).getValue();

  showDetails(param, month);
}

function _onExportValuesXSLX() {
  SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
}

function _onExportNamedFunctions() {
  SOLLibrary.exportNamedFunctionsJSON('deleteTmpExportResources');
}

function _onExportFormulas() {
  SOLLibrary.exportFormulasJSON('deleteTmpExportResources');
}

function _onToggleWriteLogsToFile() {
  SOLLibrary.toggleWriteToLogFile();
  SOLLibrary.alert('Write Logs To File',
    `Write Logs to File is ${SOLLibrary.getWriteToLogFileEnabled() ? 'enabled' : 'disabled'}`);
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;