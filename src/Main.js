// using an installable trigger since simple triggers have limited permissions (for example cannot write to log sheet)
// noinspection JSUnusedGlobalSymbols
function onInstallableOpen() {
  _createMenu();
}

// using an installable trigger since:
// simple triggers have limited permissions (for example cannot write to log sheet)
// simple triggers have a 30 seconds timeout
// noinspection JSUnusedGlobalSymbols
function onInstallableEdit(e) {
  updateUnits(e);
  updateTimelineParamNames(e);
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  SpreadsheetApp.getUi().createMenu('SOL')
    .addItem('Set Timeline Borders', '_onSetTimelineBorders')
    .addSeparator()
    .addItem('Recalculate Borrowing Payments and Tax', 'showRepaymentCalculation')
    .addItem('Sync Wise Transactions', 'syncWiseScheduled')
    .addSeparator()
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    // .addSeparator()
    // .addItem('Toggle Write Logs To File', '_onToggleWriteLogsToFile')
    // .addItem('Export Named Functions', '_onExportNamedFunctions')
    // .addItem('Export Formulas', '_onExportFormulas')
    .addToUi();
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

function _onSetTimelineBorders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (['Details', 'All Expenses'].includes(sheet.getName())) {
    setDetailsSheetBorders(sheet);
    return;
  }
  setTimelineSheetBorders();
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;
