// using an installable trigger since simple triggers have limited permissions (for example cannot write to log sheet)
// noinspection JSUnusedGlobalSymbols
function onInstallableOpen(e) {
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
    updateTimelineParamNames(e);
    updateUnitCountStatus(e);
  });
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  SpreadsheetApp.getUi().createMenu('SOL')
    // .addItem('Debug', '_debug')
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    .addItem('Toggle Write Logs to File', '_onToggleWriteLogsToFile')
    .addToUi();
}

function _onExportValuesXSLX() {
  SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
}

function _onToggleWriteLogsToFile() {
  SOLLibrary.toggleWriteToLogFile();
  SOLLibrary.alert('Write Logs To File',
    `Write Logs to File is ${SOLLibrary.getWriteToLogFileEnabled() ? 'enabled' : 'disabled'}`);
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;