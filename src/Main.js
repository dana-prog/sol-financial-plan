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
    updateUnits(e);
    updateTimelineParamNames(e);
  });

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CF_SHEET_NAME) return;
  if (e.range.getA1Notation() !== CF_TOGGLE) return;
                                                                                                               
  const value = String(e.value || '').trim();          
  if (value === 'Static Values (fast)') freezeCashFlow(sheet);
  else if (value === 'Auto Calculated Formulas (slow)') unfreezeCashFlow(sheet);
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  SpreadsheetApp.getUi().createMenu('SOL')
    // .addItem('Debug', '_debug')
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    .addItem('Export Perf Audit (Formulas + Named Ranges)', 'exportPerfAudit')
    .addSeparator()
    .addItem('PerfFix: Vectorize Interest Cascade (Dry Run)', 'vectorizeInterestCascadeBatch1DryRun')
    .addItem('PerfFix: Vectorize Interest Cascade (Apply)', 'vectorizeInterestCascadeBatch1Apply')
    .addSeparator()
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