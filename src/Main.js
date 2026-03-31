// using an installable trigger since simple triggers have limited permissions (for example cannot write to log sheet)
// noinspection JSUnusedGlobalSymbols
function onInstallableOpen(e) {
  SOLLibrary.debugDuration('onInstallableOpen', () => {
    _createMenu();
    persistColumnHeadersToNums();
  });
}

// using an installable trigger since:
// simple triggers have limited permissions (for example cannot write to log sheet)
// simple triggers have a 30 seconds timeout
// noinspection JSUnusedGlobalSymbols
function onInstallableEdit(e) {
  SOLLibrary.debugDuration('onInstallableEdit', () => {
    const oldValue = e.oldValue;
    const newValue = e.value;
    const sheetId = e.range.getSheet().getSheetId();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    switch (sheetId) {
      case CONSTRUCTION_COSTS_SHEET_ID:
        SOLLibrary.debugDuration('onEditConstructionCostsSheet', () => {
          onEditConstructionCostsSheet(oldValue, newValue, row, col);
        });
        break;
      case STAFF_SHEET_ID:
        SOLLibrary.debugDuration('onEditStaffSheet', () => {
          onEditStaffSheet(oldValue, newValue, row, col);
        });
        break;
    }
  });
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  SpreadsheetApp.getUi().createMenu('SOL')
    // .addItem('Debug', '_debug')
    .addItem('Sync Timeline Construction Params', '_onSyncTimelineConstructionParams')
    .addItem('Sync Timeline Staff Params', '_onSyncTimelineStaffParams')
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    .addItem('Toggle Write Logs to File', '_onToggleWriteLogsToFile')
    .addToUi();
}

function _onSyncTimelineConstructionParams() {
  const unitTypes = getUnitTypes();
  syncTimelineConstructionParams(unitTypes);
}

function _onSyncTimelineStaffParams() {
  const staffRoles = getStaffRoles();
  syncTimelineStaffParams(staffRoles);
}

function _onExportValuesXSLX() {
  SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
}

function _onToggleWriteLogsToFile() {
  SOLLibrary.toggleWriteToLogFile();
  SOLLibrary.alert('Write Logs To File',
    `Write Logs to File is ${SOLLibrary.getWriteToLogFileEnabled() ? 'enabled' : 'disabled'}`);
}

function _debug() {
  SOLLibrary.debugDuration(
    'getRange',
    () => {
      SpreadsheetApp.getActiveSpreadsheet().getSheetById(CONSTRUCTION_COSTS_SHEET_ID).getRange(2, 1, 1, 1)
        .setValue('test');
    });
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;