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
  SOLLibrary.alert('Done',
    `'${TIMELINE_UNITS_COUNT_CATEGORY}' and '${TIMELINE_UNITS_COSTS_CATEGORY}' params in the 'Timeline' sheet were synced according to the unit types in the 'Construction Costs' sheet`)
}

function _onSyncTimelineStaffParams() {
  const staffRoles = getStaffRoles();
  syncTimelineStaffParams(staffRoles);
  SOLLibrary.alert('Done',
    `'${TIMELINE_STAFF_CATEGORY}', '${TIMELINE_MONTHLY_NET_SALARIES_CATEGORY}' and '${TIMELINE_NET_SALARIES_CATEGORY}' params in the 'Timeline' sheet were synced according to the unit types in the 'Construction Costs' sheet`)

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