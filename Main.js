/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  _createMenu();
}

// noinspection JSUnusedGlobalSymbols
function onInstallableEdit(event) {
  const oldValue = event.oldValue;
  const newValue = event.value;
  const sheetId = event.range.getSheet().getSheetId();
  const col = event.range.getColumn();

  SOLLibrary.logArgs('Main', 'onEdit', {oldValue, newValue, sheetId, col});

  const unitTypeCol = SOLLibrary.getColNumByHeader(getConstructionCostsSheet(), CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER);
  SOLLibrary.logArgs('Main', 'onEdit', {unitTypeCol});

  if (
    sheetId === CONSTRUCTION_COSTS_SHEET_ID
    && (oldValue !== newValue && oldValue !== '' && newValue !== '')
    && col === unitTypeCol
  ) {
    SOLLibrary.log('Main', 'onEdit', `unit type changed from '${oldValue}' to '${newValue}'`);
    // unit type was changed -> update timeline construction params
    updateTimelineConstructionParams(oldValue, newValue);
  }

  const roleCol = SOLLibrary.getColNumByHeader(getStaffSheet(), STAFF_ROLE_COLUMN_HEADER);
  SOLLibrary.logArgs('Main', 'onEdit', {roleCol});

  if (
    sheetId === STAFF_SHEET_ID
    && (oldValue !== newValue && oldValue !== '' && newValue !== '')
    && col === roleCol
  ) {
    SOLLibrary.log('Main', 'onEdit', `role changed from '${oldValue}' to '${newValue}'`);
    // role was changed -> update timeline staff params
    updateTimelineStaffParams(oldValue, newValue);
  }
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  SpreadsheetApp.getUi().createMenu('SOL')
    .addItem('Sync Timeline Construction Params', '_onSyncTimelineConstructionParams')
    .addItem('Sync Timeline Staff Params', '_onSyncTimelineStaffParams')
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    .addItem('Toggle Alert Logs', '_onToggleAlertLogs')
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

function _onToggleAlertLogs() {
  SOLLibrary.toggleAlertLogs();
  SOLLibrary.alert('Alert Logs', `Alert logs are ${SOLLibrary.getLogAlerts() ? 'enabled' : 'disabled'}`);
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;