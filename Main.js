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

// function onInstallableEdit(event) {
//   SOLLibrary.log('Main', 'onEdit', 'start');
//
//   const data = {
//     sheetId: event.range.getSheet().getSheetId(),
//     oldValue: event.oldValue,
//     newValue: event.value,
//     row: event.range.getRow(),
//     col: event.range.getColumn(),
//   };
//
//   SOLLibrary.logArgs('Main', 'onEdit', data);
//
//   const trigger =
//     ScriptApp
//       .newTrigger(doOnEdit)
//       .timeBased()
//       .after(2 * 1000)
//       .create();
//
//   PropertiesService
//     .getScriptProperties()
//     .setProperty(`trigger_${trigger.getUniqueId()}`, JSON.stringify(data));
//
//   SOLLibrary.logArgs('Main', 'onEdit', {
//     triggerUid: trigger.getUniqueId(),
//     propName: `trigger_${trigger.getUniqueId()}`
//   });
// }
//
// function doOnEdit() {
//   const triggerId = event.triggerUid;
//   const propertyName = `trigger_${triggerId}`;
//   const props = PropertiesService.getScriptProperties();
//   const data = JSON.parse(props.getProperty(propertyName));
//
//   SOLLibrary.logArgs('Main', 'doOnEdit', data);
//
//   if (
//     data.sheetId === CONSTRUCTION_COSTS_SHEET_ID
//     && (data.oldValue !== data.newValue && data.oldValue !== '' && data.newValue !== '')
//     && data.col === SOLLibrary.getColNumByHeader(getConstructionCostsSheet(),
//       CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER)
//   ) {
//     SOLLibrary.log('Main', 'onEdit', `unit type changed from '${data.oldValue}' to '${data.newValue}'`);
//     // unit type was changed -> update timeline construction params
//     updateTimelineConstructionParams(data.oldValue, data.newValue);
//   }
//
//   if (
//     data.sheetId === STAFF_SHEET_ID
//     && (data.oldValue !== data.newValue && data.oldValue !== '' && data.newValue !== '')
//     && data.col === SOLLibrary.getColNumByHeader(getStaffSheet(), STAFF_ROLE_COLUMN_HEADER)
//   ) {
//     // role was changed -> update timeline staff params
//     updateTimelineStaffParams(data.oldValue, data.newValue);
//   }
//   props.deleteProperty(propertyName);
//
//
// // delete the trigger itself (time-based triggers are persistent and eventually will hit quota limits)
//   ScriptApp.getProjectTriggers().forEach(trigger => {
//     if (trigger.getUniqueId() === triggerId) {
//       ScriptApp.deleteTrigger(trigger);
//     }
//   });
// }