
/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  _createMenu();
}

function onEdit(event) {
  const oldValue = event.oldValue;
  const newValue = event.value;

  if (oldValue !== newValue && oldValue !== '' && newValue !== '') {
    updateTimelineConstructionParams(oldValue, newValue);
  }
}
/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  const ui = SpreadsheetApp.getUi();


  ui.createMenu('SOL')
    .addItem('Sync Timeline Construction Params', '_onSyncTimelineConstructionParams')
    .addItem('Export as XLSX (Values Only)', '_onExportValuesXSLX')
    .addToUi();
}

function _onSyncTimelineConstructionParams() {
  const unitTypes = getUnitTypes();
  syncTimelineConstructionParams(unitTypes);
}

function _onExportValuesXSLX() {
  SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;
