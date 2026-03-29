/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  _createMenu();
}

// noinspection JSUnusedGlobalSymbols
function onEdit(event) {
  const oldValue = event.oldValue;
  const newValue = event.value;

  if (
    event.range.getSheet() === CONSTRUCTION_COSTS_SHEET_NAME
    && (oldValue !== newValue && oldValue !== '' && newValue !== '')
    && event.range.getColumn() === SOLLibrary.getColNumByHeader(CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER)
  ) {
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
