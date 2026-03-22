/**
 * An event handler called when the spreadsheet is opened. Initializes the SOL menu.
 */
// noinspection JSUnusedGlobalSymbols
function onOpen() {
  _createMenu();
}

/**
 * Creates the SOL menu.
 * @private
 */
function _createMenu() {
  const ui = SpreadsheetApp.getUi();


  ui.createMenu('SOL')
    .addItem('Export as XLSX (Values Only)', '_exportValuesXSLX')
    .addToUi();
}

function _exportValuesXSLX() {
  SOLLibrary.exportValuesXSLX('deleteTmpExportResources');
}

// do not remove this line. see SOLLibrary.exportValuesXSLX docs for details
const deleteTmpExportResources = SOLLibrary.deleteTmpExportResources;