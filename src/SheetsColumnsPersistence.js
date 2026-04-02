const _COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME = 'columnHeadersToNums';
let colMapping;

function persistSheetsColumnsMap() {
  const mapping = {};
  const sheetIds = [UNITS_SHEET_ID, STAFF_SHEET_ID];
  sheetIds.forEach(sheetId => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(sheetId);
    mapping[sheetId] = {};
    const columnHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    columnHeaders.forEach((header, i) => {
      mapping[sheetId][header] = i + 1;
      mapping[sheetId][i + 1] = header;
    });
  });

  const timelineColMapping = {
    category: 1,
    1: 'category',
    param: 2,
    2: 'param',
  };

  mapping[CONSTRUCTION_TIMELINE_SHEET_ID] = timelineColMapping;
  mapping[INFRASTRUCTURE_TIMELINE_SHEET_ID] = timelineColMapping;
  mapping[MAIN_TIMELINE_SHEET_ID] = timelineColMapping;

  PropertiesService
    .getDocumentProperties()
    .setProperty(_COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME, JSON.stringify(mapping));

  SOLLibrary.logArgs(
    'persistColumnHeadersToNums',
    'columnHeadersToNums',
    mapping,
    `Property: ${_COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME} was set`);
}

function getColumnNumByHeader(sheetId, header) {
  return _getColumnMapValue(sheetId, header);
}

function getColumnHeaderByNum(sheetId, num) {
  return _getColumnMapValue(sheetId, num);
}

function _getColumnMapValue(sheetId, key) {
  if (!colMapping) {
    colMapping = JSON.parse(
      PropertiesService.getDocumentProperties().getProperty(_COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME));
  }

  if (!colMapping) {
    throw new Error(`Columns mapping is not initialized`);
  }

  if (!colMapping[sheetId]) {
    throw new Error(`Columns mapping for sheet ${sheetId} is not initialized`);
  }

  if (!colMapping[sheetId][key]) {
    throw new Error(`Columns mapping for column '${key}' does not exist in sheet ${sheetId}`);
  }

  return colMapping[sheetId][key];
}