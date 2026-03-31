const COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME = 'columnHeadersToNums';
let columnHeadersToNums;

function getCountNumberFormat(baseValue) {
  const abbreviation = _getUnitTypeAbbreviation(baseValue);
  return `[=1]0 "${abbreviation}";0 "${abbreviation}s"`;
}

function persistColumnHeadersToNums() {
  const res = {};
  const sheetIds = [CONSTRUCTION_COSTS_SHEET_ID, STAFF_SHEET_ID];
  sheetIds.forEach(sheetId => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(sheetId);
    res[sheetId] = {};
    const columnHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    columnHeaders.forEach((header, i) => {
      res[sheetId][header] = i + 1;
    });
  });

  SOLLibrary.logArgs(
    'persistColumnHeadersToNums',
    'columnHeadersToNums',
    res,
    `setting property: ${COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME}`);
  PropertiesService
    .getDocumentProperties()
    .setProperty(COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME, JSON.stringify(res));
}

function getColumnNumByHeader(sheetId, header) {
  if (!columnHeadersToNums) {
    const propertyValue = PropertiesService.getDocumentProperties().getProperty(COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME);
    columnHeadersToNums = JSON.parse(
      PropertiesService.getDocumentProperties().getProperty(COLUMN_HEADERS_TO_NUMS_PROPERTY_NAME));
  }

  if (!columnHeadersToNums) {
    throw new Error(`Column headers to numbers map is not initialized`);
  }

  if (!columnHeadersToNums[sheetId]) {
    throw new Error(`Column headers to numbers map for sheet ${sheetId} is not initialized`);
  }

  if (!columnHeadersToNums[sheetId][header]) {
    throw new Error(`Column header '${header}' does not exist in sheet ${sheetId}`);
  }

  return columnHeadersToNums[sheetId][header];
}

