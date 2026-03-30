/**
 * Returns a sheet by name, optionally creating it if it does not exist.
 *
 * @param {string} sheetName Sheet name.
 * @param {boolean} [create=false] Whether to create the sheet if missing.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} The sheet or null.
 */
function getSheet(sheetName, create = false) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet != null || !create) {
    return sheet;
  }

  return spreadsheet.insertSheet(sheetName);
}

/**
 * Returns the column number for the given header name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|GoogleAppsScript.Spreadsheet.Range} rangeOrSheet Sheet or range.
 * @param {string} colHeader Column header name.
 * @returns {number} Column number (1-based).
 */
function getColNumByHeader(rangeOrSheet, colHeader) {
  const headers = getHeaderMap(rangeOrSheet);
  return headers[colHeader] + 1;
}

/**
 * Appends multiple rows to the sheet.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<Array<*>>} rows Rows to append.
 */
function appendRowsToSheet(sheetName, rows) {
  const sheet = getSheet(sheetName);
  sheet.getRange(sheetName.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}

/**
 * Sorts the sheet by a column.
 *
 * @param {string} sheetName Sheet name.
 * @param {string} colName Column header name.
 * @param {boolean} [ascending=true] Sort order.
 */
function sortSheet(sheetName, colName, ascending = true) {
  const sheet = getSheet(sheetName);
  const colIndex = getColNumByHeader(sheet, colName);
  sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .sort({
      column: colIndex,
      ascending: ascending
    });
}

/**
 * Returns all values from a column.
 *
 * @param {string} sheetName Sheet name.
 * @param {string|number} colHeader Column header name or index.
 * @param {boolean} [includeHeader=true] Whether to include header row.
 * @returns {Array<*>} Column values.
 */
function getColumnValues(sheetName, colHeader, includeHeader = true) {
  const colRange = getColumnRange(sheetName, colHeader, includeHeader);
  return [].concat.apply([], colRange.getValues());
}

/**
 * Returns an object mapping headers to values for a row.
 *
 * @param {string} sheetName Sheet name.
 * @param {number} rowNum Row number (1-based).
 * @returns {Object} Row values keyed by header.
 */
function getRowValues(sheetName, rowNum) {
  const sheet = getSheet(sheetName);
  const headerMap = getHeaderMap(sheet);
  const lastCol = sheet.getLastColumn();
  const rowData = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  logArgs('SheetUtils', 'getRowValues', {rowData});

  const valuesObj = {};

  Object.entries(headerMap).forEach(([header, colIndex]) => {
    valuesObj[header] = rowData[colIndex];
  });

  return valuesObj;
}

/**
 * Returns the currently selected row number.
 * @param {string} sheetName Sheet name.
 *
 * @returns {number} Row number (1-based).
 */
function getSelectedRow(sheetName) {
  const sheet = getSheet(sheetName);
  const selectedRange = sheet.getActiveRange();
  return selectedRange.getRow();
}

/**
 * Returns a map of header to column index (0-based).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet|GoogleAppsScript.Spreadsheet.Range} rangeOrSheet Sheet or range.
 * @returns {Object<string, number>|-1} Header map or -1 if empty.
 */
function getHeaderMap(rangeOrSheet) {
  const range = rangeOrSheet.getRange ? rangeOrSheet.getRange('1:1') : rangeOrSheet;
  const rangeValues = range.getValues();

  if (rangeValues.length === 0 || rangeValues[0].length === 0) {
    log('SheetUtils', 'getHeadersMap', 'No values found in the range. Returning -1');
    return -1;
  }

  const headerValues = rangeValues[0];

  return Object.fromEntries(
    headerValues.map((header, index) => [header, index])
  );
}

/**
 * Converts objects to sheet rows based on headers.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<Object>} objects Objects to convert (prop names must match column headers).
 * @returns {Array<Array<*>>} Rows aligned with sheet headers.
 */
function generateSheetRows(sheetName, objects) {
  const sheet = getSheet(sheetName);
  const headerMap = getHeaderMap(sheet);
  const headerCount = Object.keys(headerMap).length;
  return objects.map(doc => {
    const row = Array(headerCount).fill('');

    Object.entries(doc).forEach(([propName, propValue]) => {
      if (propName in headerMap) {
        const propIndex = headerMap[propName];
        row[propIndex] = propValue;
      } else {
        log('SheetUtils', 'generateSheetRows', `No header for property: ${propName}`);
      }
    });

    return row;
  });
}

/**
 * Selects a row and scrolls it into view.
 *
 * @param {string} sheetName Sheet name.
 * @param {number} rowNum Row number (1-based).
 */
function selectRow(sheetName, rowNum) {
  const sheet = getSheet(sheetName);
  sheet.setActiveRange(
    sheet.getRange(rowNum, 1, 1, sheet.getLastColumn())
  );
  sheet.setActiveSelection(
    sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getA1Notation()
  );
}

/**
 * Appends a row, optionally runs a callback, and selects it.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<*>} row Row values.
 * @param {Function|null} [callback=null] Callback with (sheet, newRowNum).
 */
function addRow(sheetName, row, callback = null) {
  const sheet = getSheet(sheetName);
  sheet.appendRow(row);

  const newRowNum = sheet.getLastRow();
  callback && callback(sheet, newRowNum);

  selectRow(sheet, newRowNum);
}

/**
 * Updates an existing row in the given sheet with the provided values.
 *
 * @param {string} sheetName Sheet name.
 * @param {Array<*>} row - Row values.
 * @param {number} rowNum - The row number to update.
 */
function updateRow(sheetName, row, rowNum) {
  const sheet = getSheet(sheetName);
  logArgs('SheetUtils', 'updateRow', {
    sheet,
    row,
    rowNum
  });
  const range = sheet.getRange(rowNum, 1, 1, row.length);
  range.setValues([row]);
}

/**
 * Returns all row numbers where the specified column equals the given value (excluding header).
 *
 * @param {string} sheetName Sheet name.
 * @param {string} colHeader Column header name.
 * @param {*} value Value to match.
 * @returns {number[]} Matching row numbers (1-based).
 */
function getRowNumbers(sheetName, colHeader, value) {
  const sheet = getSheet(sheetName);
  const colIndex = getColNumByHeader(sheet, colHeader);

  const values = sheet
    .getRange(2, colIndex, sheet.getLastRow() - 1)
    .getValues()
    .flat();

  const rows = [];
  values.forEach((v, i) => {
    if (v === value) {
      rows.push(i + 2); // +2 → header + 1-based
    }
  });

  return rows;
}

/**
 * Returns the range of a column by header or index.
 *
 * @param {string} sheetName Sheet name.
 * @param {string|number} colHeader Column header name or index.
 * @param {boolean} [includeHeader=true] Whether to include header row.
 * @returns {GoogleAppsScript.Spreadsheet.Range} Column range.
 */
function getColumnRange(sheetName, colHeader, includeHeader = true) {
  const sheet = getSheet(sheetName);
  const colIndex = typeof colHeader === 'number' ? colHeader : getColNumByHeader(sheet, colHeader);
  const startRow = includeHeader ? 1 : 2;
  const numRows = sheet.getLastRow() - startRow + 1;
  return sheet.getRange(startRow, colIndex, numRows);
}

/**
 * Checks whether one range is fully contained in another range.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} innerRange The range to test (candidate inside).
 * @param {GoogleAppsScript.Spreadsheet.Range} outerRange The containing range.
 * @returns {boolean} TRUE if `innerRange` is completely inside `outerRange`, otherwise FALSE.
 */
function isInside(innerRange, outerRange) {
  return innerRange.getRow() >= outerRange.getRow() &&
    innerRange.getLastRow() <= outerRange.getLastRow() &&
    innerRange.getColumn() >= outerRange.getColumn() &&
    innerRange.getLastColumn() <= outerRange.getLastColumn();
}