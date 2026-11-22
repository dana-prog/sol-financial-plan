function refreshAllExpenses() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) throw new Error('Another refresh is running. Please try again shortly.');
  try {
    SpreadsheetApp.flush();
    const result = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: ["'All Expenses'"],
      fields: 'sheets(properties,basicFilter,merges,data(rowData(values(userEnteredValue,effectiveValue,formattedValue)),rowMetadata(hiddenByFilter,hiddenByUser)))',
    });
    const sheet = result.sheets?.[0];
    if (!sheet) throw new Error('Could not find All Expenses.');
    const requests = _allExpensesRefreshRequests(sheet);
    Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
    return 'All Expenses updates automatically. Formatting refreshed.';
  } finally {
    lock.releaseLock();
  }
}

function _allExpensesRefreshRequests(sheet) {
  const rows = sheet.data?.[0]?.rowData || [];
  const headers = rows[0]?.values || [];
  if (!rows[0]?.values?.[0]?.userEnteredValue?.formulaValue) {
    throw new Error('All Expenses must retain its live formula in A1. No changes were made.');
  }
  for (let index = 0; index < rows.length; index++) {
    for (const cell of rows[index].values || []) {
      if (cell.effectiveValue?.errorValue) {
        throw new Error(`All Expenses row ${index + 1}: ${cell.effectiveValue.errorValue.message || cell.effectiveValue.errorValue.type}. No changes were made.`);
      }
    }
  }
  if (!ALL_EXPENSES_HEADERS.every((name, column) => headers[column]?.effectiveValue?.stringValue === name)) {
    throw new Error('All Expenses headers do not match the expense dataset. No changes were made.');
  }
  const sheetId = sheet.properties.sheetId;
  const endRowIndex = sheet.properties.gridProperties.rowCount;
  const range = {sheetId, startRowIndex: 0, endRowIndex, startColumnIndex: 0, endColumnIndex: ALL_EXPENSES_HEADERS.length};
  const body = {...range, startRowIndex: 1};
  const requests = [];
  // Format the spill area without replacing the formula or changing filtered/hidden rows.
  const columns = [
    [0, 4, {type: 'TEXT', pattern: '@'}],
    [4, 5, {type: 'NUMBER', pattern: '0'}],
    [5, 6, {type: 'TEXT', pattern: '@'}],
    [6, 7, {type: 'DATE', pattern: 'mmmm'}],
    [7, 10, {type: 'NUMBER', pattern: '฿ #,##0,"K"'}],
    [10, 12, {type: 'TEXT', pattern: '@'}],
    [12, 13, {type: 'NUMBER', pattern: '0'}],
  ];
  for (const [startColumnIndex, endColumnIndex, numberFormat] of columns) {
    requests.push({repeatCell: {range: {...body, startColumnIndex, endColumnIndex},
      cell: {userEnteredFormat: {numberFormat}}, fields: 'userEnteredFormat.numberFormat'}});
  }
  // Include spare rows so newly spilled records remain part of the existing filter.
  const filter = {range};
  if (sheet.basicFilter?.filterSpecs) filter.filterSpecs = sheet.basicFilter.filterSpecs;
  else if (sheet.basicFilter?.criteria) filter.criteria = sheet.basicFilter.criteria;
  // Reapplying a saved sort moves the array anchor; ordering belongs to the formula.
  requests.push({setBasicFilter: {filter}});
  requests.push(..._detailsBorderRequests(sheet));
  return requests;
}
