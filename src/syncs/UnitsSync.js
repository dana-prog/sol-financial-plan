function updateUnits(editEvent) {
  const sheetId = editEvent.range.getSheet().getSheetId();
  if (sheetId !== UNITS_SHEET_ID) {
    return;
  }

  const colHeader = getColumnHeaderByNum(sheetId, editEvent.range.getColumn());
  if (colHeader !== UNITS_UNIT_TYPE_COLUMN_HEADER) {
    return;
  }

  const sheet = getUnitsSheet();
  const rowNum = editEvent.range.getRow();
  const numberFormat = getCountNumberFormat(editEvent.value);
  [UNITS_UNIT_COUNT_COLUMN_HEADER, UNITS_MISSING_UNITS_COLUMN_HEADER].forEach(colHeader => {
    const cellRange = sheet.getRange(rowNum, getColumnNumByHeader(sheetId, colHeader));
    cellRange.setNumberFormat(numberFormat);
  })
}