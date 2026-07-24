function updateUnits(editEvent) {
  const sheetId = editEvent.range.getSheet().getSheetId();
  if (sheetId !== UNITS_SHEET_ID) {
    return;
  }

  const unitsSheet = getUnitsSheet();
  const colHeader = SOLLibrary.getColHeaderByNum(unitsSheet, editEvent.range.getColumn());
  if (colHeader !== UNITS_UNIT_TYPE_COLUMN_HEADER) {
    return;
  }

  const rowNum = editEvent.range.getRow();
  const numberFormat = getCountNumberFormat(editEvent.value);
  [UNITS_UNIT_COUNT_COLUMN_HEADER, UNITS_MISSING_UNITS_COLUMN_HEADER].forEach(colHeader => {
    const cellRange = unitsSheet.getRange(rowNum, SOLLibrary.getColNumByHeader(unitsSheet, colHeader));
    cellRange.setNumberFormat(numberFormat);
  })
}