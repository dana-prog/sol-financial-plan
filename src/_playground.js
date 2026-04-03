function test() {
  const unitsSheet = getUnitsSheet();
  const unitTypes = unitsSheet.getRange(2, 2, unitsSheet.getLastRow() - 1, 1).getValues().flat();
  SOLLibrary.logArgs('_playground', 'test', {unitTypes});

  for (let i = 0; i < unitTypes.length; i++) {
    const unitType = unitTypes[i];
    const rowNum = i + 2;
    const numberFormat = getCountNumberFormat(unitType);
    [UNITS_UNIT_COUNT_COLUMN_HEADER, UNITS_MISSING_UNITS_COLUMN_HEADER].forEach(colHeader => {
      const cellRange = unitsSheet.getRange(rowNum, getColumnNumByHeader(UNITS_SHEET_ID, colHeader));
      cellRange.setNumberFormat(numberFormat);
    })
  }
}

function updateTimelineParamsFormat() {
  const unitsSheet = getUnitsSheet();
  const unitTypes = unitsSheet.getRange(2, 2, unitsSheet.getLastRow() - 1, 1).getValues().flat();
  SOLLibrary.logArgs('_playground', 'test', {unitTypes});

  unitTypes.forEach(unitType => {
    const paramName = unitType + TIMELINE_COMPLETED_UNITS_PARAM_POSTFIX;
    const paramRowNum = _getTimelineParamRowNumber(timelineSheet, paramName);
    SOLLibrary.logArgs('_playground', 'test', {unitType, paramName, paramRowNum});
    const paramRange = timelineSheet.getRange(`${paramRowNum}:${paramRowNum}`);
    const numberFormat = getCountNumberFormat(unitType);
    SOLLibrary.logArgs('_playground', 'test', {range: paramRange.getA1Notation(), numberFormat});
    paramRange.setNumberFormat(numberFormat);
  });
}