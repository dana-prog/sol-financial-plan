function updateUnitCountStatus(editEvent) {
  // check whether the change was in the unit count column of the Units sheet
  if (editEvent.range.getSheet()
    .getSheetId() !== UNITS_SHEET_ID && editEvent.range.getColumn() === getUnitTypeColNum()) {
    // get param name and unit count
    const unitType = editEvent.range.getSheet().getRange(editEvent.range.getRow(), getUnitTypeColNum).getValue();
    const unitCount = editEvent.range.getValue();

    SOLLibrary.logArgs('UnitCountSync', {
      unitType,
      unitCount
    });
  }

  // check whether the change was in a construction count param in the Construction Timeline sheet

  // get construction count sum for the param

  // get unit count for the unit type

  // if different format the unit count cell in the Units sheet
}