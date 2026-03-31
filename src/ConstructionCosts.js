const CONSTRUCTION_COSTS_SHEET_ID = 1436796628;
const CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER = 'unit type';

function onEditConstructionCostsSheet(oldValue, newValue, rowNum, colNum) {
  if (oldValue && newValue && colNum === getUnitTypeColNum()) {
    // unit type was changed -> update timeline construction params
    updateTimelineConstructionParams(oldValue, newValue);
  }
}

function getUnitTypes() {
  const sheet = getConstructionCostsSheet();
  const colRange = sheet.getRange(2, getUnitTypeColNum(), sheet.getLastRow() - 1, 1);
  return [].concat.apply([], colRange.getValues());
}

function getConstructionCostsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetById(CONSTRUCTION_COSTS_SHEET_ID);
}

function getUnitTypeColNum() {
  return getColumnNumByHeader(CONSTRUCTION_COSTS_SHEET_ID, CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER);
}

// const CONSTRUCTION_COSTS_UNIT_COUNT_COLUMN_HEADER = 'unit plan';

// function getConstructionCostsTotalUnitCount(unitTypeNameOrRow) {
//   let unitTypeIndex;
//
//   if (typeof unitTypeNameOrRow === 'number') {
//     unitTypeIndex = unitTypeNameOrRow - 2;
//   } else if (typeof unitTypeNameOrRow === 'string') {
//     unitTypeIndex = getUnitTypes().indexOf(unitTypeNameOrRow);
//     if (unitTypeIndex === -1) {
//       throw new Error(`Unit type '${unitTypeNameOrRow}' does not exist in the Construction Costs sheet`);
//     }
//   } else {
//     throw new Error(`Invalid parameter '${unitTypeNameOrRow}'`);
//   }
//
//
//   return SOLLibrary
//     .getColumnValues(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_UNIT_COUNT_COLUMN_HEADER, false)[unitTypeIndex];
// }