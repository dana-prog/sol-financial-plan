
function getUnitsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetById(UNITS_SHEET_ID);
}

// function getUnitTypeColNum() {
//   return _getColumnMapValue(UNITS_SHEET_ID, UNITS_UNIT_TYPE_COLUMN_HEADER);
// }

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