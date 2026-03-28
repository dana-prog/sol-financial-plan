const CONSTRUCTION_COSTS_SHEET_NAME = 'Construction Costs';
const CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER = 'unit type';
const CONSTRUCTION_COSTS_UNIT_COUNT_COLUMN_HEADER = 'unit plan';

function getConstructionCostsTypeColRange() {
  return SOLLibrary.getColumnRange(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER, false);
}

function getUnitTypes() {
  return SOLLibrary
    .getColumnValues(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_UNIT_TYPE_COLUMN_HEADER, false);
}

function getConstructionCostsTotalUnitCount(unitTypeNameOrRow) {
  let unitTypeIndex;

  if (typeof unitTypeNameOrRow === 'number') {
    unitTypeIndex = unitTypeNameOrRow - 2;
  } else if (typeof unitTypeNameOrRow === 'string') {
    unitTypeIndex = getUnitTypes().indexOf(unitTypeNameOrRow);
    if (unitTypeIndex === -1) {
      throw new Error(`Unit type '${unitTypeNameOrRow}' does not exist in the Construction Costs sheet`);
    }
  } else {
    throw new Error(`Invalid parameter '${unitTypeNameOrRow}'`);
  }


  return SOLLibrary
    .getColumnValues(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_UNIT_COUNT_COLUMN_HEADER, false)[unitTypeIndex];
}

function _getConstructionCostsSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONSTRUCTION_COSTS_SHEET_NAME);
}

