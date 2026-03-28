const CONSTRUCTION_COSTS_SHEET_NAME = 'Construction Costs';
const CONSTRUCTION_COSTS_TYPE_COLUMN_HEADER = 'unit type';

function getConstructionCostsTypeColRange() {
  return SOLLibrary.getColumnRange(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_TYPE_COLUMN_HEADER, false);
}

function getUnitTypes() {
  return SOLLibrary
    .getColumnValues(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_TYPE_COLUMN_HEADER, false);
}

