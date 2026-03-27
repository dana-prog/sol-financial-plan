const CONSTRUCTION_COSTS_SHEET_NAME = '_Construction Costs';
const CONSTRUCTION_COSTS_TYPE_COLUMN_HEADER = 'type';

function getConstructionCostsTypeColRange() {
  return SOLLibrary.getColumnRange(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_TYPE_COLUMN_HEADER, false);
}

function getUnitTypes() {
  return SOLLibrary
    .getColumnValues(CONSTRUCTION_COSTS_SHEET_NAME, CONSTRUCTION_COSTS_TYPE_COLUMN_HEADER, false);
}

