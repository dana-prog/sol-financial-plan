const TIMELINE_SHEET_NAME = '_Timeline';
const TIMELINE_CATEGORY_COLUMN_NUMBER = 1;
const TIMELINE_PARAM_NAME_COLUMN_NUMBER = 2;
const TIMELINE_FIRST_QUARTER_COLUMN_NUMBER = 3;
const TIMELINE_HEADER_ROW_NUM = 3;
const TIMELINE_CONSTRUCTION_PLAN_CATEGORY = 'Construction Plan';
const TIMELINE_CONSTRUCTION_COSTS_CATEGORY = 'Construction Costs';
const TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX = ' construction count';
const TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX = ' construction cost';

function syncTimelineConstructionParams(unitTypes) {
  _syncTimelineConstructionParams()
}

function _syncTimelineConstructionParams(category, paramPostfix, unitTypes) {
  const params = _getCategoryParams(TIMELINE_CONSTRUCTION_PLAN_CATEGORY);
  const expectedParams = unitTypes.map(unitType => unitType + paramPostfix);
  if (params.join() === expectedParams.join()) {
    return;
  }


}

function addTimelineConstructionTypeParams(type, paramPosition) {
  addTimelineParam(
    TIMELINE_CONSTRUCTION_PLAN_CATEGORY,
    type + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX,
    paramPosition
  );

  _getParamValuesRange(type + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX).setNumberFormat(`[=1]0 "${type}";0 "${type}s"`);

  addTimelineParam(
    TIMELINE_CONSTRUCTION_COSTS_CATEGORY,
    type + TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX,
    paramPosition
  );
}

function updateTimelineUnitTypeParams(oldType, newType) {
  updateTimelineParam(
    oldType + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX,
    newType + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX,
    `[=1]0 "${newType}";0 "${newType}s"`
  );
  updateTimelineParam(
    oldType + TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX,
    newType + TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX
  )
}

function removeTimelineConstructionTypeParams(typeIndex) {
  removeTimelineParam(TIMELINE_CONSTRUCTION_PLAN_CATEGORY, typeIndex);
  removeTimelineParam(TIMELINE_CONSTRUCTION_COSTS_CATEGORY, typeIndex);
}

function addTimelineParam(category, paramName, paramPosition) {
  SOLLibrary.log('Timeline', 'addTimelineParam', '1');

  const sheet = _getTimelineSheet();
  const {
    categoryStartRow,
    categoryEndRow,
  } = _getCategoryRowBoundaries(category);

  const categoryParamCount = categoryEndRow - categoryStartRow + 1;
  const colCount = sheet.getLastColumn();
  const newRowNum = categoryStartRow + paramPosition - 1;
  const newRowValues = Array(colCount).fill(0);
  newRowValues[TIMELINE_CATEGORY_COLUMN_NUMBER - 1] = TIMELINE_CONSTRUCTION_PLAN_CATEGORY;
  newRowValues[TIMELINE_PARAM_NAME_COLUMN_NUMBER - 1] = paramName;

  SOLLibrary.log('Timeline', 'addTimelineParam', '2');

  sheet.insertRowBefore(newRowNum);
  const newRowRange = sheet.getRange(newRowNum, 1, 1, colCount).setValues([newRowValues]);

  if (paramPosition === 1 || paramPosition === categoryParamCount + 1) {
    // if the row was added first or last - make sure all lines are merged for the category cell
    sheet
      .getRange(categoryStartRow, TIMELINE_CATEGORY_COLUMN_NUMBER, categoryParamCount + 1, 1)
      .merge();
  }

  if (paramPosition === 1) {
    // if this is the last param it will be created with a top border so cancel them
    newRowRange.setBorder(true, true, false, null, null, null, '#666666',
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  if (paramPosition === categoryParamCount + 1) {
    // if this is the last param it will be created with a top border so cancel them
    newRowRange.setBorder(false, true, null, null, null, null, '#666666',
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  SOLLibrary.log('Timeline', 'addTimelineParam', '3');

  return newRowNum;
}

function updateTimelineParam(oldName, newName, numberFormat) {
  // set the param name
  _getParamNameRange(oldName).setValue(newName);
  const paramValuesRange = _getParamValuesRange(newName);

  if (numberFormat !== undefined) {
    paramValuesRange.setNumberFormat(numberFormat);
  }
}

function removeTimelineParam(category, paramIndex) {
  SOLLibrary.log('Timeline', 'removeTimelineParam', '1');
  const sheet = _getTimelineSheet();

  const {categoryStartRow} = _getCategoryRowBoundaries(TIMELINE_CONSTRUCTION_PLAN_CATEGORY);
  SOLLibrary.log('Timeline', 'removeTimelineParam', '2');
  const paramRowNum = categoryStartRow + paramIndex;
  sheet.deleteRow(paramRowNum);

  if (paramRowNum === categoryStartRow) {
    SOLLibrary.log('Timeline', 'removeTimelineParam', '3');
    sheet
      .getRange(paramRowNum, 1, 1, sheet.getLastColumn())
      .setBorder(true, true, false, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    SOLLibrary.log('Timeline', 'removeTimelineParam', '4');
    sheet
      .getRange(paramRowNum, TIMELINE_CATEGORY_COLUMN_NUMBER, 1, 1)
      .setValue(category);
  }

  SOLLibrary.log('Timeline', 'removeTimelineParam', '5');
}

function getTimelineParamRowNumber(paramName) {
  SOLLibrary.log('Timeline', 'getTimelineParamRowNumber', '1');
  const sheet = _getTimelineSheet();
  const values = sheet
    .getRange(TIMELINE_HEADER_ROW_NUM + 1,
      TIMELINE_PARAM_NAME_COLUMN_NUMBER,
      sheet.getLastRow() - 1)
    .getValues()
    .flat();


  SOLLibrary.log('Timeline', 'getTimelineParamRowNumber', '2');

  const paramIndex = values.indexOf(paramName);
  if (paramIndex === -1) {
    throw new Error(`Parameter '${paramName}' does not exist in the Timeline sheet`);
  }

  const paramRow = paramIndex + TIMELINE_HEADER_ROW_NUM + 1;

  SOLLibrary.log('Timeline', 'getTimelineParamRowNumber', '3');


  return paramRow;
}

function _getTimelineSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TIMELINE_SHEET_NAME);
}

function _getParamRowRange(paramName) {
  const sheet = _getTimelineSheet();
  return _getParamRange(paramName, 1, sheet.getLastColumn());
}

function _getParamValuesRange(paramName) {
  const sheet = _getTimelineSheet();

  return _getParamRange(
    paramName,
    TIMELINE_FIRST_QUARTER_COLUMN_NUMBER,
    sheet.getLastColumn() - TIMELINE_FIRST_QUARTER_COLUMN_NUMBER + 1);
}

function _getParamNameRange(paramName) {
  return _getParamRange(paramName, TIMELINE_PARAM_NAME_COLUMN_NUMBER, 1);
}

function _getParamRange(paramName, startCol, colCount) {
  const sheet = _getTimelineSheet();
  const rowNum = getTimelineParamRowNumber(paramName);
  return sheet.getRange(rowNum, startCol, 1, colCount);
}

function _getCategoryRowBoundaries(category) {
  const sheet = _getTimelineSheet();
  const cell = sheet.getRange(TIMELINE_HEADER_ROW_NUM + 1, TIMELINE_CATEGORY_COLUMN_NUMBER, sheet.getLastRow(),
    sheet.getLastColumn()).createTextFinder(category).findNext();

  if (!cell) return null;

  const merged = cell.getMergedRanges()[0];

  if (merged) {
    return {
      categoryStartRow: merged.getRow(),
      categoryEndRow: merged.getLastRow()
    };
  }

  return {
    categoryStartRow: cell.getRow(),
    categoryEndRow: cell.getRow()
  };
}

function _getCategoryParams(category) {
  const sheet = _getTimelineSheet();
  const {
    categoryStartRow,
    categoryEndRow
  } = _getCategoryRowBoundaries(category);

  const paramNamesRange = sheet.getRange(
    categoryStartRow,
    TIMELINE_PARAM_NAME_COLUMN_NUMBER,
    categoryEndRow - categoryStartRow + 1,
    1
  );

  return paramNamesRange.getValues().flat();
}

function validateUnitTypesParamsSync(types) {
  const expectedConstructionPlanParams = types.map(type => type + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX);
  const expectedConstructionCostParams = types.map(type => type + TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX);
  const actualConstructionPlanParams = _getCategoryParams(TIMELINE_CONSTRUCTION_PLAN_CATEGORY);
  const actualConstructionCostParams = _getCategoryParams(TIMELINE_CONSTRUCTION_COSTS_CATEGORY);

  _assertArraysMatch(expectedConstructionPlanParams, actualConstructionPlanParams);
  _assertArraysMatch(expectedConstructionCostParams, actualConstructionCostParams);
}

function _assertArraysMatch(expected, actual) {
  SOLLibrary.log('Timeline', '_assertArraysMatch', '1');

  const expectedSet = new Set(expected);
  const actualSet = new Set(actual);

  SOLLibrary.log('Timeline', '_assertArraysMatch', '2');

  const missing = expected.filter(x => !actualSet.has(x));
  const extra = actual.filter(x => !expectedSet.has(x));

  if (missing.length || extra.length) {
    SOLLibrary.log('Timeline', '_assertArraysMatch', '3');
    throw new Error(
      (missing.length ? `Missing: ${missing.join(', ')}\n` : '') +
      (extra.length ? `Extra: ${extra.join(', ')}` : '')
    );
  }
  SOLLibrary.log('Timeline', '_assertArraysMatch', '3');
}