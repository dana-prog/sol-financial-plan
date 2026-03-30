const TIMELINE_SHEET_NAME = 'Timeline';
const TIMELINE_SHEET_ID = 224444304;

//TODO: remove hard-coded column/row numbers
const TIMELINE_CATEGORY_COLUMN_NUMBER = 1;
const TIMELINE_PARAM_NAME_COLUMN_NUMBER = 2;
const TIMELINE_FIRST_QUARTER_COLUMN_NUMBER = 3;
const TIMELINE_HEADER_ROW_NUM = 3;

// categories
const TIMELINE_CONSTRUCTION_PLAN_CATEGORY = 'Construction Plan';
const TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX = ' construction count';

const TIMELINE_CONSTRUCTION_COSTS_CATEGORY = 'Construction Costs';
const TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX = ' construction cost';

const TIMELINE_STAFF_CATEGORY = 'Staff';
const TIMELINE_STAFF_PARAM_POSTFIX = 's headcount';

const TIMELINE_MONTHLY_NET_SALARIES_CATEGORY = 'Monthly Net Salaries';
const TIMELINE_MONTHLY_NET_SALARIES_PARAM_POSTFIX = 's monthly net salaries';

const TIMELINE_NET_SALARIES_CATEGORY = 'Net Salaries';
const TIMELINE_NET_SALARIES_PARAM_POSTFIX = 's net salaries';

function updateTimelineConstructionParams(oldUnitType, newUnitType) {
  // Construction Plan Param
  _updateTimelineParam(
    oldUnitType + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX,
    newUnitType + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX,
    `[=1]0 "${newUnitType}";0 "${newUnitType}s"`
  );

  // Construction Costs Param
  _updateTimelineParam(
    oldUnitType + TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX,
    newUnitType + TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX
  );
}

function updateTimelineStaffParams(oldRoleName, newRoleName) {
  // Staff Param
  _updateTimelineParam(
    oldRoleName + TIMELINE_STAFF_PARAM_POSTFIX,
    newRoleName + TIMELINE_STAFF_PARAM_POSTFIX,
    `[=1]0 "${newRoleName}";0 "${newRoleName}s"`
  );

  // Monthly Net Salaries Param
  _updateTimelineParam(
    oldRoleName + TIMELINE_MONTHLY_NET_SALARIES_PARAM_POSTFIX,
    newRoleName + TIMELINE_MONTHLY_NET_SALARIES_PARAM_POSTFIX
  );

  // Net Salaries Param
  _updateTimelineParam(
    oldRoleName + TIMELINE_NET_SALARIES_PARAM_POSTFIX,
    newRoleName + TIMELINE_NET_SALARIES_PARAM_POSTFIX
  );
}

function syncTimelineConstructionParams(unitTypes) {
  _syncTimelineParams(
    TIMELINE_CONSTRUCTION_PLAN_CATEGORY,
    TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX,
    unitTypes,
    _getCountNumberFormat,
  );

  _syncTimelineParams(
    TIMELINE_CONSTRUCTION_COSTS_CATEGORY,
    TIMELINE_CONSTRUCTION_COST_PARAM_POSTFIX,
    unitTypes
  );

  SOLLibrary.alert('Done',
    `'${TIMELINE_CONSTRUCTION_PLAN_CATEGORY}' params and '${TIMELINE_CONSTRUCTION_COSTS_CATEGORY}' params in the 'Timeline' sheet were synced according to the unit types in the 'Construction Costs' sheet`)
}

function syncTimelineStaffParams(staffRoles) {
  _syncTimelineParams(
    TIMELINE_STAFF_CATEGORY,
    TIMELINE_STAFF_PARAM_POSTFIX,
    staffRoles,
    _getCountNumberFormat,
  );

  _syncTimelineParams(
    TIMELINE_MONTHLY_NET_SALARIES_CATEGORY,
    TIMELINE_MONTHLY_NET_SALARIES_PARAM_POSTFIX,
    staffRoles
  );

  _syncTimelineParams(
    TIMELINE_NET_SALARIES_CATEGORY,
    TIMELINE_NET_SALARIES_PARAM_POSTFIX,
    staffRoles
  );

  SOLLibrary.alert('Done',
    `'${TIMELINE_CONSTRUCTION_PLAN_CATEGORY}' params and '${TIMELINE_CONSTRUCTION_COSTS_CATEGORY}' params in the 'Timeline' sheet were synced according to the unit types in the 'Construction Costs' sheet`)
}

function getTimelineTotalUnitCount(unitType) {
  const paramName = unitType + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX;
  const paramValuesRange = _getParamValuesRange(paramName);
  return paramValuesRange
    .getValues()[0]
    .reduce((acc, val) => acc + val, 0);
}

function _updateTimelineParam(oldName, newName, numberFormat) {
  SOLLibrary.logArgs('Timeline', '_updateTimelineParam', {
    oldName,
    newName,
    numberFormat
  }, 'start');

  const sheet = getTimelineSheet();
  const paramRowNum = _getTimelineParamRowNumber(oldName);

  const paramNameRange = sheet.getRange(paramRowNum, TIMELINE_PARAM_NAME_COLUMN_NUMBER);
  paramNameRange.setValue(newName);

  if (numberFormat !== undefined) {
    const paramRange = sheet.getRange(`${paramRowNum}:${paramRowNum}`);
    paramRange.setNumberFormat(numberFormat);
  }
  SOLLibrary.logArgs('Timeline', '_updateTimelineParam', {
    oldName,
    newName,
    numberFormat
  }, 'end');
}

function _getParamValuesRange(paramName) {
  const sheet = getTimelineSheet();

  return _getParamRange(
    paramName,
    TIMELINE_FIRST_QUARTER_COLUMN_NUMBER,
    sheet.getLastColumn() - TIMELINE_FIRST_QUARTER_COLUMN_NUMBER + 1);
}

function _getParamRange(paramName, startCol, colCount) {
  const sheet = getTimelineSheet();
  const rowNum = _getTimelineParamRowNumber(paramName);
  return sheet.getRange(rowNum, startCol, 1, colCount);
}

function _getTimelineParamRowNumber(paramName) {
  const sheet = getTimelineSheet();
  const values = sheet
    .getRange(TIMELINE_HEADER_ROW_NUM + 1,
      TIMELINE_PARAM_NAME_COLUMN_NUMBER,
      sheet.getLastRow() - 1)
    .getValues()
    .flat();

  const paramIndex = values.indexOf(paramName);
  if (paramIndex === -1) {
    throw new Error(`Parameter '${paramName}' does not exist in the Timeline sheet`);
  }

  return paramIndex + TIMELINE_HEADER_ROW_NUM + 1;
}

function _syncTimelineParams(category, paramPostfix, baseValues, getValueFormat) {
  const oldParamNames = _getCategoryParamNames(category);
  const newParamNames = baseValues.map(baseValue => baseValue + paramPostfix);

  if (oldParamNames.join() === newParamNames.join()) {
    return;
  }

  _syncCategoryRowCount(category, newParamNames.length);
  _fillCategoryNewValues(category, newParamNames);
  getValueFormat && _setValuesNumberFormat(category, baseValues, getValueFormat);
}

function _syncCategoryRowCount(category, expectedParamCount) {
  const sheet = getTimelineSheet();
  const {
    categoryStartRow,
    categoryEndRow
  } = _getCategoryRowBoundaries(category);
  const actualParamCount = categoryEndRow - categoryStartRow + 1;

  if (actualParamCount === expectedParamCount) {
    return;
  }

  const paramCountDiff = expectedParamCount - actualParamCount;
  if (paramCountDiff > 0) {
    sheet.insertRows(categoryStartRow, paramCountDiff);
    // merge the inserted rows with the existing category rows in the category column
    sheet
      .getRange(categoryStartRow, TIMELINE_CATEGORY_COLUMN_NUMBER, expectedParamCount, 1)
      .merge();

  } else if (paramCountDiff < 0) {
    sheet.deleteRows(categoryStartRow, Math.abs(paramCountDiff));
    sheet.getRange(categoryStartRow, TIMELINE_CATEGORY_COLUMN_NUMBER, 1, 1).setValue(category);
  }

  // set the border of the category
  // (since we might have added a row at the beginning of the category without a border or removed the last row which had a border)
  const startRowRange = sheet.getRange(categoryStartRow, 1, 1, sheet.getLastColumn());
  startRowRange.setBorder(categoryStartRow > TIMELINE_HEADER_ROW_NUM + 1, true, expectedParamCount === 1, true, null,
    null, '#666666', SpreadsheetApp.BorderStyle.SOLID_THICK);

  // const endRowRange = sheet.getRange(
  //   categoryStartRow + totalRowsAfterSync - 1,
  //   1,
  //   1,
  //   sheet.getLastColumn()
  // );
  // endRowRange.setBorder(
  //   totalRowsAfterSync === 1,
  //   true,
  //   true,
  //   true,
  //   null,
  //   null,
  //   '#666666',
  //   SpreadsheetApp.BorderStyle.SOLID_THICK
  // );
}

function _fillCategoryNewValues(category, newParamNames) {
  const oldParamsQuarterValues = _getCategoryParamsQuarterValues(category);
  const sheet = getTimelineSheet();
  const numOfTimelineQuarters = sheet.getLastColumn() - TIMELINE_PARAM_NAME_COLUMN_NUMBER;

  const newCategoryRange = _getCategoryRange(category);

  const newValues = newParamNames.map((name) => {
    const oldParamValues = oldParamsQuarterValues[name];
    const paramTimelineValues = oldParamValues || Array(numOfTimelineQuarters).fill(0);
    newCategoryRange.clearContent();
    return [category, name, ...paramTimelineValues];
  });

  newCategoryRange.setValues(newValues);
}

function _setValuesNumberFormat(category, unitTypes, getValueFormat) {
  const sheet = getTimelineSheet();
  const numOfTimelineQuarters = sheet.getLastColumn() - TIMELINE_PARAM_NAME_COLUMN_NUMBER;
  const formats = [];

  for (let i = 0; i < unitTypes.length; i++) {
    const unitType = unitTypes[i];
    const format = getValueFormat(unitType);
    formats.push(Array(numOfTimelineQuarters).fill(format));
  }

  const {
    categoryStartRow,
    categoryEndRow
  } = _getCategoryRowBoundaries(category);
  const range = sheet.getRange(categoryStartRow, TIMELINE_FIRST_QUARTER_COLUMN_NUMBER,
    categoryEndRow - categoryStartRow + 1, numOfTimelineQuarters);
  range.setNumberFormats(formats);
}

function _getCategoryRowBoundaries(category) {
  const sheet = getTimelineSheet();
  const cell = sheet.getRange(TIMELINE_HEADER_ROW_NUM + 1, TIMELINE_CATEGORY_COLUMN_NUMBER, sheet.getLastRow(),
    sheet.getLastColumn())
    .createTextFinder(category).findNext();

  if (!cell) {
    throw new Error(`Category '${category}' does not exist in the Timeline sheet`);
  }

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

function _getCategoryRange(category) {
  const sheet = getTimelineSheet();
  const {
    categoryStartRow,
    categoryEndRow
  } = _getCategoryRowBoundaries(category);

  return sheet.getRange(categoryStartRow, 1, categoryEndRow - categoryStartRow + 1, sheet.getLastColumn());
}

function _getCategoryParamNames(category) {
  const sheet = getTimelineSheet();
  const {
    categoryStartRow,
    categoryEndRow
  } = _getCategoryRowBoundaries(category);

  return sheet
    .getRange(categoryStartRow, TIMELINE_PARAM_NAME_COLUMN_NUMBER, categoryEndRow - categoryStartRow + 1, 1)
    .getValues()
    .flat();
}

function _getCategoryParamsQuarterValues(category) {
  const sheet = getTimelineSheet();
  const numOfTimelineQuarters = sheet.getLastColumn() - TIMELINE_PARAM_NAME_COLUMN_NUMBER;
  const categoryRange = _getCategoryRange(category);
  const oldValues = categoryRange.getValues();
  // assumption: either all cells of a category contain the same formula or none contains any formula
  // take the formula from the 1st cell in the 1st row
  const oldFormulaValue = categoryRange.getFormulas()[0][TIMELINE_FIRST_QUARTER_COLUMN_NUMBER];

  const res = {};

  oldValues.forEach((rowValues) => {
    const paramName = rowValues[TIMELINE_PARAM_NAME_COLUMN_NUMBER - 1];

    res[paramName] = oldFormulaValue
      ? Array(numOfTimelineQuarters).fill(oldFormulaValue)
      : rowValues.slice(TIMELINE_PARAM_NAME_COLUMN_NUMBER);
  });

  return res;
}

function _getCountNumberFormat(baseValue) {
  const abbreviation = _getUnitTypeAbbreviation(baseValue);
  return `[=1]0 "${abbreviation}";0 "${abbreviation}s"`;
}

function _getUnitTypeAbbreviation(unitType) {
  return unitType
    .replace(/ area$/, '')
    .replace(/ and$/, '')
    .replace('surroundings', 'surr')
    .replace(/\(\d+\)$/, '').trim();
}

function getTimelineSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetById(TIMELINE_SHEET_ID);
}
