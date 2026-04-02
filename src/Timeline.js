const MAIN_TIMELINE_SHEET_ID = 224444304;
const CONSTRUCTION_TIMELINE_SHEET_ID = 1618658099;
const INFRASTRUCTURE_TIMELINE_SHEET_ID = 374017605;
const UNITS_SHEET_ID = 1436796628;
const UNITS_UNIT_TYPE_COLUMN_HEADER = 'unit type';

const TIMELINE_CATEGORY_COLUMN_NUMBER = 1;
const TIMELINE_PARAM_NAME_COLUMN_NUMBER = 2;
const TIMELINE_CATEGORY_COLUMN_HEADER = 'category';
const TIMELINE_FIRST_QUARTER_COLUMN_NUMBER = 3;
const TIMELINE_HEADER_ROW_NUM = 3;

// categories
const TIMELINE_UNITS_COUNT_CATEGORY = 'Units Count';
const TIMELINE_UNITS_COUNT_PARAM_POSTFIX = ' construction count';

const TIMELINE_UNITS_COSTS_CATEGORY = 'Units Costs';
const TIMELINE_UNITS_COST_PARAM_POSTFIX = ' construction cost';

const TIMELINE_STAFF_CATEGORY = 'Staff';
const TIMELINE_STAFF_PARAM_POSTFIX = 's headcount';

const TIMELINE_MONTHLY_NET_SALARIES_CATEGORY = 'Monthly Net Salaries';
const TIMELINE_MONTHLY_NET_SALARIES_PARAM_POSTFIX = 's monthly net salaries';

const TIMELINE_NET_SALARIES_CATEGORY = 'Net Salaries';
const TIMELINE_NET_SALARIES_PARAM_POSTFIX = 's net salaries';

const INFRASTRUCTURE_COSTS_CATEGORY = 'Infrastructure Costs';
const INFRASTRUCTURE_COSTS_PARAM_POSTFIX = '';

let _timelineSheet = null;


function syncTimelineConstructionParams(unitTypes) {
  SOLLibrary.debugDuration('syncTimelineConstructionParams', () => {
    _syncTimelineParams(
      TIMELINE_UNITS_COUNT_CATEGORY,
      TIMELINE_UNITS_COUNT_PARAM_POSTFIX,
      unitTypes,
      getCountNumberFormat,
    );

    _syncTimelineParams(
      TIMELINE_UNITS_COSTS_CATEGORY,
      TIMELINE_UNITS_COST_PARAM_POSTFIX,
      unitTypes
    );
  });
}

function syncTimelineStaffParams(staffRoles) {
  SOLLibrary.debugDuration(
    'syncTimelineStaffParams',
    () => {
      _syncTimelineParams(
        TIMELINE_STAFF_CATEGORY,
        TIMELINE_STAFF_PARAM_POSTFIX,
        staffRoles,
        getCountNumberFormat,
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
    });
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
  SOLLibrary.debugDuration('_syncCategoryRowCount', () => {

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
    const range = sheet.getRange(categoryStartRow, 1, paramCountDiff > 0 ? paramCountDiff : 1, sheet.getLastColumn());
    range.setBorder(
      categoryStartRow > TIMELINE_HEADER_ROW_NUM + 1,
      true,
      expectedParamCount === 1,
      true,
      null,
      null,
      '#666666',
      SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  });
}

function _fillCategoryNewValues(category, newParamNames) {
  SOLLibrary.debugDuration('_fillCategoryNewValues', () => {
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
  });
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
  let res;
  SOLLibrary.debugDuration('_getCategoryParamNames', () => {

    const sheet = getTimelineSheet();
    const {
      categoryStartRow,
      categoryEndRow
    } = _getCategoryRowBoundaries(category);

    res = sheet
      .getRange(categoryStartRow, TIMELINE_PARAM_NAME_COLUMN_NUMBER, categoryEndRow - categoryStartRow + 1, 1)
      .getValues()
      .flat();
  });

  return res;
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

function getTimelineSheet() {
  if (!_timelineSheet) {
    _timelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(MAIN_TIMELINE_SHEET_ID);
  }

  return _timelineSheet;
}

// function _getUnitTypeAbbreviation(unitType) {
//   return unitType
//     .replace(/ area$/, '')
//     .replace(/ and$/, '')
//     .replace('surroundings', 'surr')
//     .replace(/\(\d+\)$/, '').trim();
// }

// function getTimelineTotalUnitCount(unitType) {
//   const paramName = unitType + TIMELINE_CONSTRUCTION_PLAN_PARAM_POSTFIX;
//   const paramValuesRange = _getParamValuesRange(paramName);
//   return paramValuesRange
//     .getValues()[0]
//     .reduce((acc, val) => acc + val, 0);
// }

// function _getParamValuesRange(paramName) {
//   const sheet = getTimelineSheet();
//
//   return _getParamRange(
//     paramName,
//     TIMELINE_FIRST_QUARTER_COLUMN_NUMBER,
//     sheet.getLastColumn() - TIMELINE_FIRST_QUARTER_COLUMN_NUMBER + 1);
// }
//
// function _getParamRange(paramName, startCol, colCount) {
//   const sheet = getTimelineSheet();
//   const rowNum = _getTimelineParamRowNumber(paramName);
//   return sheet.getRange(rowNum, startCol, 1, colCount);
// }
