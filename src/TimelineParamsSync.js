const SYNC_MAPPINGS = [
  {
    source: {
      sheetId: UNITS_SHEET_ID,
      columnHeader: UNITS_UNIT_TYPE_COLUMN_HEADER,
    },
    target: {
      sheetId: CONSTRUCTION_TIMELINE_SHEET_ID,
      categories: [
        {
          name: TIMELINE_UNITS_COUNT_CATEGORY,
          paramPostfix: TIMELINE_UNITS_COUNT_PARAM_POSTFIX,
          numberFormatCallback: getCountNumberFormat
        },
        {
          name: TIMELINE_UNITS_COSTS_CATEGORY,
          paramPostfix: TIMELINE_UNITS_COST_PARAM_POSTFIX,
        }
      ]
    }
  },
  {
    source: {
      sheetId: STAFF_SHEET_ID,
      columnHeader: STAFF_ROLE_COLUMN_HEADER,
    },
    target: {
      sheetId: MAIN_TIMELINE_SHEET_ID,
      categories: [
        {
          name: TIMELINE_STAFF_CATEGORY,
          paramPostfix: TIMELINE_STAFF_PARAM_POSTFIX,
          numberFormatCallback: getCountNumberFormat
        },
        {
          name: TIMELINE_MONTHLY_NET_SALARIES_CATEGORY,
          paramPostfix: TIMELINE_MONTHLY_NET_SALARIES_PARAM_POSTFIX,
        },
        {
          name: TIMELINE_NET_SALARIES_CATEGORY,
          paramPostfix: TIMELINE_NET_SALARIES_PARAM_POSTFIX,
        }
      ]
    }
  },
  {
    source: {
      sheetId: INFRASTRUCTURE_TIMELINE_SHEET_ID,
      columnHeader: TIMELINE_CATEGORY_COLUMN_HEADER
    },
    target: {
      sheetId: MAIN_TIMELINE_SHEET_ID,
      categories: [
        {
          name: TIMELINE_INFRASTRUCTURE_COSTS_CATEGORY,
          paramPostfix: TIMELINE_INFRASTRUCTURE_COSTS_PARAM_POSTFIX,
        },
      ]
    }
  }
];

function updateTimelineParamNames(editEvent) {
  const srcSheetId = editEvent.range.getSheet().getSheetId();
  const srcCol = editEvent.range.getColumn();
  const oldValue = editEvent.oldValue;
  const newValue = editEvent.value;
  const syncMapping = this._findSyncMapping(srcSheetId, srcCol);
  SOLLibrary.logArgs('TimelineParamsSync', 'updateTimelineParamNames', {
    srcSheetId,
    srcCol,
    oldValue,
    newValue,
    syncMapping
  });
  if (!syncMapping) {
    return;
  }

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(syncMapping.target.sheetId);

  SOLLibrary.debugDuration('TimelineParamsSync.updateTimelineParamNames - Object.values iteration', () => {
    // for each target category, update the param name in the target sheet
    Object.values(syncMapping.target.categories).forEach(category => {
      this._updateTimelineParamName(
        targetSheet,
        oldValue + category.paramPostfix,
        newValue + category.paramPostfix,
        category.numberFormatCallback && category.numberFormatCallback(newValue)
      );
    });
  });
}

function _findSyncMapping(srcSheetId, srcCol) {
  const srcColHeader = (typeof srcCol === 'number') ? getColumnHeaderByNum(srcSheetId, srcCol) : srcCol;
  return SYNC_MAPPINGS.find(mapping => {
    return mapping.source.sheetId === srcSheetId &&
      mapping.source.columnHeader === srcColHeader;
  });
}

function _updateTimelineParamName(sheet, oldName, newName, numberFormat) {
  SOLLibrary.debugDuration('TimelineParamsSync._updateTimelineParamName', () => {
    let paramRowNum;
    paramRowNum = this._getTimelineParamRowNumber(sheet, oldName);

    if (paramRowNum === -1) {
      SOLLibrary.log('TimelineParamsSync', '_updateTimelineParam', `param '${oldName}' does not exist`, 'WARN');
      return;
    }

    SOLLibrary.logArgs('TimelineParamsSync', '_updateTimelineParamName', {
      oldName,
      newName,
      paramRowNum
    });

    let paramNameRange;
    paramNameRange = sheet.getRange(paramRowNum, TIMELINE_PARAM_NAME_COLUMN_NUMBER);

    paramNameRange.setValue(newName);

    if (numberFormat) {
      let paramRange;
      paramRange = sheet.getRange(`${paramRowNum}:${paramRowNum}`);
      paramRange.setNumberFormat(numberFormat);
    }
  });
}

function _getTimelineParamRowNumber(sheet, paramName) {
  const cell = this._getParamsRange(sheet).createTextFinder(paramName).findNext();
  return cell ? cell.getRow() : -1;
}

function _getParamsRange(sheet) {
  return sheet.getRange(
    TIMELINE_HEADER_ROW_NUM + 1,
    TIMELINE_PARAM_NAME_COLUMN_NUMBER,
    sheet.getLastRow() - TIMELINE_HEADER_ROW_NUM - 1
  );
}

