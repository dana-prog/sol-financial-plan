const SYNC_MAPPINGS = [
  {
    source: {
      sheetId: UNITS_SHEET_ID,
      columnHeader: UNITS_UNIT_TYPE_COLUMN_HEADER,
    },
    target: {
      sheetId: CONSTRUCTION_SHEET_ID,
      categories: [
        {
          name: TIMELINE_CONST_PLAN_CATEGORY,
          paramPostfix: TIMELINE_CONST_PLAN_PARAM_POSTFIX,
          numberFormatCallback: getCountNumberFormat
        },
        {
          name: TIMELINE_CONST_COSTS_CATEGORY,
          paramPostfix: TIMELINE_CONST_COST_PARAM_POSTFIX,
        },
        {
          name: TIMELINE_CONST_COMPLETED_UNITS_CATEGORY,
          paramPostfix: TIMELINE_CONST_COMPLETED_UNITS_PARAM_POSTFIX,
          numberFormatCallback: getCountNumberFormat
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
      sheetId: OPERATIONS_SHEET_ID,
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
      sheetId: SETUP_SHEET_ID,
      columnHeader: TIMELINE_CATEGORY_COLUMN_HEADER,
      headerRow: 3
    },
    target: {
      sheetId: OPERATIONS_SHEET_ID,
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

  if (editEvent.range.getNumRows() !== 1 || editEvent.range.getNumColumns() !== 1 ||
      typeof oldValue !== 'string' || typeof newValue !== 'string' || !oldValue.trim() || !newValue.trim()) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Rename one existing label at a time; cleared or pasted labels were not synced.', 'Rename not synced', 8);
    return;
  }
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(syncMapping.target.sheetId);
  const changes = syncMapping.target.categories.map(category => {
    const oldName = oldValue.trim().toLowerCase() + category.paramPostfix;
    const newName = newValue.trim().toLowerCase() + category.paramPostfix;
    const row = _getTimelineParamRowNumber(targetSheet, oldName);
    if (row < 0) throw new Error('Cannot sync rename: missing target ' + oldName);
    const duplicate = _getTimelineParamRowNumber(targetSheet, newName);
    if (duplicate >= 0 && duplicate !== row) throw new Error('Cannot sync duplicate target: ' + newName);
    return {row, newName, numberFormat: category.numberFormatCallback && category.numberFormatCallback(newValue)};
  });
  changes.forEach(change => {
    targetSheet.getRange(change.row, TIMELINE_PARAM_NAME_COLUMN_NUMBER).setValue(change.newName);
    if (change.numberFormat) targetSheet.getRange(change.row, 3, 1, targetSheet.getLastColumn() - 2).setNumberFormat(change.numberFormat);
  });
}

function _findSyncMapping(srcSheetId, srcCol) {
  const mapping = SYNC_MAPPINGS.find(item => item.source.sheetId === srcSheetId);
  if (!mapping) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(srcSheetId);
  const header = typeof srcCol === 'number'
    ? sheet.getRange(mapping.source.headerRow || 1, srcCol).getValue() : srcCol;
  return String(header).trim().toLowerCase() === mapping.source.columnHeader.toLowerCase() ? mapping : null;
}

function _getTimelineParamRowNumber(sheet, paramName) {
  const cell = this._getParamsRange(sheet).createTextFinder(paramName).matchEntireCell(true).useRegularExpression(false).findNext();
  return cell ? cell.getRow() : -1;
}

function _getParamsRange(sheet) {
  return sheet.getRange(
    TIMELINE_HEADER_ROW_NUM + 1,
    TIMELINE_PARAM_NAME_COLUMN_NUMBER,
    sheet.getLastRow() - TIMELINE_HEADER_ROW_NUM
  );
}
