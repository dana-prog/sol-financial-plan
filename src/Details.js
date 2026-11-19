const DETAILS_SHEET_ID = 2109063306;
const DETAILS_PERIODS = ['Month', 'Quarter', 'Year'];
const DETAILS_ALL_PERIOD_LABELS = {
  Month: 'All Month',
  Quarter: 'All Quarters',
  Year: 'All Years',
};

const params = [
  "construction expenses",
  "infrastructure expenses",
  "setup expenses",
  "direct operating expenses",
  "other operating expenses",
  "land expenses",
  "borrowing repayments",
  "loan disbursements"
];

function showDetails(param, month) {
  const aliases = {
    'planning expenses': 'setup expenses',
    'infra expenses': 'infrastructure expenses',
    'borrowings repaid': 'borrowing repayments',
    'loan disbursed': 'loan disbursements',
  };
  param = aliases[param] || param;
  if (!(month instanceof Date) || !Number.isFinite(month.getTime())) {
    SpreadsheetApp.getUi().alert('Select a monthly value in a supported expense row.');
    return;
  }
  if (!params.find((value) => param === value)) {
    SpreadsheetApp.getUi().alert(
      `"${param}" does not have a details view.`
    );
    return;
  }

  const detailsSheet = SpreadsheetApp.getActive().getSheetById(DETAILS_SHEET_ID);

  detailsSheet.getRange('A2').setValue('Month');
  detailsSheet.getRange('D2').setValue(SOLLibrary.capitalize(param));
  _updateDetailsPeriodSelection(detailsSheet, _detailsPeriodLabel(month, 'Month'));
  _refreshDetailsView(detailsSheet);

  detailsSheet.activate();
}

function updateDetailsPeriod(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getSheetId() !== DETAILS_SHEET_ID ||
      range.getRow() > 2 || range.getLastRow() < 2 || range.getColumn() > 1) return;
  _updateDetailsPeriodSelection(sheet);
}

function _detailsPeriodLabel(month, period) {
  const year = month.getFullYear();
  if (period === 'Year') return String(year);
  if (period === 'Quarter') return `Q${Math.floor(month.getMonth() / 3) + 1} ${year}`;
  const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${names[month.getMonth()]} ${year}`;
}

function updateDetailsView(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getSheetId() === DETAILS_SHEET_ID) {
    if (range.getRow() > 2 || range.getLastRow() < 2 ||
        ![1, 3, 4].some(column => range.getColumn() <= column && range.getLastColumn() >= column)) return;
    _refreshDetailsView(sheet);
    return;
  }
  // Source edits can change the position of construction rows in the All view.
  const details = sheet.getParent().getSheetById(DETAILS_SHEET_ID);
  if (details) _refreshDetailsView(details);
}

function _refreshDetailsView(sheet) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    for (let attempt = 0; attempt < 4; attempt++) {
      try {
        // Rebuild the groups after recalculation; old merges can span new types.
        sheet.getRange(5, 1, sheet.getMaxRows() - 4, 1).breakApart();
        _resizeDetailsRows(sheet);
        _refreshDetailsUnitFormats(sheet);
        _mergeDetailsTypes(sheet);
        return;
      } catch (error) {
        // Sheets can briefly reject reads while expanding a recalculating array.
        if (attempt === 3 || !String(error.message).includes('Service Spreadsheets failed')) throw error;
        Utilities.sleep(500);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function _detailsTypeGroups(types) {
  const groups = [];
  for (let start = 0; start < types.length;) {
    let end = start + 1;
    while (end < types.length && types[end] === types[start]) end++;
    if (types[start] && types[start] !== 'Total' && end - start > 1) {
      groups.push({ row: start + 5, count: end - start });
    }
    start = end;
  }
  return groups;
}

function _mergeDetailsTypes(sheet) {
  const types = sheet.getRange(5, 1, sheet.getMaxRows() - 4, 1).getValues().flat();
  for (const group of _detailsTypeGroups(types)) {
    sheet.getRange(group.row, 1, group.count, 1).merge().setVerticalAlignment('middle');
  }
}

function _resizeDetailsRows(sheet) {
  const formula = sheet.getRange('A5').getFormula();
  if (!formula) throw new Error('The Details table formula is missing.');
  // Count the formula result independently so a larger selection can recover
  // even when its array cannot yet spill into the trimmed sheet.
  const counter = sheet.getRange('E1');
  const countFormula = `=ROWS(${formula.slice(1)})`;
  if (counter.getFormula() !== countFormula) counter.setFormula(countFormula);
  for (let attempt = 0; attempt < 10; attempt++) {
    SpreadsheetApp.flush();
    const count = counter.getValue();
    if (Number.isInteger(count) && count > 0) {
      const lastRow = count + 4;
      if (lastRow > sheet.getMaxRows()) _growDetailsRows(sheet, lastRow);
      SpreadsheetApp.flush();
      if (counter.getValue() === count && sheet.getRange(lastRow, 1).getValue() === 'Total') {
        const extraRows = sheet.getMaxRows() - lastRow;
        if (extraRows > 0) {
          if (!sheet.getRange(lastRow + 1, 1, extraRows, sheet.getMaxColumns()).isBlank()) {
            throw new Error('Cannot remove Details rows containing content below Total.');
          }
          sheet.deleteRows(lastRow + 1, extraRows);
        }
        return;
      }
    }
    // flush submits writes, but dependent array formulas may still be recalculating.
    Utilities.sleep(250);
  }
  throw new Error('The Details table has not finished calculating its Total row.');
}

function _growDetailsRows(sheet, lastRow) {
  const currentRows = sheet.getMaxRows();
  sheet.insertRowsAfter(currentRows, lastRow - currentRows);
  sheet.getRange(5, 1, 1, 5).copyFormatToRange(sheet, 1, 5, currentRows + 1, lastRow);
  sheet.setRowHeights(currentRows + 1, lastRow - currentRows, sheet.getRowHeight(5));
  const rules = sheet.getConditionalFormatRules().map(rule => rule.copy().setRanges(
    rule.getRanges().map(range => range.getRow() >= 5 && range.getLastRow() === currentRows
      ? sheet.getRange(range.getRow(), range.getColumn(), lastRow - range.getRow() + 1, range.getNumColumns())
      : range)
  ).build());
  sheet.setConditionalFormatRules(rules);
}

function _detailsUnitFormats(rows, params, formats) {
  const byParam = new Map(params.map((row, index) => [String(row[0]).toLowerCase(), formats[index][0]]));
  return rows.map(row => {
    if (row[0] !== 'Construction Expenses') return ['General'];
    const param = String(row[4]).toLowerCase().replace(/ const\. cost$/, ' const. count');
    const format = byParam.get(param);
    if (!format) throw new Error(`Missing planned-unit number format for ${param}.`);
    return [format.replace(/complex room/g, 'complex home')];
  });
}

function _refreshDetailsUnitFormats(sheet) {
  const rowCount = sheet.getMaxRows() - 4;
  if (rowCount <= 0) return;
  SpreadsheetApp.flush();
  const rows = sheet.getRange(5, 1, rowCount, 5).getValues();
  const target = sheet.getRange(5, 3, rowCount, 1);
  if (!rows.some(row => row[0] === 'Construction Expenses')) {
    target.setNumberFormat('General');
    return;
  }
  const source = sheet.getParent().getRangeByName('S_CONSTRUCTION_EXP_P');
  if (!source) throw new Error('The construction timeline named range is missing.');
  const sourceSheet = source.getSheet();
  const headers = sourceSheet.getRange(source.getRow(), source.getColumn(), 1, source.getNumColumns()).getValues()[0];
  const paramColumn = headers.findIndex(value => String(value).toLowerCase() === 'param');
  const monthColumn = headers.findIndex(value => value instanceof Date);
  if (paramColumn < 0 || monthColumn < 0) throw new Error('Construction timeline headers are incomplete.');
  const firstRow = source.getRow() + 1;
  const sourceRows = source.getNumRows() - 1;
  const params = sourceSheet.getRange(firstRow, source.getColumn() + paramColumn, sourceRows, 1).getValues();
  const formats = sourceSheet.getRange(firstRow, source.getColumn() + monthColumn, sourceRows, 1).getNumberFormats();
  target.setNumberFormats(_detailsUnitFormats(rows, params, formats));
}

function _detailsPeriodSelection(months, period, selected) {
  if (!DETAILS_PERIODS.includes(period)) throw new Error(`Unsupported details period: ${period}`);
  const dates = months.filter(month => month instanceof Date && Number.isFinite(month.getTime()));
  if (!dates.length) throw new Error('The project timeline has no months to select.');
  const allPeriods = DETAILS_ALL_PERIOD_LABELS[period];
  const options = [allPeriods, ...new Set(dates.map(month => _detailsPeriodLabel(month, period)))];
  if (Object.values(DETAILS_ALL_PERIOD_LABELS).includes(selected)) {
    return { options, value: allPeriods };
  }
  // Keep the selected date within the same calendar period when changing granularity.
  const anchor = dates.find(month => DETAILS_PERIODS.some(
    value => _detailsPeriodLabel(month, value) === String(selected)
  )) || dates[0];
  return { options, value: _detailsPeriodLabel(anchor, period) };
}

function _updateDetailsPeriodSelection(sheet, selectedValue) {
  const period = sheet.getRange('A2').getValue();
  const selected = sheet.getRange('C2');
  const months = sheet.getParent().getRangeByName('R_MONTHS');
  if (!months) throw new Error('The project timeline named range R_MONTHS is missing.');
  const selection = _detailsPeriodSelection(months.getValues().flat(), period,
    selectedValue === undefined ? selected.getValue() : selectedValue);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(selection.options, true)
    .setAllowInvalid(false)
    .build();
  selected.setDataValidation(validation).setNumberFormat('@').setValue(selection.value);
  sheet.getRange('C1').setValue(`Select ${period}`);
}
