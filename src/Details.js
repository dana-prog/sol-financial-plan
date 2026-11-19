const DETAILS_SHEET_ID = 2109063306;
const DETAILS_PERIOD_VALUE_FORMULA = `=LET(
  months, FILTER(R_MONTHS, ISNUMBER(R_MONTHS)),
  saved, TO_TEXT(__SELECTION__),
  period, A2,
  month_labels, ARRAYFORMULA(TEXT(months, "mmm yyyy")),
  quarter_labels, ARRAYFORMULA("Q"&ROUNDUP(MONTH(months)/3)&" "&YEAR(months)),
  year_labels, ARRAYFORMULA(TEXT(months, "yyyy")),
  all_label, SWITCH(period, "Month", "All Month", "Quarter", "All Quarters", "Year", "All Years"),
  anchor, IFNA(MATCH(TRUE, ARRAYFORMULA(((month_labels=saved)+(quarter_labels=saved)+(year_labels=saved))>0), 0), 1),
  IF(OR(saved="All Month", saved="All Quarters", saved="All Years"), all_label,
    SWITCH(period, "Month", INDEX(month_labels, 1, anchor), "Quarter", INDEX(quarter_labels, 1, anchor), "Year", INDEX(year_labels, 1, anchor)))
)`;

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
  _updateDetailsPeriodOptions(detailsSheet);
  _refreshDetailsView(detailsSheet);

  detailsSheet.activate();
}

function updateDetailsPeriod(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getSheetId() !== DETAILS_SHEET_ID ||
      range.getRow() > 2 || range.getLastRow() < 2 ||
      ![1, 3].some(column => range.getColumn() <= column && range.getLastColumn() >= column)) return;
  const selected = sheet.getRange('C2');
  // Choosing a dropdown item replaces its formula. Remember that choice and
  // restore the formula so the next period change needs no script round trip.
  if (!selected.getFormula()) _updateDetailsPeriodSelection(sheet, selected.getValue());
  if (range.getColumn() === 1) _updateDetailsPeriodOptions(sheet);
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
  if (details) {
    _updateDetailsPeriodOptions(details);
    _refreshDetailsView(details);
  }
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
  // Sheets expands the array itself; wait for its Total before trimming blanks.
  for (let attempt = 0; attempt < 10; attempt++) {
    SpreadsheetApp.flush();
    const currentRows = sheet.getMaxRows();
    const types = sheet.getRange(5, 1, currentRows - 4, 1).getValues().flat();
    const totalIndex = types.indexOf('Total');
    if (totalIndex >= 0) {
      const lastRow = totalIndex + 5;
      const extraRows = currentRows - lastRow;
      if (extraRows > 0) {
        if (!sheet.getRange(lastRow + 1, 1, extraRows, sheet.getMaxColumns()).isBlank()) {
          throw new Error('Cannot remove Details rows containing content below Total.');
        }
        sheet.deleteRows(lastRow + 1, extraRows);
      }
      return;
    }
    Utilities.sleep(250);
  }
  throw new Error('The Details table has not finished calculating its Total row.');
}

function _detailsUnitFormats(rows, source, formats) {
  const headers = source[0];
  const paramColumn = headers.findIndex(value => String(value).toLowerCase() === 'param');
  const categoryColumn = headers.findIndex(value => String(value).toLowerCase() === '_category');
  if (paramColumn < 0 || categoryColumn < 0) throw new Error('Construction timeline headers are incomplete.');
  const byParam = new Map(source.map((row, index) => [String(row[paramColumn]).toLowerCase(), { row, format: formats[index][0] }]));
  const costs = source.slice(1).filter(row => String(row[categoryColumn]).toLowerCase() === 'construction expenses')
    .sort((a, b) => String(a[paramColumn]).localeCompare(String(b[paramColumn])));
  const monthColumns = new Map(headers.flatMap((value, index) => value instanceof Date ? [[value.getTime(), index]] : []));
  const nextByMonth = new Map();
  return rows.map(row => {
    if (row[0] !== 'Construction Expenses') return ['General'];
    const month = row[1] instanceof Date ? row[1].getTime() : NaN;
    const column = monthColumns.get(month);
    if (column === undefined) throw new Error('Construction expense month is outside the source timeline.');
    // The formula orders construction rows by month and param. Match that order
    // against the source amounts/counts without adding a visible identifier column.
    for (let index = nextByMonth.get(month) || 0; index < costs.length; index++) {
      const cost = costs[index];
      const param = String(cost[paramColumn]).toLowerCase().replace(/ const\. cost$/, ' const. count');
      const units = byParam.get(param);
      if (cost[column] !== row[3] || !units || units.row[column] !== row[2]) continue;
      if (!units.format) throw new Error(`Missing planned-unit number format for ${param}.`);
      nextByMonth.set(month, index + 1);
      return [units.format.replace(/complex room/g, 'complex home')];
    }
    throw new Error('Construction expense could not be matched to its planned-unit format.');
  });
}

function _refreshDetailsUnitFormats(sheet) {
  const rowCount = sheet.getMaxRows() - 4;
  if (rowCount <= 0) return;
  SpreadsheetApp.flush();
  const rows = sheet.getRange(5, 1, rowCount, 4).getValues();
  const target = sheet.getRange(5, 3, rowCount, 1);
  if (!rows.some(row => row[0] === 'Construction Expenses')) {
    target.setNumberFormat('General');
    return;
  }
  const source = sheet.getParent().getRangeByName('S_CONSTRUCTION_EXP_P');
  if (!source) throw new Error('The construction timeline named range is missing.');
  const values = source.getValues();
  const monthColumn = values[0].findIndex(value => value instanceof Date);
  if (monthColumn < 0) throw new Error('Construction timeline month headers are missing.');
  const formats = source.getSheet().getRange(source.getRow(), source.getColumn() + monthColumn, source.getNumRows(), 1).getNumberFormats();
  target.setNumberFormats(_detailsUnitFormats(rows, values, formats));
}

function _updateDetailsPeriodSelection(sheet, selectedValue) {
  const literal = '"' + String(selectedValue).replace(/"/g, '""') + '"';
  sheet.getRange('C2').setFormula(DETAILS_PERIOD_VALUE_FORMULA.replace('__SELECTION__', literal));
  SpreadsheetApp.flush();
}

function _detailsPeriodOptions(months, period) {
  const allLabels = { Month: 'All Month', Quarter: 'All Quarters', Year: 'All Years' };
  if (!allLabels[period]) throw new Error(`Unsupported details period: ${period}`);
  const dates = months.filter(month => month instanceof Date && Number.isFinite(month.getTime()));
  if (!dates.length) throw new Error('The project timeline has no months to select.');
  return [allLabels[period], ...new Set(dates.map(month => _detailsPeriodLabel(month, period)))];
}

function _updateDetailsPeriodOptions(sheet) {
  const period = sheet.getRange('A2').getValue();
  const months = sheet.getParent().getRangeByName('R_MONTHS');
  if (!months) throw new Error('The project timeline named range R_MONTHS is missing.');
  const options = _detailsPeriodOptions(months.getValues().flat(), period);
  const validation = SpreadsheetApp.newDataValidation().requireValueInList(options, true).setAllowInvalid(false).build();
  sheet.getRange('C2').setDataValidation(validation);
  SpreadsheetApp.flush();
}
