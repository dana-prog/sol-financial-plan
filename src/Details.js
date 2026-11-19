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
  "land expenses"
];

function showDetails(param, month) {
  const aliases = {
    'planning expenses': 'setup expenses',
    'infra expenses': 'infrastructure expenses',
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

  detailsSheet.getRange('B1').setValue('Month');
  detailsSheet.getRange('B3').setValue(SOLLibrary.capitalize(param));
  _updateDetailsPeriodSelection(detailsSheet, _detailsPeriodLabel(month, 'Month'));

  detailsSheet.activate();
}

function updateDetailsPeriod(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getSheetId() !== DETAILS_SHEET_ID ||
      range.getRow() > 1 || range.getColumn() > 2 || range.getLastColumn() < 2) return;
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
  const period = sheet.getRange('B1').getValue();
  const selected = sheet.getRange('B2');
  const months = sheet.getParent().getRangeByName('R_MONTHS');
  if (!months) throw new Error('The project timeline named range R_MONTHS is missing.');
  const selection = _detailsPeriodSelection(months.getValues().flat(), period,
    selectedValue === undefined ? selected.getValue() : selectedValue);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(selection.options, true)
    .setAllowInvalid(false)
    .build();
  selected.setDataValidation(validation).setNumberFormat('@').setValue(selection.value);
  sheet.getRange('A2').setValue(`Select ${period}`);
}
