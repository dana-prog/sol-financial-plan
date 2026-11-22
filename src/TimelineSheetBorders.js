const YEARS_ROW_NUM = 1;
const QUARTERS_ROW_NUM = 2;
const HEADER_ROW = 3;
const FIRST_CATEGORY_ROW = 4;
const BORDER_COLOR = "#666666";

// assumptions:
// 1. The first row is years (starting from the 4th column)
// 2. The second row is quarters (starting from the 4th column)
// 3. Row 3 contains a category header; category labels start in row 4.
function setTimelineSheetBorders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const lastRow = sheet.getLastRow();
  const firstDateCol = _getFirstNumericColumn(sheet, YEARS_ROW_NUM);
  const lastCol = sheet.getLastColumn();

  if (firstDateCol < 1 || lastRow < FIRST_CATEGORY_ROW) throw new Error('Select a populated timeline sheet. No borders were changed.');
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const categoryCol = headers.findIndex(value => String(value).trim().toLowerCase() === 'category') + 1;
  if (!categoryCol) throw new Error('Could not find the Category column. No borders were changed.');
  if (sheet.getName() === 'In Loans' || sheet.getName() === 'Out Loans') {
    if (!headers.some(value => String(value).trim().toLowerCase() === 'param')) {
      throw new Error('Could not find the Param column. No borders were changed.');
    }
  }

  // clear all existing borders
  sheet
    .getRange(1, 1, lastRow, lastCol)
    .setBorder(false, false, false, false, false, false);

  // Vertical borders (years & quarters)
  const years = sheet
    .getRange(YEARS_ROW_NUM, firstDateCol, 1, lastCol - firstDateCol + 1)
    .getValues()[0];

  const quarters = sheet
    .getRange(QUARTERS_ROW_NUM, firstDateCol, 1, lastCol - firstDateCol + 1)
    .getValues()[0];

  for (let i = 1; i < years.length; i++) {

    if (years[i] !== years[i - 1]) {
      // Set year border (thick)
      sheet
        .getRange(1, firstDateCol + i, lastRow, 1)
        .setBorder(null, true, null, null, null, null, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THICK);

      continue;
    }


    if (quarters[i] !== quarters[i - 1]) {
      // Set quarter border (thin)
      sheet
        .getRange(1, firstDateCol + i, lastRow, 1)
        .setBorder(null, true, null, null, null, null, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // Horizontal borders (categories)
  const categories = sheet
    .getRange(FIRST_CATEGORY_ROW, categoryCol, lastRow - FIRST_CATEGORY_ROW + 1, 1)
    .getValues()
    .flat();

  for (let i = 1; i < categories.length; i++) {
    // Set category border (thick)
    if (categories[i] !== "") {
      sheet
        .getRange(FIRST_CATEGORY_ROW + i, 1, 1, lastCol)
        .setBorder(true, null, null, null, null, null, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }


  if (sheet.getName() === "In Loans" || sheet.getName() === "Out Loans") {
    setLoanSectionBorders(sheet, lastRow, lastCol, categoryCol);
  }
}

function setLoanSectionBorders(sheet, lastRow, lastCol, categoryCol) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const paramCol = headers.findIndex(
    value => String(value).trim().toLowerCase() === "param"
  ) + 1;

  if (paramCol === 0) {
    throw new Error('Could not find the "Param" column.');
  }

  const rows = sheet
    .getRange(FIRST_CATEGORY_ROW, 1, lastRow - FIRST_CATEGORY_ROW + 1, Math.max(paramCol, categoryCol))
    .getValues();

  const borderAfter = new Set([
    "end balance",
    "principal carried forward",
    "interest carried forward"
  ]);

  let inLoanSection = false;

  rows.forEach((row, index) => {
    const category = String(row[categoryCol - 1]).trim().toLowerCase();
    const param = String(row[paramCol - 1]).trim().toLowerCase();

    if (category === "totals") {
      inLoanSection = false;
    } else if (param === "start balance") {
      inLoanSection = true;
    }

    if (inLoanSection && borderAfter.has(param)) {
      sheet
        .getRange(FIRST_CATEGORY_ROW + index, paramCol, 1, lastCol - paramCol + 1)
        .setBorder(
          null, null, true, null, null, null,
          BORDER_COLOR,
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
  });
}

function _getFirstNumericColumn(sheet, rowNum) {
  const headers = sheet
    .getRange(rowNum, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const index = headers.findIndex(v => typeof v === 'number' && !isNaN(v));

  return index === -1 ? -1 : index + 1;
}
