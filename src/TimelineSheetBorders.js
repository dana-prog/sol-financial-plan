const YEARS_ROW_NUM = 1;
const QUARTERS_ROW_NUM = 2;
const HEADER_ROW = 3;
const FIRST_CATEGORY_ROW = 4;
const CATEGORY_COL = 1;
const BORDER_COLOR = "#666666";

// assumptions:
// 1. The first row is years (starting from the 4th column)
// 2. The second row is quarters (starting from the 4th column)
// 3. The first column is category (starting from the 4th row)
function setTimelineSheetBorders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const lastRow = sheet.getLastRow();
  const firstDateCol = _getFirstNumericColumn(sheet, YEARS_ROW_NUM);
  const lastCol = sheet.getLastColumn();

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
    .getRange(FIRST_CATEGORY_ROW, CATEGORY_COL, lastRow - FIRST_CATEGORY_ROW + 1, 1)
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
}

function _getFirstNumericColumn(sheet, rowNum) {
  const headers = sheet
    .getRange(rowNum, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const index = headers.findIndex(v => typeof v === 'number' && !isNaN(v));

  return index === -1 ? -1 : index + 1;
}