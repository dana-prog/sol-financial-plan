// assumptions:
// 1. The first row is years (starting from the 4th column)
// 2. The second row is quarters (starting from the 4th column)
// 3. The first column is category (starting from the 4th row)
function setTimelineSheetBorders() {
  const sheetName = 'Construction';
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const HEADER_COL = 4;
  const CATEGORY_ROW = 4;

  // Clear existing borders
  sheet
    .getRange(1, 1, lastRow, lastCol)
    .setBorder(false, false, false, false, false, false);

  //
  // Vertical borders (years / quarters)
  //
  const years = sheet
    .getRange(1, HEADER_COL, 1, lastCol - HEADER_COL + 1)
    .getValues()[0];

  const quarters = sheet
    .getRange(2, HEADER_COL, 1, lastCol - HEADER_COL + 1)
    .getValues()[0];

  for (let i = 1; i < years.length; i++) {
    const col = HEADER_COL + i;

    if (years[i] !== years[i - 1]) {
      // Thick border before this column
      sheet
        .getRange(1, col, lastRow, 1)
        .setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else if (quarters[i] !== quarters[i - 1]) {
      // Thin border before this column
      sheet
        .getRange(1, col, lastRow, 1)
        .setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  //
  // Horizontal borders (categories)
  //
  const categories = sheet
    .getRange(CATEGORY_ROW, 1, lastRow - CATEGORY_ROW + 1, 1)
    .getValues()
    .flat();

  for (let i = 1; i < categories.length; i++) {
    if (categories[i] !== categories[i - 1]) {
      const row = CATEGORY_ROW + i;

      sheet
        .getRange(row, 1, 1, lastCol)
        .setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    }
  }
}