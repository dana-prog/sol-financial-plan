function debugCell() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cash Flow");
  const cell = sheet.getRange("AI3");

  Logger.log("Display value : %s", cell.getDisplayValue());
  Logger.log("Raw value     : %s", cell.getValue());
  Logger.log("Formula       : %s", cell.getFormula());
  Logger.log("FormulaR1C1   : %s", cell.getFormulaR1C1());
  Logger.log("DisplayValues : %s", sheet.getDataRange().getDisplayValues()[41][6]);
}