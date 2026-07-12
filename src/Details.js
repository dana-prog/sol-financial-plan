const DETAILS_SHEET_ID = 2109063306;

const paramMap = {
  "construction expenses": "Construction Expenses",
  "infrastructure expenses": "Infrastructure Expenses",
  "operational expenses": "Operational Expenses",
  "land expenses": "Land Expenses",
};

function showDetails(param, month) {
  const detailsParam = paramMap[param];

  if (!detailsParam) {
    SpreadsheetApp.getUi().alert(
      `"${param}" does not have a details view.`
    );
    return;
  }

  const detailsSheet = SpreadsheetApp.getActive().getSheetById(DETAILS_SHEET_ID);

  detailsSheet.getRange("B1").setValue(detailsParam);
  detailsSheet.getRange("B2").setValue(month);

  detailsSheet.activate();
}