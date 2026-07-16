const DETAILS_SHEET_ID = 2109063306;

const params = [
  "construction expenses",
  "setup expenses",
  "operational expenses",
  "land expenses"
];

function showDetails(param, month) {
  if (!params.find((value) => param === value)) {
    SpreadsheetApp.getUi().alert(
      `"${param}" does not have a details view.`
    );
    return;
  }

  const detailsSheet = SpreadsheetApp.getActive().getSheetById(DETAILS_SHEET_ID);

  detailsSheet.getRange("B1").setValue(SOLLibrary.capitalize(param));
  detailsSheet.getRange("B2").setValue(month);

  detailsSheet.activate();
}