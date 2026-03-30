const STAFF_SHEET_NAME = 'Staff';
const STAFF_SHEET_ID = 1162211165;
const STAFF_ROLE_COLUMN_HEADER = 'role';

function getStaffRoles() {
  return SOLLibrary
    .getColumnValues(STAFF_SHEET_NAME, STAFF_ROLE_COLUMN_HEADER, false);
}

function getStaffSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetById(STAFF_SHEET_ID);
}

