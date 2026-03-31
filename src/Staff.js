const STAFF_SHEET_NAME = 'Staff';
const STAFF_SHEET_ID = 1162211165;
const STAFF_ROLE_COLUMN_HEADER = 'role';
const STAFF_HEADCOUNT_MODE_COLUMN_HEADER = 'headcount mode';
const STAFF_HEADCOUNT_VALUE_COLUMN_HEADER = 'headcount value';

function getStaffRoles() {
  return SOLLibrary
    .getColumnValues(STAFF_SHEET_NAME, STAFF_ROLE_COLUMN_HEADER, false);
}

function onEditStaffSheet(oldValue, newValue, rowNum, colNum) {
  const sheet = getStaffSheet();
  // const roleCol = SOLLibrary.getColNumByHeader(sheet, STAFF_ROLE_COLUMN_HEADER);
  // const headcountModeCol = SOLLibrary.getColNumByHeader(sheet, STAFF_HEADCOUNT_MODE_COLUMN_HEADER);
  const roleCol = 1;
  const headcountModeCol = 2;

  if (colNum !== roleCol && colNum !== headcountModeCol) {
    SOLLibrary.log('Staff', 'onEditStaffSheet', `colNum ${colNum} is not a role or headcount mode column`);
    return;
  }

  switch (colNum) {
    case roleCol:
      SOLLibrary.log('Staff', 'onEditStaffSheet', `role changed from '${oldValue}' to '${newValue}'`);
      const headcountMode = sheet.getRange(rowNum, headcountModeCol).getValue();
      _updateHeadcountValueFormat(rowNum, newValue, headcountMode === 'fixed');
      if (oldValue !== '' && newValue !== '') {
        updateTimelineStaffParams(oldValue, newValue);
      }
      break;
    case headcountModeCol:
      SOLLibrary.log('Staff', 'onEditStaffSheet', `headcount mode changed from '${oldValue}' to '${newValue}'`);
      const role = sheet.getRange(rowNum, roleCol).getValue();
      _updateHeadcountValueFormat(rowNum, role, newValue === 'fixed');
      break;
  }
}

function getStaffSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetById(STAFF_SHEET_ID);
}

function _updateHeadcountValueFormat(rowNum, role, fixed) {
  const sheet = getStaffSheet();
  // const headcountValueCol = SOLLibrary.getColNumByHeader(sheet, STAFF_HEADCOUNT_VALUE_COLUMN_HEADER);
  const headcountValueCol = 3;
  sheet.getRange(rowNum, headcountValueCol)
    .setNumberFormat(fixed ? getCountNumberFormat(role) : `"1 ${role} per" 0 "guests"`);
}
