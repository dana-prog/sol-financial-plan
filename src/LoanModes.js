function handleLoanModeEdit(e) {
  if (!e || !e.range || e.range.getRow() !== 1 ||
      ![1, 4].includes(e.range.getColumn()) || e.range.getGridId() !== IN_LOANS_SHEET_ID) return false;

  const column = e.range.getColumn();
  const label = column === 1 ? 'Repayment Calculations' : 'Planned Loans';
  if (e.value !== label + ' Enabled' && e.value !== label + ' Disabled') {
    SpreadsheetApp.getUi().alert('Select ' + label + ' Enabled or ' + label + ' Disabled.');
    return true;
  }
  const template = HtmlService.createTemplateFromFile('LoanModeProgress');
  const enabled = e.value === label + ' Enabled';
  template.action = column === 1
    ? (enabled ? 'Restoring repayment formulas…' : 'Freezing repayments…')
    : (enabled ? 'Enabling Planned Loans…' : 'Disabling Planned Loans…');
  // D1 already controls the draw formulas. Confirm the edit without a server round trip.
  template.completionMessage = column === 4
    ? e.value + '. Sheets recalculates the affected formulas automatically.'
    : '';
  // Open the progress UI before any sheet reads can wait for recalculation.
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(105), label);
  return true;
}

// Read the current selection so an older dialog cannot apply a stale edit.
function applyRepaymentMode() {
  const sheet = getRepaymentSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const mode = sheet.getRange(REPAYMENT_FREEZE.toggleCell).getValue();
  if (mode === REPAYMENT_FREEZE.frozenValue) return freezeLoanRepayments();
  if (mode === REPAYMENT_FREEZE.liveValue) return unfreezeLoanRepayments();
  throw new Error('Select Repayment Calculations Enabled or Repayment Calculations Disabled.');
}
