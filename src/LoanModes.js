function handleLoanModeEdit(e) {
  if (!e || !e.range || e.range.getRow() !== 1 || e.range.getColumn() !== 1 ||
      e.range.getGridId() !== IN_LOANS_SHEET_ID) return false;
  if (e.value !== 'Planned Loans Enabled' && e.value !== 'Planned Loans Disabled') {
    SpreadsheetApp.getUi().alert('Select Planned Loans Enabled or Planned Loans Disabled.');
    return true;
  }
  const enabled = e.value === 'Planned Loans Enabled';
  const template = HtmlService.createTemplateFromFile('LoanModeProgress');
  template.action = enabled ? 'Enabling Planned Loans…' : 'Disabling Planned Loans…';
  template.completionMessage = (enabled
    ? 'Planned loans enabled.'
    : 'Planned loans disabled.')
    + ' Recalculate loan repayments to update.';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(420).setHeight(105), 'Planned Loans');
  return true;
}
