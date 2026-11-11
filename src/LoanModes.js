function handleLoanModeEdit(e) {
  if (!e || !e.range || e.range.getSheet().getName() !== 'In Loans' ||
      e.range.getRow() !== 1 || ![1, 4].includes(e.range.getColumn())) return false;

  const column = e.range.getColumn();
  const cell = e.range.getSheet().getRange(1, column);
  const label = column === 1 ? 'Repayment Calculations' : 'Planned Loans';
  const mode = cell.getValue();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let message;
  let failure;
  try {
    if (mode !== label + ' Enabled' && mode !== label + ' Disabled') {
      throw new Error('Select ' + label + ' Enabled or ' + label + ' Disabled.');
    }
    if (column === 1) {
      const disabled = mode === label + ' Disabled';
      spreadsheet.toast(disabled ? 'Freezing repayments. Please wait.' :
        'Restoring repayment formulas. Please wait.', label, -1);
      message = disabled ? freezeLoanRepayments() : unfreezeLoanRepayments();
    } else {
      spreadsheet.toast('Applying ' + mode + '. Please wait.', label, -1);
      // Loan draw formulas read D1 directly; no formula rewrite is needed.
      SpreadsheetApp.flush();
      message = mode === label + ' Disabled'
        ? 'Planned loans are disabled. Their loan draws are set to zero.'
        : 'Planned loans are enabled. Their loan draws use the original calculations.';
    }
    cell.setNote('');
  } catch (error) {
    failure = error;
    cell.setNote('The selected mode was not applied successfully. ' + error.message);
  } finally {
    spreadsheet.toast('', label, 1);
  }

  // Alerts suspend execution, so show them only after the repayment lock is released.
  const ui = SpreadsheetApp.getUi();
  ui.alert(label + (failure ? ' — Failed' : ' — Done'),
    failure ? failure.message : message, ui.ButtonSet.OK);
  if (failure) throw failure;
  return true;
}
