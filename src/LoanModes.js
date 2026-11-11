function handleLoanModeEdit(e) {
  if (!e || !e.range || e.range.getRow() !== 1 ||
      ![1, 4].includes(e.range.getColumn()) || e.range.getGridId() !== IN_LOANS_SHEET_ID) return false;

  const column = e.range.getColumn();
  const label = column === 1 ? 'Repayment Calculations' : 'Planned Loans';
  const spreadsheet = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const startedAt = Date.now();
  console.info(label + ': edit received');
  let cell;
  let message;
  let failure;
  try {
    // Range coordinates and grid ID come from the event; avoid sheet access before feedback.
    spreadsheet.toast('Applying your selection. Please wait.', label, -1);
    console.info(label + ': progress displayed after ' + (Date.now() - startedAt) + ' ms');
    cell = e.range.getSheet().getRange(1, column);
    const mode = cell.getValue();
    console.info(label + ': selection read after ' + (Date.now() - startedAt) + ' ms');
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
    if (cell) cell.setNote('The selected mode was not applied successfully. ' + error.message);
  } finally {
    spreadsheet.toast('', label, 1);
  }

  // Alerts suspend execution, so show them only after the repayment lock is released.
  console.info(label + ': operation finished after ' + (Date.now() - startedAt) + ' ms');
  const ui = SpreadsheetApp.getUi();
  ui.alert(label + (failure ? ' — Failed' : ' — Done'),
    failure ? failure.message : message, ui.ButtonSet.OK);
  if (failure) throw failure;
  return true;
}
