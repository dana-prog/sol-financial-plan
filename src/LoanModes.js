function handleLoanModeEdit(e) {
  if (!e || !e.range || e.range.getRow() !== 1 ||
      ![1, 4].includes(e.range.getColumn()) || e.range.getGridId() !== IN_LOANS_SHEET_ID) return false;

  const column = e.range.getColumn();
  const template = HtmlService.createTemplateFromFile('LoanModeProgress');
  template.column = column;
  // Open the progress UI before any sheet reads can wait for recalculation.
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(210),
    column === 1 ? 'Repayment Calculations' : 'Planned Loans');
  return true;
}

// Called by the progress dialog after its spinner has rendered.
function applyLoanMode(column) {
  if (![1, 4].includes(column)) throw new Error('Invalid loan control.');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetById(IN_LOANS_SHEET_ID);
  if (!sheet) throw new Error('The In Loans sheet could not be found.');
  const cell = sheet.getRange(1, column);
  const label = column === 1 ? 'Repayment Calculations' : 'Planned Loans';
  try {
    const mode = cell.getValue();
    if (mode !== label + ' Enabled' && mode !== label + ' Disabled') {
      throw new Error('Select ' + label + ' Enabled or ' + label + ' Disabled.');
    }
    let message;
    if (column === 1) {
      message = mode === label + ' Disabled' ? freezeLoanRepayments() : unfreezeLoanRepayments();
    } else {
      // Loan draw formulas read D1 directly; no formula rewrite is needed.
      SpreadsheetApp.flush();
      message = mode === label + ' Disabled'
        ? 'Planned loans are disabled. Their loan draws are set to zero.'
        : 'Planned loans are enabled. Their loan draws use the original calculations.';
    }
    cell.setNote('');
    return message;
  } catch (error) {
    cell.setNote('The selected mode was not applied successfully. ' + error.message);
    throw error;
  }
}
