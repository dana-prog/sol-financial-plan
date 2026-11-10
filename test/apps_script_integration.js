// Run in the bound Apps Script project; creates a separate validation spreadsheet.
function validateInLoansFreezeIntegration() {
  const spreadsheet = SpreadsheetApp.create('In Loans freeze-unfreeze validation', 4, 5);
  console.log('Validation spreadsheet: ' + spreadsheet.getUrl());
  const sheet = spreadsheet.getSheets()[0].setName('In Loans');
  sheet.getRange('A1').setValue(5);
  sheet.getRange('B1').setFormula('=SEQUENCE(2,2,A1,1)');
  sheet.getRange('A4').setFormula('="=literal"');
  sheet.getRange('B4').setFormula('=""');
  sheet.getRange('C4').setFormula('=FALSE');
  SpreadsheetApp.flush();
  _freezeInLoans(spreadsheet);
  _freezeInLoans(spreadsheet);
  if (sheet.getRange('B1').getFormula()) throw new Error('Formula was not frozen.');
  if (sheet.getRange('C2').getValue() !== 8) throw new Error('Spill result was not retained.');
  sheet.getRange('A1').setValue(10);
  SpreadsheetApp.flush();
  if (sheet.getRange('B1').getValue() !== 5) throw new Error('Frozen result changed.');
  _unfreezeInLoans(spreadsheet);
  SpreadsheetApp.flush();
  if (sheet.getRange('A1').getValue() !== 10) throw new Error('Input edit was lost.');
  if (sheet.getRange('B1').getFormula() !== '=SEQUENCE(2,2,A1,1)') throw new Error('Formula not restored.');
  if (sheet.getRange('C2').getValue() !== 13) throw new Error('Spill did not recalculate.');
  if (sheet.getRange('A4').getValue() !== '=literal') throw new Error('Literal text changed.');
  if (sheet.getRange('C4').getValue() !== false) throw new Error('Boolean result changed.');
  _freezeInLoans(spreadsheet);
  sheet.getRange('C2').setValue(999);
  SpreadsheetApp.flush();
  let refused = false;
  try { _unfreezeInLoans(spreadsheet); } catch (error) { refused = /was changed/.test(error.message); }
  if (!refused) throw new Error('Changed frozen result was overwritten.');
  sheet.getRange('C2').setValue(13);
  SpreadsheetApp.flush();
  _unfreezeInLoans(spreadsheet);
  console.log('PASS: freeze, repeated freeze, input preservation, array restoration, conflict detection, second cycle.');
  return spreadsheet.getUrl();
}

function validateInLoansDropdownIntegration() {
  const spreadsheet = SpreadsheetApp.create('In Loans dropdown validation', 5, 5);
  console.log('Validation spreadsheet: ' + spreadsheet.getUrl());
  const sheet = spreadsheet.getSheets()[0].setName('In Loans');
  const toggle = sheet.getRange('A1').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Frozen Values', 'Live Values'], true).setAllowInvalid(false).build());
  toggle.setValue('Live Values');
  sheet.getRange('B1').setValue(5);
  sheet.getRange('B3').setFormula('=SEQUENCE(2,2,B1,1)');
  SpreadsheetApp.flush();
  const select = value => {
    toggle.setValue(value);
    SpreadsheetApp.flush();
    updateInLoansCalculationMode({range: toggle, source: spreadsheet, value});
  };
  select('Frozen Values');
  if (_readInLoansBackup(spreadsheet)?.state !== 'frozen') throw new Error('Dropdown did not freeze.');
  if (toggle.getValue() !== 'Frozen Values') throw new Error('Dropdown is out of sync.');
  sheet.getRange('B1').setValue(10);
  select('Live Values');
  if (_readInLoansBackup(spreadsheet)?.state !== 'live') throw new Error('Dropdown did not unfreeze.');
  if (sheet.getRange('C4').getValue() !== 13) throw new Error('Array did not recalculate.');
  _applyInLoansCalculationAction(spreadsheet, _freezeInLoans);
  if (toggle.getValue() !== 'Frozen Values') throw new Error('Menu action did not sync dropdown.');
  sheet.getRange('C4').setValue(999);
  select('Live Values');
  if (toggle.getValue() !== 'Frozen Values') throw new Error('Failed operation left incorrect dropdown state.');
  sheet.getRange('C4').setValue(13);
  select('Live Values');
  if (toggle.getValue() !== 'Live Values') throw new Error('Final unfreeze failed.');
  console.log('PASS: dropdown freeze/unfreeze, menu sync, failure recovery, and preserved input edits.');
}
