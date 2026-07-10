// ============================================================                                                                                                  
  //  Cash Flow Timeline — freeze / thaw via A1 dropdown                                                                                                         
  // ============================================================                                                                                                  
  const CF_SHEET_NAME = 'Cash Flow';
  const CF_TOGGLE ='B1';                                                                                                                                      
  const CF_FIRST_ROW  = 4;                                                                                                                                         
  const CF_FIRST_COL  = 4;    // D
  const CF_NUM_COLS   = 96;   // D:CU                                                                                                                              
  const CF_PROP_KEY   = 'frozenCashFlowTimelineFormulas';                

function freezeCashFlow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CF_SHEET_NAME);
  if (!sheet) {
    SOLLibrary.log('_CashFlowFreeze', 'Sheet not found: ' + CF_SHEET_NAME, 'WARN');
    return;
  }

  SOLLibrary.log('_CashFlowFreeze', 'Before getLastRow');
  const lastRow = sheet.getLastRow();
  if (lastRow < CF_FIRST_ROW) return;
  SOLLibrary.logArgs('_CashFlowFreeze', 'freezeCashFlow', {lastRow});

  const colRange = sheet.getRange(CF_FIRST_ROW, CF_FIRST_COL, lastRow - CF_FIRST_ROW + 1, 1);
  const allRange = sheet.getRange(CF_FIRST_ROW, CF_FIRST_COL, lastRow - CF_FIRST_ROW + 1, CF_NUM_COLS);
  const formulas = colRange.getFormulas();
  SOLLibrary.logArgs('_CashFlowFreeze', 'freezeCashFlow', {formulas});

  saveChunked_(CF_PROP_KEY, JSON.stringify(formulas));
  const allValues = allRange.getValues();
  SOLLibrary.logArgs('_CashFlowFreeze', 'freezeCashFlow', {allValues});
  allRange.setValues(allValues);           // formulas → plain values
  SOLLibrary.alert('Cash Flow Timeline', 'Sheet was set with static values');
}

function unfreezeCashFlow() {
  const json = loadChunked_(CF_PROP_KEY);
  if (!json) {
    SOLLibrary.alert('Cash Flow Timeline', 'No formulas found to set to the sheet');
    return;
  }
  const formulas = JSON.parse(json);
  const numRows = formulas.length;
  const numCols = numRows ? formulas[0].length : 0;
  if (!numRows || !numCols) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CF_SHEET_NAME);
  const range = sheet.getRange(CF_FIRST_ROW, CF_FIRST_COL, numRows, numCols);
  const values = range.getValues();
  // Mix formulas + values: setValues writes "=…" strings as formulas
  const merged = formulas.map((row, i) =>
    row.map((f, j) => f || values[i][j])
  );
  range.setValues(merged);
  deleteChunked_(CF_PROP_KEY);
  SOLLibrary.alert('Cash Flow Timeline', 'Sheet was set with auto calculated formulas');
}
                                                                                                                                                              
   
  // ============================================================                                                                                                  
  //  Chunked property storage (9KB per-value limit in PropertiesService)
  // ============================================================
  const CHUNK_SIZE = 8000;                                                                                                                                         
   
  function saveChunked_(key, str) {                                                                                                                                
    const props = PropertiesService.getDocumentProperties();
    deleteChunked_(key);
    let count = 0;                                                                                                                                                 
    for (let pos = 0; pos < str.length; pos += CHUNK_SIZE) {
      props.setProperty(`${key}_${count}`, str.substr(pos, CHUNK_SIZE));                                                                                           
      count++;                                                                                                                                                     
    }
    props.setProperty(`${key}_count`, String(count));                                                                                                              
  }                                                      

  function loadChunked_(key) {                                                                                                                                     
    const props = PropertiesService.getDocumentProperties();
    const count = +props.getProperty(`${key}_count`) || 0;                                                                                                         
    if (!count) return null;                             
    let str = '';
    for (let i = 0; i < count; i++) {
      str += props.getProperty(`${key}_${i}`) || '';                                                                                                               
    }
    return str;                                                                                                                                                    
  }                                                      

  function deleteChunked_(key) {                                                                                                                                   
    const props = PropertiesService.getDocumentProperties();
    const count = +props.getProperty(`${key}_count`) || 0;                                                                                                         
    for (let i = 0; i < count; i++) props.deleteProperty(`${key}_${i}`);
    props.deleteProperty(`${key}_count`);                                                                                                                          
  }