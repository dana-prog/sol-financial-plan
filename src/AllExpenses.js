const ALL_EXPENSES_TYPES = [
  'Borrowing Repayments', 'Loan Disbursements', 'Setup Expenses', 'Land Expenses',
  'Infrastructure Expenses', 'Construction Expenses', 'Direct Operating Expenses',
  'Other Operating Expenses',
];
const ALL_EXPENSES_CALCULATION_ROWS = 20000;

function refreshAllExpenses() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) throw new Error('Another refresh is running. Please try again shortly.');
  try {
    const metadata = Sheets.Spreadsheets.get(spreadsheetId, {fields: 'sheets(properties,basicFilter,rowGroups,merges)'});
    const sheet = metadata.sheets.find(sheet => sheet.properties.title === 'All Expenses') ||
      metadata.sheets.find(sheet => sheet.properties.title === 'Details');
    if (!sheet) throw new Error('Could not find All Expenses or Details.');
    const expenses = _calculateAllExpenses(spreadsheetId);
    const layout = _allExpensesLayout(expenses);
    const requests = _allExpensesWriteRequests(sheet, layout);
    Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
    _createMenu();
    spreadsheet.toast(`Updated ${expenses.length} expense items.`, 'All Expenses', 5);
  } finally {
    lock.releaseLock();
  }
}

function _calculateAllExpenses(spreadsheetId) {
  // Evaluate the existing named-function formula in a temporary sheet. Never
  // clear the last successful display until the complete new result is valid.
  const sheetId = Math.floor(Math.random() * 2000000000);
  const title = `_expenses_calculation_${Date.now()}`;
  Sheets.Spreadsheets.batchUpdate({requests: [
    {addSheet: {properties: {sheetId, title, hidden: true,
      gridProperties: {rowCount: ALL_EXPENSES_CALCULATION_ROWS, columnCount: 6}}}},
    {updateCells: {start: {sheetId, rowIndex: 0, columnIndex: 0},
      rows: [{values: [{userEnteredValue: {formulaValue: ALL_EXPENSES_FORMULA}}]}], fields: 'userEnteredValue'}},
  ]}, spreadsheetId);
  try {
    const deadline = Date.now() + 120000;
    let rows;
    do {
      const range = `'${title}'!A1:F${ALL_EXPENSES_CALCULATION_ROWS}`;
      const result = Sheets.Spreadsheets.Values.get(spreadsheetId, range, {valueRenderOption: 'UNFORMATTED_VALUE'});
      rows = result.values || [];
      if (rows[0]?.[0] === '__NO_EXPENSES__') return [];
      if (ALL_EXPENSES_TYPES.includes(rows[0]?.[0])) {
        _validateAllExpenses(rows);
        return rows;
      }
      Utilities.sleep(1000);
    } while (Date.now() < deadline);
    throw new Error(`Expense calculation did not finish (${rows?.[0]?.[0] || 'blank result'}). Existing rows were preserved.`);
  } finally {
    Sheets.Spreadsheets.batchUpdate({requests: [{deleteSheet: {sheetId}}]}, spreadsheetId);
  }
}

function _validateAllExpenses(rows) {
  for (const row of rows) {
    if (row.length !== 6 || !ALL_EXPENSES_TYPES.includes(row[0]) ||
        typeof row[1] !== 'number' || typeof row[5] !== 'number' || !Number.isFinite(row[5])) {
      throw new Error('Expense calculation contains incomplete or invalid rows. Existing rows were preserved.');
    }
  }
}

function _allExpensesLayout(expenses) {
  _validateAllExpenses(expenses);
  const rows = [];
  const subtotals = [];
  const sections = [];
  for (const type of ALL_EXPENSES_TYPES) {
    const items = expenses.filter(row => row[0] === type)
      .sort((a, b) => Number(a[3]) - Number(b[3]) || String(a[4]).localeCompare(String(b[4])));
    if (!items.length) continue;
    const typeStart = rows.length + 2;
    for (const year of [...new Set(items.map(row => row[1]))]) {
      const yearly = items.filter(row => row[1] === year);
      const yearStart = rows.length + 2;
      for (const quarter of [...new Set(yearly.map(row => row[2]))]) {
        const quarterStart = rows.length + 2;
        rows.push(...yearly.filter(row => row[2] === quarter));
        sections.push({startRow: quarterStart, endRow: rows.length + 1, depth: 3});
        subtotals.push(rows.length + 2);
        rows.push([type, year, quarter, '', 'Quarter subtotal', `=SUBTOTAL(9,F${quarterStart}:F${rows.length + 1})`]);
      }
      sections.push({startRow: yearStart, endRow: rows.length + 1, depth: 2});
      subtotals.push(rows.length + 2);
      rows.push([type, year, '', '', 'Year subtotal', `=SUBTOTAL(9,F${yearStart}:F${rows.length + 1})`]);
    }
    sections.push({startRow: typeStart, endRow: rows.length + 1, depth: 1});
    subtotals.push(rows.length + 2);
    rows.push([type, '', '', '', 'Type subtotal', `=SUBTOTAL(9,F${typeStart}:F${rows.length + 1})`]);
  }
  const totalRow = rows.length + 2;
  rows.push(['Total', '', '', '', '', totalRow === 2 ? '=0' : `=SUBTOTAL(9,F2:F${totalRow - 1})`]);
  return { rows, subtotals, sections, totalRow };
}

function _allExpensesWriteRequests(sheet, layout) {
  const sheetId = sheet.properties.sheetId;
  const oldRowCount = sheet.properties.gridProperties.rowCount;
  const bodyRange = {sheetId, startRowIndex: 1, endRowIndex: Math.max(oldRowCount, layout.totalRow), startColumnIndex: 0, endColumnIndex: 6};
  const requests = [];
  if (sheet.basicFilter) requests.push({clearBasicFilter: {sheetId}});
  for (const range of sheet.merges || []) requests.push({unmergeCells: {range}});
  for (const group of [...(sheet.rowGroups || [])].sort((a,b) => b.depth - a.depth)) {
    requests.push({deleteDimensionGroup: {range: group.range}});
  }
  if (layout.totalRow > oldRowCount) {
    requests.push({appendDimension: {sheetId, dimension: 'ROWS', length: layout.totalRow - oldRowCount}});
  }
  requests.push({repeatCell: {range: bodyRange, cell: {}, fields: 'userEnteredValue,textFormatRuns'}});
  const rows = layout.rows.map((row,index) => ({values: row.map((value,column) => {
    const rowNumber = index + 2;
    const summary = layout.subtotals.includes(rowNumber) || rowNumber === layout.totalRow;
    const format = {
      backgroundColor: {red: summary ? 0.93 : 1, green: summary ? 0.93 : 1, blue: summary ? 0.93 : 1},
      verticalAlignment: 'MIDDLE', horizontalAlignment: column === 5 ? 'RIGHT' : column >= 1 && column <= 3 ? 'CENTER' : 'LEFT',
      wrapStrategy: column === 4 ? 'WRAP' : 'OVERFLOW_CELL',
      textFormat: {fontFamily: 'Arial', fontSize: 11, bold: summary, foregroundColor: {red: 0.4, green: 0.4, blue: 0.4}},
    };
    if (column === 1) format.numberFormat = {type: 'NUMBER', pattern: '0'};
    if (column === 2) format.numberFormat = {type: 'TEXT', pattern: '@'};
    if (column === 3) format.numberFormat = {type: 'DATE', pattern: 'mmmm yyyy'};
    if (column === 5) format.numberFormat = {type: 'NUMBER', pattern: '฿ #,##0,"K"'};
    const cell = {userEnteredFormat: format, userEnteredValue: typeof value === 'number' ? {numberValue: value}
      : column === 5 && summary ? {formulaValue: value} : {stringValue: String(value)}};
    if (column === 4 && value) cell.textFormatRuns = _allExpensesTextRuns(String(value));
    return cell;
  })}));
  requests.push({updateCells: {range: {...bodyRange, endRowIndex: layout.totalRow}, rows,
    fields: 'userEnteredValue,userEnteredFormat,textFormatRuns'}});
  const borderRows = _detailsBorderRows([['Type', 'Year'], ...layout.rows]);
  for (const border of borderRows) {
    requests.push({updateBorders: {range: {sheetId, startRowIndex: border.row - 1, endRowIndex: border.row, startColumnIndex: 0, endColumnIndex: 6},
      bottom: {style: border.style, color: {red: 0.4, green: 0.4, blue: 0.4}}}});
  }
  requests.push({updateDimensionProperties: {range: {sheetId, dimension: 'ROWS', startIndex: 1, endIndex: layout.totalRow},
    properties: {hiddenByUser: false}, fields: 'hiddenByUser'}});
  for (const section of layout.sections) {
    requests.push({mergeCells: {range: {sheetId, startRowIndex: section.startRow - 1, endRowIndex: section.endRow + 1,
      startColumnIndex: section.depth - 1, endColumnIndex: section.depth}, mergeType: 'MERGE_ALL'}});
  }
  if (layout.totalRow > 2) {
    const filterSpecs = sheet.basicFilter?.filterSpecs || Object.entries(sheet.basicFilter?.criteria || {})
      .map(([column, filterCriteria]) => ({columnIndex: Number(column), filterCriteria}));
    const preserved = filterSpecs.map(spec => ({...spec, filterCriteria: {...spec.filterCriteria}}));
    for (const spec of preserved) {
      if (spec.columnIndex >= 1 && spec.columnIndex <= 3 && spec.filterCriteria.hiddenValues) {
        spec.filterCriteria.hiddenValues = spec.filterCriteria.hiddenValues.filter(value => value !== '');
      }
    }
    requests.push({setBasicFilter: {filter: {range: {sheetId, startRowIndex: 0, endRowIndex: layout.totalRow - 1, startColumnIndex: 0, endColumnIndex: 6}, filterSpecs: preserved}}});
  }
  requests.push({updateSheetProperties: {properties: {sheetId, gridProperties: {rowCount: layout.totalRow}},
    fields: 'gridProperties.rowCount'}});
  return requests;
}

function _parenthesisSpans(text) {
  const spans = [];
  let start = -1;
  let depth = 0;
  for (let index = 0; index < text.length; index++) {
    if (text[index] === '(') {
      if (depth === 0) start = index;
      depth++;
    } else if (text[index] === ')' && depth > 0) {
      if (--depth === 0) spans.push([start, index + 1]);
    }
  }
  return spans;
}

function _allExpensesTextRuns(text) {
  const runs = [{startIndex: 0, format: {fontSize: 11}}];
  for (const [start,end] of _parenthesisSpans(text)) {
    if (runs[runs.length - 1].startIndex === start) runs[runs.length - 1].format.fontSize = 10;
    else runs.push({startIndex: start, format: {fontSize: 10}});
    if (end < text.length) runs.push({startIndex: end, format: {fontSize: 11}});
  }
  return runs;
}
