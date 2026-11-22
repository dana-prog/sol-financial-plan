const ALL_EXPENSES_TYPES = [
  'In Loan Repayments', 'Out Loan Disbursements', 'Setup Expenses', 'Land Expenses',
  'Infrastructure Expenses', 'Construction Expenses', 'Direct Operating Expenses',
  'Other Operating Expenses',
];

function refreshAllExpenses() {
  const started = Date.now();
  const timings = {};
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) throw new Error('Another refresh is running. Please try again shortly.');
  try {
    let phase = Date.now();
    const metadata = Sheets.Spreadsheets.get(spreadsheetId, {fields: 'properties(timeZone),sheets(properties,basicFilter,rowGroups,merges)'});
    const sheet = metadata.sheets.find(sheet => sheet.properties.title === 'All Expenses') ||
      metadata.sheets.find(sheet => sheet.properties.title === 'Details');
    if (!sheet) throw new Error('Could not find All Expenses or Details.');
    const title = sheet.properties.title.replace(/'/g, "''");
    const snapshot = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [`'${title}'!A1:F${sheet.properties.gridProperties.rowCount}`],
      fields: 'sheets(data(rowData(values(userEnteredValue,textFormatRuns))))',
    });
    sheet.data = snapshot.sheets[0].data;
    timings.snapshotMs = Date.now() - phase;
    phase = Date.now();
    const source = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
      ranges: ALL_EXPENSES_SOURCE_RANGES, valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER',
    });
    timings.readSourcesMs = Date.now() - phase;
    phase = Date.now();
    const inputs = Object.fromEntries(ALL_EXPENSES_SOURCE_RANGES.map((name, index) => [name, source.valueRanges?.[index]?.values || []]));
    const monthStart = Utilities.formatDate(new Date(started), metadata.properties.timeZone, 'yyyy-MM-01');
    const currentMonth = Date.parse(monthStart + 'T00:00:00Z') / 86400000 + 25569;
    const expenses = _allExpensesFromSources(inputs, currentMonth);
    const layout = _allExpensesLayout(expenses);
    const requests = _allExpensesRefreshRequests(sheet, layout);
    timings.prepareMs = Date.now() - phase;
    timings.items = expenses.length;
    timings.writeRequests = requests.length;
    phase = Date.now();
    // Source validation completes before the one atomic display update.
    if (requests.length) Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
    timings.saveMs = Date.now() - phase;
    return requests.length ? `Updated ${expenses.length} expense items.` : 'All Expenses is already up to date.';
  } finally {
    lock.releaseLock();
    console.log('All Expenses refresh timings (ms): ' + JSON.stringify({...timings, totalMs: Date.now() - started}));
  }
}

function _allExpensesRefreshRequests(sheet, layout) {
  const previous = sheet.data?.[0]?.rowData?.slice(1) || [];
  const filterRange = sheet.basicFilter?.range;
  const filterMatches = layout.totalRow === 2 ? !sheet.basicFilter : filterRange &&
    (filterRange.startRowIndex || 0) === 0 && filterRange.endRowIndex === layout.totalRow - 1 &&
    (filterRange.startColumnIndex || 0) === 0 && filterRange.endColumnIndex === 6;
  const sameLayout = !sheet.merges?.length && !sheet.rowGroups?.length && filterMatches &&
    sheet.properties.gridProperties.rowCount === layout.totalRow && previous.length === layout.rows.length &&
    layout.rows.every((row, index) => row.every((value, column) => {
      if (column === 5 && index !== layout.rows.length - 1) return true;
      const old = previous[index].values?.[column]?.userEnteredValue || {};
      return value === (old.numberValue ?? old.stringValue ?? old.formulaValue ?? '');
    }));
  if (sameLayout && layout.rows.slice(0, -1).every((row, index) => {
    const old = previous[index].values || [];
    const normalize = runs => (runs || []).map(run => [run.startIndex || 0, run.format?.fontSize]);
    return old[5]?.userEnteredValue?.numberValue === row[5] &&
      JSON.stringify(normalize(old[4]?.textFormatRuns)) === JSON.stringify(normalize(_allExpensesTextRuns(String(row[4]))));
  })) return [];
  const requests = _allExpensesWriteRequests(sheet, layout);
  // Labels and periods are unchanged, so retain the user's existing formatting.
  return sameLayout ? requests.filter(request => !request.updateBorders && !request.repeatCell?.fields.includes('userEnteredFormat')) : requests;
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
  for (const type of ALL_EXPENSES_TYPES) {
    const items = expenses.filter(row => row[0] === type)
      .sort((a, b) => Number(a[3]) - Number(b[3]) || String(a[4]).localeCompare(String(b[4])));
    rows.push(...items);
  }
  const totalRow = rows.length + 2;
  rows.push(['Total', '', '', '', '', totalRow === 2 ? '=0' : `=SUBTOTAL(9,F2:F${totalRow - 1})`]);
  return { rows, totalRow };
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
  // Clearing the formula anchor releases its spill before writing static values.
  requests.push({repeatCell: {range: {sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 1},
    cell: {}, fields: 'userEnteredValue,textFormatRuns'}});
  const displayRange = {...bodyRange, endRowIndex: layout.totalRow};
  const format = {backgroundColor: {red: 1, green: 1, blue: 1}, verticalAlignment: 'MIDDLE',
    horizontalAlignment: 'LEFT', wrapStrategy: 'OVERFLOW_CELL',
    textFormat: {fontFamily: 'Arial', fontSize: 11, bold: false, foregroundColor: {red: 0.4, green: 0.4, blue: 0.4}}};
  requests.push({repeatCell: {range: displayRange, cell: {userEnteredFormat: format}, fields: 'userEnteredFormat'}});
  const columns = [
    [1, {horizontalAlignment: 'CENTER', numberFormat: {type: 'NUMBER', pattern: '0'}}],
    [2, {horizontalAlignment: 'CENTER', numberFormat: {type: 'TEXT', pattern: '@'}}],
    [3, {horizontalAlignment: 'CENTER', numberFormat: {type: 'DATE', pattern: 'mmmm'}}],
    [4, {wrapStrategy: 'WRAP'}],
    [5, {horizontalAlignment: 'RIGHT', numberFormat: {type: 'NUMBER', pattern: '฿ #,##0,"K"'}}],
  ];
  for (const [column, columnFormat] of columns) {
    requests.push({repeatCell: {range: {...displayRange, startColumnIndex: column, endColumnIndex: column + 1},
      cell: {userEnteredFormat: columnFormat}, fields: Object.keys(columnFormat).map(key => 'userEnteredFormat.' + key).join(',')}});
  }
  requests.push({repeatCell: {range: {...displayRange, startRowIndex: layout.totalRow - 1},
    cell: {userEnteredFormat: {backgroundColor: {red: 0.93, green: 0.93, blue: 0.93}, textFormat: {bold: true}}},
    fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold'}});
  const rows = layout.rows.map((row,index) => ({values: row.map((value,column) => {
    const summary = index + 2 === layout.totalRow;
    const cell = {userEnteredValue: typeof value === 'number' ? {numberValue: value}
      : column === 5 && summary ? {formulaValue: value} : {stringValue: String(value)}};
    if (column === 4 && value) cell.textFormatRuns = _allExpensesTextRuns(String(value));
    return cell;
  })}));
  requests.push({updateCells: {range: displayRange, rows, fields: 'userEnteredValue,textFormatRuns'}});
  const borderRows = _detailsBorderRows([['Type', 'Year'], ...layout.rows]);
  for (const border of borderRows) {
    requests.push({updateBorders: {range: {sheetId, startRowIndex: border.row - 1, endRowIndex: border.row, startColumnIndex: 0, endColumnIndex: 6},
      bottom: {style: border.style, color: {red: 0.4, green: 0.4, blue: 0.4}}}});
  }
  requests.push({updateDimensionProperties: {range: {sheetId, dimension: 'ROWS', startIndex: 1, endIndex: layout.totalRow},
    properties: {hiddenByUser: false}, fields: 'hiddenByUser'}});
  if (layout.totalRow > 2) {
    const filterSpecs = sheet.basicFilter?.filterSpecs || Object.entries(sheet.basicFilter?.criteria || {})
      .map(([column, filterCriteria]) => ({columnIndex: Number(column), filterCriteria}));
    requests.push({setBasicFilter: {filter: {range: {sheetId, startRowIndex: 0, endRowIndex: layout.totalRow - 1, startColumnIndex: 0, endColumnIndex: 6}, filterSpecs}}});
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
