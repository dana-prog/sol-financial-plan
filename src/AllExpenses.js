const ALL_EXPENSES_TYPES = [
  'Borrowing Repayments', 'Loan Disbursements', 'Setup Expenses', 'Land Expenses',
  'Infrastructure Expenses', 'Construction Expenses', 'Direct Operating Expenses',
  'Other Operating Expenses',
];
const ALL_EXPENSES_CALCULATION_ROWS = 20000;

function refreshAllExpenses() {
  const started = Date.now();
  const timings = {};
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) throw new Error('Another refresh is running. Please try again shortly.');
  let snapshot;
  let calculating = false;
  try {
    let phase = Date.now();
    const target = spreadsheet.getSheetByName('All Expenses') || spreadsheet.getSheetByName('Details');
    if (!target) throw new Error('Could not find All Expenses or Details.');
    const title = target.getName().replace(/'/g, "''");
    const response = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [`'${title}'!A1:F${target.getMaxRows()}`],
      fields: 'sheets(properties,basicFilter,rowGroups,merges,data(rowData(values(userEnteredValue,textFormatRuns))))',
    });
    snapshot = response.sheets[0];
    timings.readMs = Date.now() - phase;
    const capacity = Math.max(ALL_EXPENSES_CALCULATION_ROWS, snapshot.properties.gridProperties.rowCount);
    const calculationRange = `'${title}'!A2:F${capacity}`;
    phase = Date.now();
    const result = Sheets.Spreadsheets.batchUpdate({
      requests: _allExpensesCalculationRequests(snapshot, capacity),
      includeSpreadsheetInResponse: true, responseIncludeGridData: true, responseRanges: [calculationRange],
    }, spreadsheetId, {fields: 'updatedSpreadsheet(sheets(data(rowData(values(effectiveValue)))))'});
    calculating = true;
    timings.calculateMs = Date.now() - phase;
    const expenses = _readCalculatedExpenses(spreadsheetId, calculationRange, result, timings);
    phase = Date.now();
    const layout = _allExpensesLayout(expenses);
    const requests = _allExpensesRefreshRequests(snapshot, layout, capacity);
    timings.prepareMs = Date.now() - phase;
    timings.items = expenses.length;
    timings.writeRequests = requests.length;
    phase = Date.now();
    Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
    calculating = false;
    timings.saveMs = Date.now() - phase;
    return `Updated ${expenses.length} expense items.`;
  } catch (error) {
    if (calculating) {
      try {
        Sheets.Spreadsheets.batchUpdate({requests: _allExpensesRestoreRequests(snapshot)}, spreadsheetId);
      } catch (restoreError) {
        throw new Error(`${error.message} Could not restore the previous display: ${restoreError.message}`);
      }
    }
    throw error;
  } finally {
    lock.releaseLock();
    console.log('All Expenses refresh timings (ms): ' + JSON.stringify({...timings, totalMs: Date.now() - started}));
  }
}

function _allExpensesCalculationRequests(sheet, capacity) {
  const sheetId = sheet.properties.sheetId;
  const requests = [];
  if (sheet.basicFilter) requests.push({clearBasicFilter: {sheetId}});
  for (const range of sheet.merges || []) requests.push({unmergeCells: {range}});
  const extraRows = capacity - sheet.properties.gridProperties.rowCount;
  if (extraRows > 0) requests.push({appendDimension: {sheetId, dimension: 'ROWS', length: extraRows}});
  requests.push({repeatCell: {range: {sheetId, startRowIndex: 1, endRowIndex: capacity, startColumnIndex: 0, endColumnIndex: 6},
    cell: {}, fields: 'userEnteredValue,textFormatRuns'}});
  requests.push({updateCells: {start: {sheetId, rowIndex: 1, columnIndex: 0},
    rows: [{values: [{userEnteredValue: {formulaValue: ALL_EXPENSES_FORMULA}}]}], fields: 'userEnteredValue'}});
  return requests;
}

function _readCalculatedExpenses(spreadsheetId, range, result, timings) {
  const started = Date.now();
  const deadline = started + 120000;
  let grid = result.updatedSpreadsheet?.sheets?.[0]?.data?.[0]?.rowData || [];
  timings.polls = 0;
  while (true) {
    const error = grid[0]?.values?.[0]?.effectiveValue?.errorValue;
    if (error && error.type !== 'LOADING') throw new Error(`Expense calculation failed: ${error.message}`);
    let rows = grid.map(row => (row.values || []).map(cell => {
      const value = cell.effectiveValue || {};
      return value.numberValue ?? value.stringValue ?? '';
    }));
    // Formatted empty rows may be included in GridData after an in-place write.
    while (rows.length && rows[rows.length - 1].every(value => value === '')) rows.pop();
    if (rows[0]?.[0] === '__NO_EXPENSES__') rows = [];
    else if (!ALL_EXPENSES_TYPES.includes(rows[0]?.[0])) {
      if (Date.now() >= deadline) throw new Error('Expense calculation did not finish.');
      if (timings.polls) Utilities.sleep(1000);
      const response = Sheets.Spreadsheets.get(spreadsheetId, {
        ranges: [range], fields: 'sheets(data(rowData(values(effectiveValue))))',
      });
      grid = response.sheets?.[0]?.data?.[0]?.rowData || [];
      timings.polls++;
      continue;
    }
    _validateAllExpenses(rows);
    timings.waitMs = Date.now() - started;
    return rows;
  }
}

function _allExpensesRefreshRequests(sheet, layout, capacity) {
  const previous = sheet.data?.[0]?.rowData?.slice(1) || [];
  const sameLayout = !sheet.merges?.length && !sheet.rowGroups?.length && sheet.basicFilter &&
    sheet.properties.gridProperties.rowCount === layout.totalRow && previous.length === layout.rows.length &&
    layout.rows.every((row, index) => row.every((value, column) => {
      if (column === 5 && index !== layout.rows.length - 1) return true;
      const old = previous[index].values?.[column]?.userEnteredValue || {};
      return value === (old.numberValue ?? old.stringValue ?? old.formulaValue ?? '');
    }));
  const preparedSheet = {...sheet, merges: [], basicFilter: undefined,
    properties: {...sheet.properties, gridProperties: {...sheet.properties.gridProperties, rowCount: capacity}}};
  const requests = _allExpensesWriteRequests(preparedSheet, layout);
  const filter = requests.find(request => request.setBasicFilter)?.setBasicFilter.filter;
  if (filter && sheet.basicFilter) {
    filter.filterSpecs = sheet.basicFilter.filterSpecs || Object.entries(sheet.basicFilter.criteria || {})
      .map(([column, filterCriteria]) => ({columnIndex: Number(column), filterCriteria}));
  }
  // Labels and periods are unchanged, so retain the user's existing formatting.
  return sameLayout ? requests.filter(request => !request.updateBorders && !request.repeatCell?.fields.includes('userEnteredFormat')) : requests;
}

function _allExpensesRestoreRequests(sheet) {
  const sheetId = sheet.properties.sheetId;
  const rowCount = sheet.properties.gridProperties.rowCount;
  const requests = [
    {updateCells: {start: {sheetId, rowIndex: 1, columnIndex: 0}, rows: [{values: [{}]}], fields: 'userEnteredValue'}},
    {updateCells: {range: {sheetId, startRowIndex: 1, endRowIndex: rowCount, startColumnIndex: 0, endColumnIndex: 6},
      rows: (sheet.data?.[0]?.rowData || []).slice(1), fields: 'userEnteredValue,textFormatRuns'}},
    {updateSheetProperties: {properties: {sheetId, gridProperties: {rowCount}}, fields: 'gridProperties.rowCount'}},
  ];
  for (const range of sheet.merges || []) requests.push({mergeCells: {range, mergeType: 'MERGE_ALL'}});
  if (sheet.basicFilter) requests.push({setBasicFilter: {filter: sheet.basicFilter}});
  return requests;
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
    [3, {horizontalAlignment: 'CENTER', numberFormat: {type: 'DATE', pattern: 'mmmm yyyy'}}],
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
