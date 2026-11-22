function setDetailsSheetBorders(sheet) {
  const spreadsheetId = sheet.getParent().getId();
  const sheetName = sheet.getName().replace(/'/g, "''");
  const result = Sheets.Spreadsheets.get(spreadsheetId, {
    ranges: [`'${sheetName}'`],
    fields: 'sheets(properties(sheetId),merges,data(rowData(values(formattedValue)),rowMetadata(hiddenByFilter,hiddenByUser)))',
  });
  const requests = _detailsBorderRequests(result.sheets[0]);
  if (requests.length) Sheets.Spreadsheets.batchUpdate({requests}, spreadsheetId);
}

function _detailsBorderRequests(sheet) {
  const grid = sheet.data?.[0];
  const values = (grid?.rowData || []).map(row => (row.values || []).map(cell => cell.formattedValue || ''));
  while (values.length && !values[values.length - 1].some(value => value !== '')) values.pop();
  if (values.length < 2) return [];

  const sheetId = sheet.properties.sheetId;
  const lastColumn = Math.max(...values.map(row => row.length));
  const hiddenRows = new Set();
  for (let row = 1; row <= values.length; row++) {
    const metadata = grid.rowMetadata?.[row - 1];
    if (metadata?.hiddenByFilter || metadata?.hiddenByUser) hiddenRows.add(row);
  }
  // Expand labels in memory; never unmerge or rewrite the actual cells.
  for (const merge of sheet.merges || []) {
    if (merge.endColumnIndex - merge.startColumnIndex !== 1) continue;
    const value = values[merge.startRowIndex]?.[merge.startColumnIndex] || '';
    for (let row = merge.startRowIndex + 1; row < Math.min(merge.endRowIndex, values.length); row++) {
      values[row][merge.startColumnIndex] = value;
    }
  }
  const borders = _detailsBorderRows(values, hiddenRows);
  const requests = [{updateBorders: {
    range: {sheetId, startRowIndex: 1, endRowIndex: values.length, startColumnIndex: 0, endColumnIndex: lastColumn},
    top: {style: 'NONE'}, bottom: {style: 'NONE'}, innerHorizontal: {style: 'NONE'},
  }}];
  const visibleRows = values.slice(1).map((_, index) => index + 2).filter(row => !hiddenRows.has(row));
  const color = {red: 0.4, green: 0.4, blue: 0.4};

  function addEdge(row, edge, style) {
    let previous;
    for (let column = 0; column < lastColumn; column++) {
      const merge = (sheet.merges || []).find(range =>
        row - 1 >= range.startRowIndex && row - 1 < range.endRowIndex &&
        column >= range.startColumnIndex && column < range.endColumnIndex);
      let range = {sheetId, startRowIndex: row - 1, endRowIndex: row, startColumnIndex: column, endColumnIndex: column + 1};
      if (merge) {
        const visibleMergedRows = visibleRows.filter(index => index > merge.startRowIndex && index <= merge.endRowIndex);
        // A year divider must not cut through the merged expense type label.
        if (row !== (edge === 'top' ? visibleMergedRows[0] : visibleMergedRows.at(-1))) continue;
        if (column !== merge.startColumnIndex) continue;
        range = {...merge};
      }
      // Join adjacent cells only when their full vertical extents match.
      if (previous && previous.range.startRowIndex === range.startRowIndex &&
          previous.range.endRowIndex === range.endRowIndex && previous.range.endColumnIndex === range.startColumnIndex) {
        previous.range.endColumnIndex = range.endColumnIndex;
      } else {
        previous = {range, [edge]: {style, color}};
        requests.push({updateBorders: previous});
      }
    }
  }

  if (visibleRows.length) addEdge(visibleRows[0], 'top', 'SOLID_THICK');
  for (const border of borders) {
    addEdge(border.row, 'bottom', border.style);
    const nextRow = visibleRows.find(row => row > border.row);
    if (nextRow) addEdge(nextRow, 'top', border.style);
  }
  return requests;
}

function _detailsBorderRows(values, hiddenRows = new Set()) {
  const headers = values[0].map(value => String(value).trim().toLowerCase());
  const category = ['budget category', 'type', 'expense type'].map(name => headers.indexOf(name))
    .find(index => index >= 0) ?? -1;
  const columns = [category, headers.indexOf('year')];
  if (columns.includes(-1)) {
    throw new Error('Could not find Budget Category (or Type) and Year columns. No borders were changed.');
  }
  const [group, year] = columns;
  const rows = values.slice(1).map((values, index) => ({ values, row: index + 2 }))
    .filter(row => row.values[group] !== '' && row.values[group] !== undefined && !hiddenRows.has(row.row));
  const borders = [];
  for (let index = 0; index < rows.length; index++) {
    const current = rows[index];
    const next = rows[index + 1];
    // A shared boundary receives only its strongest divider.
    if (!next || current.values[group] !== next.values[group]) {
      borders.push({ row: current.row, style: 'SOLID_THICK' });
    } else if (next.values[year] !== '' && current.values[year] !== next.values[year]) {
      borders.push({ row: current.row, style: 'SOLID' });
    }
  }
  return borders;
}
