function setDetailsSheetBorders(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2) return;

  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  const values = range.getDisplayValues();
  for (const merge of range.getMergedRanges()) {
    const row = merge.getRow() - 1;
    const column = merge.getColumn() - 1;
    // Merged labels are stored only in their top cell, but borders still need
    // the category and year associated with every detail and subtotal row.
    if (merge.getNumColumns() === 1) {
      for (let index = row + 1; index < row + merge.getNumRows(); index++) values[index][column] = values[row][column];
    }
  }
  const borders = _detailsBorderRows(values);

  // Rebuild horizontal dividers without changing the header or vertical borders.
  sheet.getRange(2, 1, lastRow - 1, lastColumn)
    .setBorder(null, null, false, null, null, false);

  const lastColumnName = sheet.getRange(1, lastColumn).getA1Notation().replace(/\d+$/, '');
  const styles = SpreadsheetApp.BorderStyle;
  for (const style of ['SOLID_THICK', 'SOLID']) {
    const ranges = borders.filter(border => border.style === style)
      .map(border => `A${border.row}:${lastColumnName}${border.row}`);
    if (ranges.length) {
      sheet.getRangeList(ranges)
        .setBorder(null, null, true, null, null, null, BORDER_COLOR, styles[style]);
    }
  }
}

function _detailsBorderRows(values) {
  const headers = values[0].map(value => String(value).trim().toLowerCase());
  const columns = [headers.indexOf('type') >= 0 ? headers.indexOf('type') : headers.indexOf('expense type'), headers.indexOf('year')];
  if (columns.includes(-1)) {
    throw new Error('Could not find Type and Year columns. No borders were changed.');
  }
  const [type, year] = columns;
  const rows = values.slice(1).map((values, index) => ({ values, row: index + 2 }))
    .filter(row => row.values[type] !== '');
  const borders = [];
  for (let index = 0; index < rows.length; index++) {
    const current = rows[index];
    const next = rows[index + 1];
    // A shared boundary receives only its strongest divider.
    if (!next || current.values[type] !== next.values[type]) {
      borders.push({ row: current.row, style: 'SOLID_THICK' });
    } else if (next.values[year] !== '' && current.values[year] !== next.values[year]) {
      borders.push({ row: current.row, style: 'SOLID' });
    }
  }
  return borders;
}
