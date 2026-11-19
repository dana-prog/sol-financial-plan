function setDetailsSheetBorders(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2) return;

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
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
  const columns = ['type', 'year'].map(header => headers.indexOf(header));
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
    } else if (current.values[year] !== next.values[year]) {
      borders.push({ row: current.row, style: 'SOLID' });
    }
  }
  return borders;
}
