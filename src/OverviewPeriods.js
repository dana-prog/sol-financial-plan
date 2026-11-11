// The dependent dropdown is formula-driven; reset its selection when its type changes.
function handleOverviewPeriodEdit(e) {
  if (!e || !e.range || e.range.getGridId() !== 190902611) return false;

  const range = e.range;
  const typeRow = 11;
  const typeColumn = 2;
  if (range.getRow() > typeRow || range.getLastRow() < typeRow ||
      range.getColumn() > typeColumn || range.getLastColumn() < typeColumn) return false;

  range.getSheet().getRange('F11').setValue('All');
  return true;
}
