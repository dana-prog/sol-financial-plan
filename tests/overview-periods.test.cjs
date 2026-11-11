const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const context = vm.createContext({});
vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/OverviewPeriods.js'), 'utf8'), context);

function edit(row, column, lastRow = row, lastColumn = column, sheetId = 190902611) {
  const writes = [];
  return {writes, event: {range: {
    getGridId: () => sheetId,
    getRow: () => row,
    getLastRow: () => lastRow,
    getColumn: () => column,
    getLastColumn: () => lastColumn,
    getSheet: () => ({getRange: cell => ({setValue: value => writes.push({cell, value})})})
  }}};
}

test('changing the period type resets the selected period to All', () => {
  for (const fixture of [edit(11, 2), edit(11, 2, 12, 4), edit(10, 1, 12, 8)]) {
    assert.equal(context.handleOverviewPeriodEdit(fixture.event), true);
    assert.deepEqual(fixture.writes, [{cell: 'F11', value: 'All'}]);
  }
});

test('period selections, unrelated cells and other sheets are left unchanged', () => {
  for (const fixture of [edit(11, 6), edit(10, 2), edit(12, 2), edit(11, 1), edit(11, 3), edit(11, 2, 11, 2, 706577296)]) {
    assert.equal(context.handleOverviewPeriodEdit(fixture.event), false);
    assert.deepEqual(fixture.writes, []);
  }
  assert.equal(context.handleOverviewPeriodEdit(), false);
  assert.equal(context.handleOverviewPeriodEdit({}), false);
});
