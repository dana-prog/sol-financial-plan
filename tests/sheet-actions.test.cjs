const fs = require("node:fs"),
  vm = require("node:vm");
const assert = require("node:assert/strict");
const test = require("node:test");
const root = require("node:path").join(__dirname, "../src");
function evalFiles(names, globals = {}) {
  const ctx = { console, ...globals };
  vm.createContext(ctx);
  for (const name of names)
    vm.runInContext(
      fs.readFileSync(
        root + "/" + name.replace(/\.gs$/, ".js").replace(/^library\//, ""),
        "utf8",
      ),
      ctx,
    );
  return ctx;
}

test("timeline label range includes final row", () => {
  let args;
  const ctx = evalFiles([
    "Constants.gs",
    "Utils.gs",
    "syncs/TimelineParamsSync.gs",
  ]);
  ctx._getParamsRange({ getLastRow: () => 37, getRange: (...a) => (args = a) });
  assert.equal(args[0] + args[2] - 1, 37);
});
test("blank and bulk rename events report skipped sync without writes", () => {
  let warnings = 0,
    writes = 0;
  const ctx = evalFiles(
    ["Constants.gs", "Utils.gs", "syncs/TimelineParamsSync.gs"],
    {
      SOLLibrary: { logArgs() {} },
      SpreadsheetApp: {
        getActiveSpreadsheet: () => ({
          toast: () => warnings++,
          getSheetById: () => {
            writes++;
          },
        }),
      },
    },
  );
  ctx._findSyncMapping = () => ({ target: { sheetId: 1, categories: [] } });
  const range = {
    getSheet: () => ({ getSheetId: () => 1 }),
    getColumn: () => 1,
    getNumRows: () => 1,
    getNumColumns: () => 1,
  };
  ctx.updateTimelineParamNames({ range, oldValue: "Chef" });
  ctx.updateTimelineParamNames({ range: { ...range, getNumRows: () => 2 } });
  assert.equal(warnings, 2);
  assert.equal(writes, 0);
});
test("non-timeline border action fails before any format write", () => {
  const actions = [];
  const sh = {
    getLastRow: () => 10,
    getLastColumn: () => 2,
    getRange: () => ({
      setBorder() {
        actions.push("clear");
      },
      getValues: () => [["Name", "Value"]],
    }),
  };
  const ctx = evalFiles(["TimelineSheetBorders.gs"], {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getActiveSheet: () => sh }),
    },
  });
  assert.throws(() => ctx.setTimelineSheetBorders(), /No borders were changed/);
  assert.equal(actions.length, 0);
});
