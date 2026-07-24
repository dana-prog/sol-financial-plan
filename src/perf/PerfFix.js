/**
 * PerfFix — vectorizes duplicate 32-quarter row formulas in Main Timeline
 * so they use the MAIN_TIMELINE_PARAM_ROW named function instead of 32
 * separate MAIN_TIMELINE_PARAM(..., QUARTER()) calls.
 *
 * Prerequisite: the named function MAIN_TIMELINE_PARAM_ROW must exist.
 * Add it manually via Data → Named functions before running:
 *
 *   name:        MAIN_TIMELINE_PARAM_ROW
 *   arguments:   name
 *   definition:
 *     LAMBDA(name, LET(
 *       range, S_MAIN_TIMELINE,
 *       params, INDEX(range,,2),
 *       param_row, MATCH(name, params, 0),
 *       row_values, INDEX(range, param_row,),
 *       quarter_count, COLUMNS(range) - 2,
 *       CHOOSECOLS(row_values, SEQUENCE(1, quarter_count, 3, 1))
 *     ))
 *
 * Usage:
 *   1. Add the named function as described above.
 *   2. SOL → PerfFix → Vectorize Interest Cascade (Dry Run)   — shows what would change
 *   3. SOL → PerfFix → Vectorize Interest Cascade (Apply)     — actually applies the change
 *
 * The Apply action creates a backup sheet (PerfFix_Backup_<ts>) containing the original
 * formulas for the affected rows, so you can manually revert if needed.
 */

// Batch 1: pure-arithmetic rows. No IF, no PHASE, no pre-existing inconsistencies.
// ARRAYFORMULA is required to make + and - broadcast across the 1x32 rows
// returned by MAIN_TIMELINE_PARAM_ROW. Without it, only the first cell receives a value.
const _PERF_FIX_BATCH_1 = {
  105: '=ARRAYFORMULA(MAIN_TIMELINE_PARAM_ROW("principal opening") + MAIN_TIMELINE_PARAM_ROW("loan draw") - MAIN_TIMELINE_PARAM_ROW("principal paid"))',
  109: '=ARRAYFORMULA(MAIN_TIMELINE_PARAM_ROW("accrued interest opening") + MAIN_TIMELINE_PARAM_ROW("interest generated") - MAIN_TIMELINE_PARAM_ROW("interest paid"))',
  111: '=ARRAYFORMULA(MAIN_TIMELINE_PARAM_ROW("cash opening") - MAIN_TIMELINE_PARAM_ROW("total setup cost") + MAIN_TIMELINE_PARAM_ROW("gross profit") + MAIN_TIMELINE_PARAM_ROW("loan draw"))',
  113: '=ARRAYFORMULA(MAIN_TIMELINE_PARAM_ROW("cash before debt service") - MAIN_TIMELINE_PARAM_ROW("interest paid") - MAIN_TIMELINE_PARAM_ROW("principal paid") - MAIN_TIMELINE_PARAM_ROW("company tax"))',
};

const _PERF_FIX_SHEET_NAME = 'Main Timeline';
const _PERF_FIX_FIRST_QUARTER_COL = 3;   // C
const _PERF_FIX_LAST_QUARTER_COL = 34;   // AH

// noinspection JSUnusedGlobalSymbols
function vectorizeInterestCascadeBatch1DryRun() {
  _perfFixRun(_PERF_FIX_BATCH_1, true);
}

// noinspection JSUnusedGlobalSymbols
function vectorizeInterestCascadeBatch1Apply() {
  _perfFixRun(_PERF_FIX_BATCH_1, false);
}

function _perfFixRun(rowMap, dryRun) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(_PERF_FIX_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found: ' + _PERF_FIX_SHEET_NAME);
  }

  const ui = SpreadsheetApp.getUi();
  const log = [];
  log.push(dryRun ? '=== DRY RUN — no changes will be made ===' : '=== APPLYING CHANGES ===');
  log.push('');

  const preflight = _perfFixVerifyNamedFunction(sheet);
  if (!preflight.ok) {
    log.push('ABORT: MAIN_TIMELINE_PARAM_ROW preflight failed.');
    log.push('  reason: ' + preflight.error);
    log.push('');
    log.push('Please add the MAIN_TIMELINE_PARAM_ROW named function via Data → Named functions');
    log.push('before running this script. See PerfFix.js header for the definition.');
    ui.alert('PerfFix — Vectorize Interest Cascade', log.join('\n'), ui.ButtonSet.OK);
    return;
  }
  log.push('Preflight OK: MAIN_TIMELINE_PARAM_ROW works.');
  log.push('');

  const rows = Object.keys(rowMap).map(function (r) { return parseInt(r, 10); }).sort(function (a, b) { return a - b; });

  if (!dryRun) {
    const backupName = 'PerfFix_Backup_' + new Date().getTime();
    const backup = ss.insertSheet(backupName);
    backup.getRange(1, 1, 1, 3).setValues([['row', 'cell', 'original_formula']]);
    let brow = 2;
    rows.forEach(function (row) {
      const range = sheet.getRange(
        row,
        _PERF_FIX_FIRST_QUARTER_COL,
        1,
        _PERF_FIX_LAST_QUARTER_COL - _PERF_FIX_FIRST_QUARTER_COL + 1
      );
      const formulas = range.getFormulas()[0];
      for (let i = 0; i < formulas.length; i++) {
        if (formulas[i]) {
          const colLetter = _perfFixColLetter(_PERF_FIX_FIRST_QUARTER_COL + i);
          backup.getRange(brow, 1, 1, 3).setValues([[row, colLetter + row, formulas[i]]]);
          brow++;
        }
      }
    });
    log.push('Backup sheet created: ' + backupName);
    log.push('');
  }

  rows.forEach(function (row) {
    const newFormula = rowMap[row];
    const cRange = sheet.getRange(row, _PERF_FIX_FIRST_QUARTER_COL);
    const currentFormula = cRange.getFormula();

    log.push('Row ' + row + ':');
    log.push('  current C' + row + ': ' + _perfFixTruncate(currentFormula, 90));
    log.push('  new     C' + row + ': ' + _perfFixTruncate(newFormula, 90));

    if (!dryRun) {
      const clearRange = sheet.getRange(
        row,
        _PERF_FIX_FIRST_QUARTER_COL + 1,
        1,
        _PERF_FIX_LAST_QUARTER_COL - _PERF_FIX_FIRST_QUARTER_COL
      );
      clearRange.clearContent();
      cRange.setFormula(newFormula);
      log.push('  status: applied');
    }
    log.push('');
  });

  log.push('Done. Rows processed: ' + rows.length);
  if (!dryRun) {
    log.push('');
    log.push('Next: reload the spreadsheet, spot-check values in the affected rows');
    log.push('against the values in the backup sheet, then time a recalc by editing');
    log.push('a Params cell (e.g. quarterly interest rate).');
  }
  ui.alert('PerfFix — Vectorize Interest Cascade', log.join('\n'), ui.ButtonSet.OK);
}

/**
 * Writes a test formula calling MAIN_TIMELINE_PARAM_ROW to a scratch cell in the
 * bottom-right corner, reads the result, and restores the cell. Returns ok=true
 * iff the formula evaluates to a numeric value.
 */
function _perfFixVerifyNamedFunction(sheet) {
  const scratchRow = sheet.getMaxRows();
  const scratchCol = sheet.getMaxColumns();
  const scratch = sheet.getRange(scratchRow, scratchCol);
  const originalFormula = scratch.getFormula();
  const originalValue = scratch.getValue();
  try {
    scratch.setFormula('=IFERROR(COUNTA(MAIN_TIMELINE_PARAM_ROW("interest paid")), -1)');
    SpreadsheetApp.flush();
    const value = scratch.getValue();
    if (typeof value === 'number' && value > 0) {
      return { ok: true };
    }
    return { ok: false, error: 'test formula returned: ' + value };
  } catch (e) {
    return { ok: false, error: String(e) };
  } finally {
    if (originalFormula) {
      scratch.setFormula(originalFormula);
    } else if (originalValue !== '' && originalValue !== null && originalValue !== undefined) {
      scratch.setValue(originalValue);
    } else {
      scratch.clearContent();
    }
  }
}

function _perfFixColLetter(idx) {
  let label = '';
  let rem = idx;
  while (rem > 0) {
    const mod = (rem - 1) % 26;
    label = String.fromCharCode(65 + mod) + label;
    rem = Math.floor((rem - 1) / 26);
  }
  return label;
}

function _perfFixTruncate(str, n) {
  if (!str) return '';
  const flat = str.replace(/\s+/g, ' ');
  return flat.length > n ? flat.substring(0, n) + '…' : flat;
}
