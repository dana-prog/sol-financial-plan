# Expense source refactor

Budget Items (formerly All Expenses) is now Budget's source for ordinary expense amounts and category contingency. Company Tax remains a Cash Flow result displayed under Operations; revenue and funding retain their original sources. Cash Flow remains on its original expense sources; migrating that consumer is outside this restoration.

Current names: the sheet anchor is `=BUDGET_ITEMS_ROWS(R_MONTHS)` and Budget reads the `BUDGET_ITEMS` named range through `BUDGET_ITEMS_SUM(status, amount_column)`. The Actual/Planned column is named `Status`. Earlier names in the migration history below describe the same generator and dataset.

The baseline is the SOL financial plan on 22 November 2026, after the Road correction and the user's numeric Water Tanks correction. The spreadsheet timezone is Asia/Bangkok. Cash Flow covers January 2026 through December 2033; Budget currently covers 2026 through 2029. Do not extend Budget's horizon as part of this migration.

## Reconciliation results

Evidence: [original baseline and formula mapping](../migrations/2026-11-22-expense-source-refactor-baseline.json) and [applied dataset upgrade, refreshed reconciliation and rollback](../migrations/2026-11-22-all-expenses-dataset.json). The original baseline remains an immutable record from before the Water Tanks correction.

- Independently reconstructed 1,064 nonzero expense records from source values, using actual Setup before November 2026 and planned Setup thereafter.
- All eight Cash Flow expense types reconcile over all 96 months: 768 comparisons, no difference greater than THB 0.01.
- All 86 Budget detail lines reconcile for both past and future amounts: 172 comparisons, no difference greater than THB 0.01. Company Tax is compared to its existing Cash Flow source and remains outside the proposed expense dataset.
- No nonzero source expense within Budget's existing years is missing a matching Budget line after the Road correction. Out Loan Disbursements and out loans disbursements require an explicit alias; do not rely on singular/plural display labels as keys.
- No negative expense records were present in the baseline. Live validation temporarily set Water Tanks to -96,000, zero, nonnumeric text and a division error: the signed total was -110,400; zero produced no record; both invalid inputs surfaced errors. The original numeric 96,000 was restored and the reports rechecked.
- The live expanded dataset was checked record by record against independently reconstructed source amounts, categories, basis and provenance. Budget and Cash Flow formulas and effective values remained unchanged after validation. This does not establish a performance improvement or validate the future report-consumer formulas.

## Corrections and source exceptions

### Budget consumer migration completed

The migration was applied on 22 November, [reversed for Budget](../migrations/2026-11-23-restore-budget-source-formulas.json) on 23 November, and then [reapplied at the user's request](../migrations/2026-11-23-reapply-budget-all-expenses.json) the same day. Before reapplication, all 189 formulas matched the saved original versions; afterward, each exactly matched the saved migration. Every Budget value remained unchanged within THB 0.000001, with no formula errors or unintended formula, formatting, validation or link changes. The native Budget display was checked at 100% zoom. Existing named-function definitions were reused; no Apps Script changes or clasp synchronization were required.

`BUDGET_EXPENSE_SUM(basis, amount_column)` reads the existing `ALL_EXPENSES` named range using the current Budget row's year, category and line. Blank line means the whole category; blank basis means both bases. It normalizes the existing plural Out Loans label, returns zero for missing matches and propagates source errors. Its native definition is maintained in `src/BudgetFormula.js`.

All 83 ordinary detail lines now use that function for Cost in Past Expenses (Actual) and Future Expenses (Planned). The 23 category contingency cells read Contingency from the dataset, without applying rates again. The three Company Tax lines retain their existing formulas. Existing subtotal functions, source drill-down links, funding, carry-forward, headroom, statuses and sheet layout are preserved. Source ranges remain in link formulas only to locate drill-down destinations, not to retrieve expense amounts.

`ALL_EXPENSES_ROWS` explicitly accepts only Actual records before the current month and Planned records from the current month onward. It never relabels past plans as actuals or substitutes a plan for a missing actual. All 1,064 existing records already satisfied this rule, so no current amount changed. Setup is currently the only actual expense source; the other sources will need actual records as their months become historical.

Validation compared every Budget result, all expense records and the complete Cash Flow snapshot before and after (numeric tolerance THB 0.000001). Twelve native Sheets formula cases cover signed credits, duplicate records, basis/year matching, category contingency, aliases, empty results, source errors and the month boundary; thirteen existing local tests passed. Native function exports match local source. See [migration and rollback](../migrations/2026-11-22-budget-all-expenses.json).

### Budget recalculation optimization — 23 November 2026

`BUDGET_EXPENSE_SUM` now finds populated rows from the category column, then selects each requested column before calling the existing `COL_VALUES`. The 10,272-row source capacity remains available for growth, but the five matching/amount arrays only contain populated records plus the first data row. Retaining that row lets the existing column function handle an empty dataset. Blank gaps, later-added records, signed amounts and source errors remain supported. Unknown amount headers still surface an error, although the column-selection path may return a different error type.

The 166 ordinary Past/Future Expense formulas now return zero before the expense lookup when their period contains no eligible months. Currently this avoids work for 78 future-year Past Expense cells. All named-function calls remain in place, including source drill-down links. Company Tax and the All Expenses generator are unchanged.

The C17 clear test settled in 21.716 seconds with the original formulas and 5.219 seconds with the final Budget-only optimization; restoration took 5.266 seconds. Timing used the visible Sheets recalculation bar followed by 2.5 seconds of quiet. This small sample includes browser/network overhead and is not a general performance guarantee. A separate loan-detail named range showed no clear additional benefit and was removed.

Fourteen native formula cases and thirteen related local tests passed. Every Budget result, Cash Flow value and all 1,064 expense records matched the pre-change snapshot within THB 0.000001. Budget formatting, validation and links were unchanged. The native Budget definition matches local source, which was synchronized through clasp. See the [performance migration and rollback record](../migrations/2026-11-23-budget-expense-performance.json).

### Individual Budget formulas restored — 23 November 2026

At the user's request after reporting slower performance, Budget again uses individual cell formulas for details, category/year totals, contingency, funding, headroom and statuses. The 189 expense lookups still call `BUDGET_EXPENSE_SUM` against the existing `ALL_EXPENSES` named range. This restores the optimized version immediately before the single-formula experiment, not the older direct-source expense formulas. Company Tax and funding retain their existing named-function sources.

The 27 label merges and original C/D widths are restored. M53 and M88 remain blank as previously approved. J13 retains its zero Required Budget subtotal through `BUDGET_VALUE("total expenses") + BUDGET_VALUE("contingency")`, for 463 formulas in total. All Budget values match the live pre-restoration snapshot within THB 0.000001, and the 110 source links match the original individual-formula baseline. Cash Flow and All Expenses formulas and values are unchanged. Native visual verification completed at 100% zoom. No new performance claim is made.

The unused `BUDGET_ROWS` definitions remain available for historical reference; Budget no longer calls them. No Apps Script or named-function definitions changed, so no clasp synchronization was required. See the [restoration record](../migrations/2026-11-23-budget-restore-cell-formulas.json).

### Single-formula Budget experiment — 23 November 2026 (subsequently reversed)

During this experiment, Budget A1 contained `=BUDGET_ROWS()`, generating the existing 114 × 13 report. Native definitions are maintained in `src/BudgetRowsFormula.js`. The individual Budget cell formulas have since been restored as described above.

`BUDGET_EXPENSE_TOTALS(ALL_EXPENSES)` groups the existing named range once by year, category and normalized line, with separate Actual Cost, Planned Cost and Contingency amounts. It does not call `ALL_EXPENSES_ROWS`. `BUDGET_REPORT_AMOUNT` reads this compact summary. `BUDGET_ROW_LAYOUT()` preserves the current 2026–2029 catalog, including zero-cost lines. New categories, lines or years require updating this catalog and the sheet formatting/range together.

`BUDGET_ROWS()` builds detail amounts and subtotals, adds Company Tax from the existing Cash Flow named function, and calculates funding and cumulative headroom using the existing funding/revenue named functions. Company Tax remains under Operations with no contingency. The output creates links through `BUDGET_SOURCE_LINK` and the existing `BUDGET_TIMELINE_LINK` at the final `MAKEARRAY` step: building links before stacking the arrays loses their hyperlink metadata in native Sheets.

All numeric amounts reconcile within THB 0.000001, and all 110 source links are preserved. The two literal `))` status cells (M53 and M88) are blank as approved. The previously empty J13 now shows a zero Required Budget subtotal for 2027 In Loans, without changing annual totals. The 27 label merges were removed to accommodate the array, and columns C/D widened to 100/116 pixels; formatting, groups, filter and conditional formatting are otherwise retained. Do not enter values inside the generated output, because that would block the array.

Thirteen native formula cases passed for credits, duplicates, period separation, contingency, subtotals, aliases, missing lines, empty data, gaps and error propagation; 48 related local tests passed. Every existing named range/function was unchanged, and the five new native definitions match source. Cash Flow and All Expenses values and formulas are unchanged. The final report was visually checked at 100% zoom. See the [migration and rollback record](../migrations/2026-11-23-budget-single-formula.json).

No speed improvement is claimed yet. The initial timing probe did not observe the recalculation bar, and user browser interaction interrupted the subsequent attempt, so those measurements were discarded. The single report reduces repeated All Expenses scans but still recalculates the whole report when its dependencies change; upstream All Expenses generation remains part of the dependency chain.

### Company Tax grouped under Operations in Budget

Company Tax is now the first detail line under Operations for 2027–2029 (rows 33, 68, and 103). The separate Tax category headers were removed. This is a Budget reporting change only: the detail formulas still read `CASH_FLOW_PARAM_VALUES("company tax")`, and All Expenses still excludes tax.

At the grouping stage, Operations contingency subtracted company-tax payments before applying the operating rate. The consumer migration now reads contingency directly from All Expenses, which excludes company tax. Company Tax still increases the Operations subtotal without receiving contingency. Annual expense totals, contingency, required budget, available budget, headroom, and statuses were preserved. The row groups and named Budget range adjusted to the new layout. See [grouping migration](../migrations/2026-11-22-budget-tax-under-operations.json).

### Road correction completed

Budget now includes Road in its 2027, 2028, and 2029 infrastructure sections. The 2028 amount is THB 2,000,000 plus THB 300,000 contingency. After the subsequent Water Tanks correction, the 2028 infrastructure required budget is THB 16,802,650. Annual required budget is THB 231,912,065.99853653 and headroom is THB 33,979,241.219913065. The headroom carried into 2029 leaves THB 17,028,904.4375.

The migration record is [2026-11-22-budget-road-expense.json](../migrations/2026-11-22-budget-road-expense.json), committed separately as `80ea91b`. Existing named formulas were preserved. New rows use the existing infrastructure formulas and their source links resolve to the Road row.

### Setup basis correction completed in All Expenses

All Expenses now uses actual Setup amounts before the current month and planned amounts from the current month onward. This reduced its January–September 2026 base total by THB 591,150 and brought all nine previously differing months into agreement with Budget and Cash Flow.

The current month belongs to the planned period. A missing or zero actual amount must remain missing or zero according to the existing source semantics; do not fall back to planned values. A record's basis must come from its source, not merely from its date.

### Water Tanks correction completed by the user

The user changed `Infra Exp (P)!AE9` from text `฿ 96K` to numeric 96,000 for February 2028. All Expenses now includes base cost 96,000 and contingency 14,400; Cash Flow contains total cost 110,400. No ignored nonnumeric source amount remains. Eligible nonempty text now surfaces an error instead of silently disappearing.

## Source mapping

| Expense type | Source and selection | Budget mapping | Contingency |
| --- | --- | --- | --- |
| Setup Expenses | Setup Exp (A) before the current month; Setup Exp (P) from the current month onward; exclude Totals | Preserve source `_category`: Accounting, Legal, Permits, Planning | Actual: zero. Planned: `PARAM("setup contingency")`, currently 10% |
| Construction Expenses | Construction Exp (P) parameters ending in ` const. cost`; exclude counts and summary rows | Construction / Construction | `PARAM("construction contingency")`, currently 15% |
| Infrastructure Expenses | Infra Exp (P) numeric detail amounts; exclude Totals | Infra / source `_category`, including Road | `PARAM("infra contingency")`, currently 15% |
| Direct Operating Expenses | Operational Exp (P) rows in that category | Operations / parameter mapped to its Budget line | `PARAM("operating contingency")`, currently 5% |
| Other Operating Expenses | Operational Exp (P) rows in that category | Operations / parameter mapped to its Budget line | `PARAM("operating contingency")`, currently 5% |
| Land Expenses | Obligations (P), inherited `_category` = Land | Land / Land | Zero |
| Out Loan Disbursements | Out Loans, detail parameter `out loans disbursements`; exclude Totals | Out Loans / normalized disbursement key | Zero |
| In Loan Repayments | In Loans, detail parameter `in loans repayments`; exclude Totals | In Loans / repayment key | Zero |

Operating parameter mappings are Insurance, Wifi, Software, Legal, Electricity, Water, Food, Maintenance, Wear & tear, Marketing, and Salaries. The two records `total guest services salaries` and `total operations support salaries` map to Salaries. These are the selected salary aggregates: do not additionally include individual salary rows or the combined `salaries expenses` row.

Preserve refunds and negative corrections as signed amounts. Zero-valued months may be omitted from the detail dataset but must return zero from monthly aggregations. Do not suppress source errors with blanket `IFERROR(...,0)` handling.

## All Expenses schema

The displayed dataset has 13 columns in the user's requested order. Type is retained only inside the named-function calculation for ordering and descriptions; it is no longer a sheet column. Source identifiers remain separate from display descriptions. No staging or temporary sheet is required.

| Column | Field | Meaning |
| --- | --- | --- |
| A | Budget Category | Stable category corresponding to Budget's `_category` |
| B | Budget Line | Explicit matching key, with financing aliases normalized |
| C | Param | Readable description; not an aggregation key |
| D | Basis | Actual or Planned; financing and land use Planned |
| E | Year | Derived from the real Month date |
| F | Quarter | Derived display value, Q1–Q4 |
| G | Month | First-of-month date, displayed as `mmmm` |
| H | Cost | Base cost, excluding contingency |
| I | Contingency | Signed base cost multiplied by the applicable rate; zero for actual Setup and financing/land |
| J | Total Cost | Cost plus Contingency |
| K | Source Sheet | Origin for provenance and validation |
| L | Source Parameter | Unmodified source parameter or obligation name |
| M | Source Row | Source location for traceability, not a persistent identity |

Use one record per selected source row and month. Identical display descriptions may legitimately appear more than once. Source row positions can change; regenerate provenance from the source rather than treating row numbers as permanent IDs. Keep current type/date ordering. The dataset has no embedded grand-total record. Report consumers must aggregate the appropriate amount column.

Base and contingency stay separate. Budget now reads Cost for its expense columns and Contingency for its category contingency cells. Do not apply rates a second time or feed Total Cost to Budget's base-expense columns. Cash Flow's future consumer should read Total Cost.

## Formula and refresh contract

All Expenses is formula-driven at A1, including its column-title row. Its formula is stored in `src/AllExpensesFormula.js`. The four existing detail functions (`SETUP_EXPENSES_D`, `CONSTRUCTION_EXPENSES_D`, `OPERATIONAL_EXPENSES_D`, `LAND_EXPENSES_D`) were extended to six outputs: type, parameter, party, base amount, source row and source category. The inventory found All Expenses A2 to be their only consumer. Their definitions and original versions are retained in the migration record. Existing named-function calls are preserved.

Current and proposed interfaces:

- `S_ALL_EXPENSES` is live: All Expenses A1:M10273, with a header and no total record. Its capacity covers every currently eligible source row in all 96 months. Revisit capacity when source rows or the timeline expand.
- `ALL_EXPENSES_MONTHLY_VALUES(expense_type, amount_column)`: returns one value for every month in `R_MONTHS`, including zero months.
- `BUDGET_EXPENSE_SUM(basis, amount_column)` is live: sums the selected amount for the current Budget row's year, category, optional line and optional basis.

Reuse existing header lookup functions such as `COL_VALUES` and `DATA_ROWS`. Keep `BUDGET_VALUE`, `BUDGET_CATEGORY_SUM`, `BUDGET_YEAR_SUM`, `BUDGET_YEAR_CONTINGENCY`, `BUDGET_TIMELINE_LINK`, and `CASH_FLOW_PARAM_VALUES` in their existing roles. Retain source drill-down links. Replace source-specific retrieval calls only with a reviewed named-function interface, never direct ranges or inlined implementations. Do not globally redirect source functions such as `CONSTRUCTION_PARAM_VALUES` or `OPERATIONS_PARAM_VALUES` to All Expenses: its own calculations depend on those functions.

The former `refreshAllExpenses()` validated the live anchor and headers, then updated formatting, the filter range and borders without writing expense values. The duplicate static financial calculation was removed. Filter criteria were retained, but saved sort specifications were not reapplied: a live test demonstrated that doing so moves the array anchor. Type/date ordering belongs to the formula. The corrected refresh was executed and verified to preserve A2 and all 1,064 records. The refresh command and implementation were subsequently removed as recorded below.

## Named-function simplification completed

All Expenses A1 now contains only `=ALL_EXPENSES_ROWS(R_MONTHS)`. The named definitions and the A1 formula are maintained in `src/AllExpensesFormula.js`; they are native Sheets named functions, not Apps Script custom functions. Existing source named functions and calls remain intact.

| Function | Responsibility |
| --- | --- |
| `EXPENSE_DETAIL_ROWS(details, eligible)` | Validate selected amounts, preserve errors and signed adjustments, omit zeros, and return an empty row when no records qualify |
| `INFRA_EXPENSES_D(month)` | Read infrastructure details through the existing named range and header lookup functions |
| `FINANCING_EXPENSES_D(month, source, param_name, party_header, type_label)` | Read scheduled financing details using the live source labels |
| `EXPENSE_RECORDS(expenses, month, source_name, category, basis, rate)` | Add the Budget mapping, source, basis, contingency and total |
| `EXPENSES_FOR_MONTH(month)` | Combine all expense sources for one month |
| `EXPENSE_DESCRIPTION(type, month, param, party)` | Produce the readable description, retaining `CONSTRUCTION_PARAM_VALUES` for construction quantities |
| `ALL_EXPENSES_ROWS(months)` | Return the title row followed by the current 13-column records for the supplied months; return titles only for an empty result |

The dependency graph has no reference back to `S_ALL_EXPENSES`, Budget or Cash Flow. `IN_LOANS_EXPENSES_D` was inspected but left unchanged: it is a legacy four-column function, while the current financing calculation needs the six-column source-provenance contract and the live parameter labels.

At the named-function simplification stage, all 1,064 rows were compared across the then-current 14 columns before and after, with exact equality. Budget and Cash Flow formulas and effective values were unchanged. Native named-function tests covered negative values, duplicate descriptions from separate source rows, zero omission, empty selections, signed contingency, nonnumeric amounts, source formula errors and a single-month input. See the [named-function migration record](../migrations/2026-11-22-all-expenses-named-functions.json).

### Column order and Basis update

The subsequent [column-order migration](../migrations/2026-11-22-all-expenses-column-order.json) removed the displayed Type field and applied the 13-column schema above. All 1,064 records retained the same row order, amounts and source metadata. The 106 Scheduled basis labels for financing and land were renamed Planned. The filter and `S_ALL_EXPENSES` now span A:M; field widths and formatting follow their original fields. Section borders use Budget Category and Year. An empty result still has 13 columns, and the anchor at that stage remained A2.

Budget and Cash Flow formulas and values are unchanged. Before their consumer migration, adapt aggregation interfaces to this schema; no visible Type column remains for a type-based lookup. Internal expense records still retain that classification.

### Generated title row

The title row is now part of `ALL_EXPENSES_ROWS` output, with the single call `=ALL_EXPENSES_ROWS(R_MONTHS)` in A1. Static title cells were cleared and the A2 formula removed; expense rows still begin at row 2. Title values and order are generated from `ALL_EXPENSES_HEADERS` in the local definition source. An empty result returns exactly the 13 titles, without a blank detail row. Refresh validates the live anchor at A1 and also checks for errors there. The existing filter and named range still begin at row 1.

All titles, 1,064 expense records, and Budget/Cash Flow values were unchanged. See the [title-row migration](../migrations/2026-11-22-all-expenses-header-row.json).

#### Proposed filter interface — not implemented

The user wants to replace the months argument with `include_filter` and `exclude_filter`; the function would use `R_MONTHS` internally. Proposed syntax: `ALL_EXPENSES_ROWS(ExpenseInclude, ExpenseExclude)`, with each argument a three-column range or array of Field, Operator, Value rules, excluding any control-table header row. Blank include means all records; blank exclude means no exclusions. All rules within a block must match; matching the nonempty exclusion block removes a record even when it matches inclusion. Numeric/date values remain typed, and invalid field names/operators should surface errors. The current months argument remains until the filter format is settled; no filter controls or filtering behavior were added by the title-row change.

### Future display filtering

Keep `=ALL_EXPENSES_ROWS(R_MONTHS)` complete as the shared report source. A future display function can accept the generated rows or the complete `S_ALL_EXPENSES` dataset and filter the structured fields before displaying them. Prefer reading the existing complete dataset for a separate view to avoid recalculating the raw sources twice. No filter controls, filtered view or additional sheet were added in this refactor.

Filtering records out of the source sheet would also remove them from future Budget/Cash Flow aggregations. Decide whether a filter is a viewing preference or a financial-scope rule before applying it there. No performance improvement is claimed from splitting formulas into named functions.

## Dependency boundaries

1. Raw expense calculations and All Expenses must remain independent of Budget and Cash Flow.
2. Budget reads Cost and Contingency from All Expenses and retains its subtotal, funding, carry-forward, and status calculations. Company Tax continues to read Cash Flow.
3. Cash Flow reads total expense amounts by month and type. Revenue and funding inflows retain their existing sources. Interest, principal, depreciation, and tax retain their specialized calculations.
4. The repayment script reads Cash Flow expense inputs excluding in-loan debt service, calculates repayments and tax in memory, and writes scheduled outputs. Keep that boundary: never feed total expenses including in-loan repayments back into its pre-debt-service input.
5. A formula-driven All Expenses row must not read Cash Flow tax or available cash. That would couple the upstream expense table to downstream results.

At the original migration, repayment and tax files contained unrelated local work. That work was subsequently committed separately in `ebee234`. Verify remote changes and local scope before any future Apps Script deployment.

## Implementation sequence and acceptance

Steps 1–4 are complete. Budget's consumer migration and explicit actual/planned eligibility are live and verified. Cash Flow remains on its original expense sources; its consumer migration requires a separate monthly aggregation interface and validation.

1. Inventory the live named functions/ranges and all consumers of All Expenses; retain formula snapshots and source-contract evidence.
2. Implement the expanded formula, source mappings, and refresh behavior locally where Apps Script is involved. Validate signed adjustments, empty datasets, errors, duplicate descriptions, salary mapping, financing aliases, and the actual/planned boundary offline.
3. Apply the All Expenses schema/formula in place. Reconcile all 768 monthly amounts and 172 Budget past/future values against the corrected baseline before switching report consumers. The only expected change to existing All Expenses base amounts is the documented historical Setup correction.
4. Migrate Budget detail expense retrieval while retaining named subtotal, contingency, and link functions. Verify annual required budget, available budget, carry-forward, headroom, and statuses.
5. Migrate Cash Flow expense retrieval. Verify every monthly expense, cash balance, operating profit, and dependent output. Do not run repayment recalculation against an unresolved script/source contract mismatch.
6. Measure representative source-edit recalculation before and after, using the same edits and restoring inputs. A timing benefit is not established by this reconciliation. Validate the native display and source links.
7. Synchronize authorized Apps Script changes using an isolated remote-matched clasp snapshot; verify remote content. Commit and push focused migration records and source changes.

Rollback must restore each changed formula and named-range definition from a fresh pre-write snapshot. Row and column expansion must preserve validation, formatting, named references, filters, and existing data. Take no dependency on temporary tabs or a manually refreshed cache.

## Expense function and range names — 22 November 2026

The generator is now `ALL_EXPENSES_ROWS(months)`, called from All Expenses A1 as `=ALL_EXPENSES_ROWS(R_MONTHS)`. `ALL_EXPENSES` is a named range covering All Expenses A1:M10273, including the title row and spare capacity. `S_ALL_EXPENSES` remains an alias of the same range for existing consumers. The rename changes no expense selection or calculations; all 1,064 records and 13 headers were verified unchanged.

## Sheet renamed to Budget Items — 23 November 2026

The All Expenses tab is now Budget Items, retaining sheet ID 1623286087 and its existing data, formulas, formatting and named ranges. `ALL_EXPENSES` and `S_ALL_EXPENSES` still identify the same dataset, and `ALL_EXPENSES_ROWS` remains the generator. The SOL refresh menu label and refresh sheet lookup/messages use the new sheet title; internal function names remain compatible. Historical references to All Expenses above describe this same sheet.

## Budget Items border automation removed — 23 November 2026

At the user's request, refreshing Budget Items no longer generates border requests, and Set Timeline Borders does nothing when Budget Items is active. Refresh still applies number formats and preserves the filter criteria. The Details and timeline border tools remain available for other sheets. Existing cell borders were not cleared; this change removes the code paths that modify them.

## Budget Items refresh removed — 23 November 2026

The SOL refresh command, `refreshAllExpenses()`, its formatting/filter request builder and the tests dedicated to that deleted implementation have been removed at the user's request. Budget Items continues to update through `ALL_EXPENSES_ROWS(R_MONTHS)`; the generator, named ranges, Budget formulas and existing sheet formatting are unchanged. Earlier refresh descriptions above are historical.

## Budget Items function and range renamed — 23 November 2026

`ALL_EXPENSES_ROWS(months)` is now `BUDGET_ITEMS_ROWS(months)`, and the primary named range `ALL_EXPENSES` is now `BUDGET_ITEMS`. Sheets updated the A1 caller and the references inside `BUDGET_EXPENSE_SUM` and the unused `BUDGET_ROWS` function. The existing `S_ALL_EXPENSES` compatibility alias is retained. All native named definitions were compared against the prior export: only the requested name substitutions changed. Every Budget Items, Budget and Cash Flow value was unchanged, with no formula errors. Local definitions and formula fixtures use the new names and were synchronized through clasp.

## Basis renamed to Status — 23 November 2026

The user renamed the live lookup to `BUDGET_ITEMS_SUM(status, amount_column)` and changed its header lookup to Status. The generator still emitted Basis, causing CHOOSEROWS errors. The generated fourth header is now Status, and the compact summary function `BUDGET_EXPENSE_TOTALS` uses that header too. Local source preserves the user's function/argument rename. Internal record parameters retain their existing names and behavior.

All Budget errors cleared and its values match the last healthy snapshot. Budget Items differs only in D1's header, and Cash Flow values are unchanged. Formula fixtures use the new header and lookup argument. The sheet anchor and all existing named-function calls remain in place.

## Budget timeline-link interface simplified — 23 November 2026

`BUDGET_TIMELINE_LINK(source_table, sheet_id, lookup_col, lookup_value, budget_year, period, amount)` now derives the start month and month count internally. `period` is `"Past"` (before the current month), `"Future"` (current month onward), or `"Full Year"`. Empty periods return zero before source-row lookup; unknown periods remain errors. The native definition is maintained in `src/BudgetTimelineLinkFormula.js`.

All 176 live Budget callers now pass `BUDGET_VALUE("_year")` and a period directly, retaining their existing named amount calculations and source destinations. Their outer LET blocks and duplicated month calculations were removed. `BUDGET_SOURCE_LINK` accepts the same year/period pair, and `BUDGET_ROWS` uses the simplified wrapper. The live report's existing `COMPANY_TAX_VALUES()` call and MAP-based output were preserved when updating local source.

Thirteen native boundary checks cover current, past and future years, January, December and invalid periods. Four additional native checks compare exact link destinations and values with the old interface, including an empty period with an error-valued amount. All Budget values, links and formatting were unchanged; Budget Items and Cash Flow were unchanged. See [formulas, definitions and rollback](../migrations/2026-11-23-budget-timeline-link.json).
