# Expense source refactor

The migration will make All Expenses the shared expense dataset for Budget and Cash Flow. The All Expenses dataset upgrade is live and verified. Budget and Cash Flow still read their original sources; their consumer migration is the next stage.

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

Retain the existing first six columns to minimize disruption to the visible report. Append structured fields to the same sheet; no staging or temporary sheet is required. Keep source identifiers separate from display descriptions.

| Column | Field | Meaning |
| --- | --- | --- |
| A | Type | Existing expense-type label |
| B | Year | Derived from the real Month date |
| C | Quarter | Derived display value |
| D | Month | First-of-month date, displayed as `mmmm` |
| E | Param | Existing readable description; not an aggregation key |
| F | Cost | Base cost, excluding contingency |
| G | Budget category | Stable category corresponding to Budget's `_category` |
| H | Budget line | Explicit matching key, with financing aliases normalized |
| I | Source parameter | Unmodified source parameter or obligation name |
| J | Source sheet | Origin for provenance and validation |
| K | Basis | Actual, Planned, or Scheduled; financing and obligations use Scheduled unless their sources establish actual status |
| L | Contingency | Signed base cost multiplied by the applicable rate; zero for actual Setup and financing/land |
| M | Total cost | Cost plus Contingency |
| N | Source row | Source location for traceability, not a persistent identity |

Use one record per selected source row and month. Identical display descriptions may legitimately appear more than once. Source row positions can change; regenerate provenance from the source rather than treating row numbers as permanent IDs. Keep current type/date ordering. The dataset has no embedded grand-total record. Report consumers must aggregate the appropriate amount column.

Base and contingency stay separate. Budget's existing category contingency formulas may continue to apply once to base expenses; Cash Flow consumes Total cost. Do not feed Total cost to Budget's base-expense columns.

## Formula and refresh contract

All Expenses is formula-driven at A2. Its formula is stored in `src/AllExpensesFormula.js`. The four existing detail functions (`SETUP_EXPENSES_D`, `CONSTRUCTION_EXPENSES_D`, `OPERATIONAL_EXPENSES_D`, `LAND_EXPENSES_D`) were extended to six outputs: type, parameter, party, base amount, source row and source category. The inventory found All Expenses A2 to be their only consumer. Their definitions and original versions are retained in the migration record. Existing named-function calls are preserved.

Current and proposed interfaces:

- `S_ALL_EXPENSES` is live: All Expenses A1:N10273, with a header and no total record. Its capacity covers every currently eligible source row in all 96 months. Revisit capacity when source rows or the timeline expand.
- `ALL_EXPENSES_MONTHLY_VALUES(expense_type, amount_column)`: returns one value for every month in `R_MONTHS`, including zero months.
- `ALL_EXPENSES_BUDGET_SUM(budget_category, budget_line, start_month, month_count, amount_column)`: sums the base-cost field for the same date windows as the existing Budget formulas; zero months returns zero.

Reuse existing header lookup functions such as `COL_VALUES` and `DATA_ROWS`. Keep `BUDGET_VALUE`, `BUDGET_CATEGORY_SUM`, `BUDGET_YEAR_SUM`, `BUDGET_YEAR_CONTINGENCY`, `BUDGET_TIMELINE_LINK`, and `CASH_FLOW_PARAM_VALUES` in their existing roles. Retain source drill-down links. Replace source-specific retrieval calls only with a reviewed named-function interface, never direct ranges or inlined implementations. Do not globally redirect source functions such as `CONSTRUCTION_PARAM_VALUES` or `OPERATIONS_PARAM_VALUES` to All Expenses: its own calculations depend on those functions.

`refreshAllExpenses()` now validates the live anchor and headers, then updates formatting, the filter range and borders without writing expense values. The duplicate static financial calculation was removed. Filter criteria are retained, but saved sort specifications are not reapplied: a live test demonstrated that doing so moves the array anchor. Type/date ordering belongs to the formula. The corrected refresh was executed and verified to preserve A2 and all 1,064 records.

## Dependency boundaries

1. Raw expense calculations and All Expenses must remain independent of Budget and Cash Flow.
2. Budget reads base expense amounts and retains its current contingency, funding, carry-forward, and status calculations. Company Tax continues to read Cash Flow.
3. Cash Flow reads total expense amounts by month and type. Revenue and funding inflows retain their existing sources. Interest, principal, depreciation, and tax retain their specialized calculations.
4. The repayment script reads Cash Flow expense inputs excluding in-loan debt service, calculates repayments and tax in memory, and writes scheduled outputs. Keep that boundary: never feed total expenses including in-loan repayments back into its pre-debt-service input.
5. A formula-driven All Expenses row must not read Cash Flow tax or available cash. That would couple the upstream expense table to downstream results.

The local repayment and tax files contain pre-existing uncommitted work. The live labels and local script expectations also differ in places. Preserve that work; verify the currently deployed source before any related Apps Script deployment. Do not publish the entire dirty working tree.

## Implementation sequence and acceptance

Steps 1–3 and the Apps Script synchronization for this stage are complete. The next step is to define and validate the named aggregation interfaces, then migrate Budget before Cash Flow. No report-consumer replacement has been applied.

1. Inventory the live named functions/ranges and all consumers of All Expenses; retain formula snapshots and source-contract evidence.
2. Implement the expanded formula, source mappings, and refresh behavior locally where Apps Script is involved. Validate signed adjustments, empty datasets, errors, duplicate descriptions, salary mapping, financing aliases, and the actual/planned boundary offline.
3. Apply the All Expenses schema/formula in place. Reconcile all 768 monthly amounts and 172 Budget past/future values against the corrected baseline before switching report consumers. The only expected change to existing All Expenses base amounts is the documented historical Setup correction.
4. Migrate Budget detail expense retrieval while retaining named subtotal, contingency, and link functions. Verify annual required budget, available budget, carry-forward, headroom, and statuses.
5. Migrate Cash Flow expense retrieval. Verify every monthly expense, cash balance, operating profit, and dependent output. Do not run repayment recalculation against an unresolved script/source contract mismatch.
6. Measure representative source-edit recalculation before and after, using the same edits and restoring inputs. A timing benefit is not established by this reconciliation. Validate the native display and source links.
7. Synchronize authorized Apps Script changes using an isolated remote-matched clasp snapshot; verify remote content. Commit and push focused migration records and source changes.

Rollback must restore each changed formula and named-range definition from a fresh pre-write snapshot. Row and column expansion must preserve validation, formatting, named references, filters, and existing data. Take no dependency on temporary tabs or a manually refreshed cache.
