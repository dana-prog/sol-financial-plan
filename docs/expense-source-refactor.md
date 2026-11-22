# Expense source refactor

The migration will make All Expenses the shared expense dataset for Budget and Cash Flow. This document completes the mapping and reconciliation phase; it does not change the live All Expenses formula or either report's data source.

The baseline is the SOL financial plan on 22 November 2026, after the separately applied Road correction. The spreadsheet timezone is Asia/Bangkok. Cash Flow covers January 2026 through December 2033; Budget currently covers 2026 through 2029. Do not extend Budget's horizon as part of this migration.

## Reconciliation results

Evidence: [captured baseline and formula mapping](../migrations/2026-11-22-expense-source-refactor-baseline.json).

- Independently reconstructed 1,063 nonzero expense records from source values, using actual Setup before November 2026 and planned Setup thereafter.
- All eight Cash Flow expense types reconcile over all 96 months: 768 comparisons, no difference greater than THB 0.01.
- All 86 Budget detail lines reconcile for both past and future amounts: 172 comparisons, no difference greater than THB 0.01. Company Tax is compared to its existing Cash Flow source and remains outside the proposed expense dataset.
- No nonzero source expense within Budget's existing years is missing a matching Budget line after the Road correction. Out Loan Disbursements and out loans disbursements require an explicit alias; do not rely on singular/plural display labels as keys.
- No negative expense records were present in this baseline. Signed adjustments still require synthetic validation before rollout.
- These are offline comparisons of captured effective values, not live replacement-formula tests or performance measurements. Budget contingency, annual totals, and headroom still require end-to-end migration verification.

## Corrections and source exceptions

### Road correction completed

Budget now includes Road in its 2027, 2028, and 2029 infrastructure sections. The 2028 amount is THB 2,000,000 plus THB 300,000 contingency. The 2028 infrastructure required budget is THB 16,692,250. Annual required budget is THB 231,801,665.99853653 and headroom is THB 34,089,641.219913065. The lower headroom carries into 2029, reducing its headroom to THB 17,139,304.4375.

The migration record is [2026-11-22-budget-road-expense.json](../migrations/2026-11-22-budget-road-expense.json), committed separately as `80ea91b`. Existing named formulas were preserved. New rows use the existing infrastructure formulas and their source links resolve to the Road row.

### Setup basis correction required in All Expenses

Current All Expenses shows planned Setup amounts in historical months. Nine months from January through September 2026 differ from the actual-source amounts already used by Budget and Cash Flow; their aggregate difference is THB 591,150. Moving All Expenses to the agreed cutoff reduces its Setup base total by that amount while preserving the current report amounts.

The current month belongs to the planned period. A missing or zero actual amount must remain missing or zero according to the existing source semantics; do not fall back to planned values. A record's basis must come from its source, not merely from its date.

### Numeric-looking text remains a separate source issue

`Infra Exp (P)!AE9` contains the text `฿ 96K` for February 2028. Existing source totals ignore it, and it is explicitly excluded from the reconciliation. Do not silently parse it as THB 96,000. Record it as a validation exception and resolve whether it is intended expenditure before adopting a policy that rejects all text. Changing it would be a separate financial correction.

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

Use one record per selected source row and month. Identical display descriptions may legitimately appear more than once. Source row positions can change; regenerate provenance from the source rather than treating row numbers as permanent IDs. Keep current type/date ordering. Do not use the visible filter-sensitive grand total as an input to either report.

Base and contingency stay separate. Budget's existing category contingency formulas may continue to apply once to base expenses; Cash Flow consumes Total cost. Do not feed Total cost to Budget's base-expense columns.

## Formula and refresh contract

Keep All Expenses formula-driven for this migration. Extend its existing formula and named-function approach. Inventory live named ranges and definitions before implementing new names; the cell snapshot records calls, not the complete live named-function definitions.

Candidate interfaces, subject to that inventory:

- `S_ALL_EXPENSES`: header plus data, excluding the display total.
- `ALL_EXPENSES_MONTHLY_VALUES(expense_type, amount_column)`: returns one value for every month in `R_MONTHS`, including zero months.
- `ALL_EXPENSES_BUDGET_SUM(budget_category, budget_line, start_month, month_count, amount_column)`: sums the base-cost field for the same date windows as the existing Budget formulas; zero months returns zero.

Reuse existing header lookup functions such as `COL_VALUES` and `DATA_ROWS`. Keep `BUDGET_VALUE`, `BUDGET_CATEGORY_SUM`, `BUDGET_YEAR_SUM`, `BUDGET_YEAR_CONTINGENCY`, `BUDGET_TIMELINE_LINK`, and `CASH_FLOW_PARAM_VALUES` in their existing roles. Retain source drill-down links. Replace source-specific retrieval calls only with a reviewed named-function interface, never direct ranges or inlined implementations. Do not globally redirect source functions such as `CONSTRUCTION_PARAM_VALUES` or `OPERATIONS_PARAM_VALUES` to All Expenses: its own calculations depend on those functions.

The current `refreshAllExpenses()` writes static values and can overwrite the formula anchor. It must be adapted before report consumers are switched. It also hardcodes legacy financing labels (`in loans repaid`, `out loans disbursed`) that differ from the live source parameters. Do not run it as a migration shortcut. Preserve its useful formatting behavior without maintaining a second independent financial calculation. Review related border/filter assumptions when expanding beyond six columns.

## Dependency boundaries

1. Raw expense calculations and All Expenses must remain independent of Budget and Cash Flow.
2. Budget reads base expense amounts and retains its current contingency, funding, carry-forward, and status calculations. Company Tax continues to read Cash Flow.
3. Cash Flow reads total expense amounts by month and type. Revenue and funding inflows retain their existing sources. Interest, principal, depreciation, and tax retain their specialized calculations.
4. The repayment script reads Cash Flow expense inputs excluding in-loan debt service, calculates repayments and tax in memory, and writes scheduled outputs. Keep that boundary: never feed total expenses including in-loan repayments back into its pre-debt-service input.
5. A formula-driven All Expenses row must not read Cash Flow tax or available cash. That would couple the upstream expense table to downstream results.

The local repayment and tax files contain pre-existing uncommitted work. The live labels and local script expectations also differ in places. Preserve that work; verify the currently deployed source before any related Apps Script deployment. Do not publish the entire dirty working tree.

## Implementation sequence and acceptance

1. Inventory the live named functions/ranges and all consumers of All Expenses; retain formula snapshots and source-contract evidence.
2. Implement the expanded formula, source mappings, and refresh behavior locally where Apps Script is involved. Validate signed adjustments, empty datasets, errors, duplicate descriptions, salary mapping, financing aliases, and the actual/planned boundary offline.
3. Apply the All Expenses schema/formula in place. Reconcile all 768 monthly amounts and 172 Budget past/future values against the corrected baseline before switching report consumers. The only expected change to existing All Expenses base amounts is the documented historical Setup correction.
4. Migrate Budget detail expense retrieval while retaining named subtotal, contingency, and link functions. Verify annual required budget, available budget, carry-forward, headroom, and statuses.
5. Migrate Cash Flow expense retrieval. Verify every monthly expense, cash balance, operating profit, and dependent output. Do not run repayment recalculation against an unresolved script/source contract mismatch.
6. Measure representative source-edit recalculation before and after, using the same edits and restoring inputs. A timing benefit is not established by this reconciliation. Validate the native display and source links.
7. Synchronize authorized Apps Script changes using an isolated remote-matched clasp snapshot; verify remote content. Commit and push focused migration records and source changes.

Rollback must restore each changed formula and named-range definition from a fresh pre-write snapshot. Row and column expansion must preserve validation, formatting, named references, filters, and existing data. Take no dependency on temporary tabs or a manually refreshed cache.
