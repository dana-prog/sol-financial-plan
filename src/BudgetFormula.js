// Native Sheets definition: Budget reads the materialized expense dataset, not its generator.
const BUDGET_EXPENSE_SUM_DEFINITION = {
  name: 'BUDGET_EXPENSE_SUM',
  arguments: ['basis', 'amount_column'],
  description: 'Sums All Expenses for the current Budget year, category and budget line. A blank budget line selects the category; blank basis selects Actual and Planned. Use Cost or Contingency. Missing matches return zero; source errors remain visible.',
  formula: String.raw`=LET(
  category, BUDGET_VALUE("_category"),
  budget_year, BUDGET_VALUE("_year"),
  budget_line, LOWER(TRIM(BUDGET_VALUE("budget line"))),
  line_key, IF(budget_line="out loans disbursements", "out loan disbursements", budget_line),
  categories, COL_VALUES(ALL_EXPENSES, "Budget Category"),
  lines, COL_VALUES(ALL_EXPENSES, "Budget Line"),
  years, COL_VALUES(ALL_EXPENSES, "Year"),
  bases, COL_VALUES(ALL_EXPENSES, "Basis"),
  amounts, COL_VALUES(ALL_EXPENSES, amount_column),
  matches, ARRAYFORMULA((categories=category)*(years=budget_year)*
    IF(line_key="", 1, LOWER(TRIM(lines))=line_key)*
    IF(basis="", 1, bases=basis)),
  IF(SUM(matches)=0, 0, SUM(FILTER(amounts, matches)))
)`,
};
