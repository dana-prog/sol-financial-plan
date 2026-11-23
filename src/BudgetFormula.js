// Native Sheets definition: Budget reads the materialized expense dataset, not its generator.
// Retain the first data row even when blank so COL_VALUES can handle an empty dataset.
const BUDGET_EXPENSE_SUM_DEFINITION = {
  name: 'BUDGET_EXPENSE_SUM',
  arguments: ['basis', 'amount_column'],
  description: 'Sums All Expenses for the current Budget year, category and budget line. A blank budget line selects the category; blank basis selects Actual and Planned. Use Cost or Contingency. Missing matches return zero; source errors remain visible.',
  formula: String.raw`=LET(
  category, BUDGET_VALUE("_category"),
  budget_year, BUDGET_VALUE("_year"),
  budget_line, LOWER(TRIM(BUDGET_VALUE("budget line"))),
  line_key, IF(budget_line="out loans disbursements", "out loan disbursements", budget_line),
  headers, HEADER_ROW(ALL_EXPENSES),
  positions, SEQUENCE(ROWS(ALL_EXPENSES)),
  category_column, CHOOSECOLS(ALL_EXPENSES, MATCH("Budget Category", headers, 0)),
  populated_rows, FILTER(positions, ARRAYFORMULA((category_column<>"")+(positions=2))),
  expense_column, LAMBDA(header,
    COL_VALUES(CHOOSEROWS(CHOOSECOLS(ALL_EXPENSES, MATCH(header, headers, 0)), populated_rows), header)),
  categories, expense_column("Budget Category"),
  lines, expense_column("Budget Line"),
  years, expense_column("Year"),
  bases, expense_column("Basis"),
  amounts, expense_column(amount_column),
  matches, ARRAYFORMULA((categories=category)*(years=budget_year)*
    IF(line_key="", 1, LOWER(TRIM(lines))=line_key)*
    IF(basis="", 1, bases=basis)),
  IF(SUM(matches)=0, 0, SUM(FILTER(amounts, matches)))
)`,
};
