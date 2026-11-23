// The report catalog is explicit so a zero-cost line does not disappear.
const BUDGET_REPORT_YEARS = [2026, 2027, 2028, 2029];
const BUDGET_REPORT_CATEGORIES = [
  ['Out Loans', ['out loans disbursements']],
  ['In Loans', ['In Loan Repayment']],
  ['Land', ['Land']],
  ['Setup', ['Accounting', 'Legal', 'Permits', 'Planning']],
  ['Infra', ['Electricity', 'Water', 'Wifi', 'Landscaping', 'Construction Management', 'Road', 'Cars']],
  ['Construction', ['Construction']],
  ['Operations', ['Company Tax', 'Insurance', 'Wifi', 'Software', 'Legal', 'Electricity', 'Water', 'Food', 'Maintenance', 'Wear & tear', 'Marketing', 'Salaries']],
];
const BUDGET_REPORT_HEADERS = ['_category', '_year', 'Year', 'Category', 'Budget line',
  'Past Expenses', 'Future Expenses', 'Total Expenses', 'Contingency', 'Required Budget',
  'Available Budget', 'Headroom', 'Projected Status'];

const BUDGET_REPORT_LAYOUT = BUDGET_REPORT_YEARS.flatMap(year => [
  [year, '', '', 'Year', false],
  ...BUDGET_REPORT_CATEGORIES
    .filter(([category]) => year !== 2026 || ['Out Loans', 'Setup'].includes(category))
    .flatMap(([category, lines]) => [
      [year, category, '', 'Category', false],
      ...lines.map(line => [year, category,
        year === 2026 && category === 'Out Loans' ? 'Out Loan Disbursements' : line,
        'Line', year === 2029 && line !== 'Company Tax']),
    ]),
]);

const BUDGET_ROWS_FORMULA = '=BUDGET_ROWS()';
const BUDGET_ROWS_NAMED_FUNCTIONS = [
  {
    name: 'BUDGET_ROW_LAYOUT',
    arguments: [],
    description: 'Budget report catalog: year, category, line, row kind, and full-year total link. Includes zero-cost lines and the existing 2026–2029 layout.',
    formula: '={' + BUDGET_REPORT_LAYOUT.map(row => row.map(value =>
      typeof value === 'string' ? '"' + value.replaceAll('"', '""') + '"' : String(value).toUpperCase()
    ).join(',')).join(';\n') + '}',
  },
  {
    name: 'BUDGET_EXPENSE_TOTALS',
    arguments: ['expenses'],
    description: 'Groups an All Expenses table once by year, category and line, returning Past Expenses, Future Expenses and Contingency. Preserves credits and source errors.',
    formula: String.raw`=LET(
  positions, SEQUENCE(ROWS(expenses)),
  category_column, CHOOSECOLS(expenses, MATCH("Budget Category", HEADER_ROW(expenses), 0)),
  source, FILTER(expenses, ARRAYFORMULA((category_column<>"")+(positions=2))),
  years, COL_VALUES(source, "Year"),
  categories, COL_VALUES(source, "Budget Category"),
  lines, COL_VALUES(source, "Budget Line"),
  bases, COL_VALUES(source, "Basis"),
  costs, COL_VALUES(source, "Cost"),
  contingency, COL_VALUES(source, "Contingency"),
  fields, TOCOL(HSTACK(years, categories, lines, bases, costs, contingency)),
  errors, ARRAYFORMULA(ISERROR(fields)),
  IF(SUM(ARRAYFORMULA(N(errors)))>0, INDEX(fields, MATCH(TRUE, errors, 0)),
    IF(SUM(ARRAYFORMULA((categories<>"")*(NOT(ISNUMBER(costs))+NOT(ISNUMBER(contingency)))))>0, NA(),
      QUERY(HSTACK(ARRAYFORMULA(IF(categories="", 0, years)), categories,
        ARRAYFORMULA(LOWER(TRIM(lines))),
        ARRAYFORMULA(IF(bases="Actual", costs, 0)),
        ARRAYFORMULA(IF(bases="Planned", costs, 0)),
        ARRAYFORMULA(IF(categories="", 0, contingency))),
        "select Col1,Col2,Col3,sum(Col4),sum(Col5),sum(Col6) group by Col1,Col2,Col3 label Col1 'Year',Col2 'Category',Col3 'Line',sum(Col4) 'Past Expenses',sum(Col5) 'Future Expenses',sum(Col6) 'Contingency'", 0)
    )
  )
)`,
  },
  {
    name: 'BUDGET_REPORT_AMOUNT',
    arguments: ['totals', 'budget_year', 'category', 'budget_line', 'amount_column'],
    description: 'Reads an amount from a compact Budget totals table. Blank category selects the year; blank line selects the category. Normalizes the existing Out Loans label.',
    formula: String.raw`=LET(
  line, LOWER(TRIM(budget_line)),
  line_key, IF(line="out loans disbursements", "out loan disbursements", line),
  matches, ARRAYFORMULA((COL_VALUES(totals, "Year")=budget_year)*
    IF(category="", 1, COL_VALUES(totals, "Category")=category)*
    IF(line_key="", 1, LOWER(TRIM(COL_VALUES(totals, "Line")))=line_key)),
  IF(SUM(matches)=0, 0, SUM(FILTER(COL_VALUES(totals, amount_column), matches)))
)`,
  },
  {
    name: 'BUDGET_SOURCE_LINK',
    arguments: ['category', 'budget_line', 'start_month', 'month_count', 'amount'],
    description: 'Preserves Budget drill-down destinations through BUDGET_TIMELINE_LINK. Source tables locate links only; amount is already calculated.',
    formula: String.raw`=IF(month_count=0, 0,
  SWITCH(category,
    "Out Loans", BUDGET_TIMELINE_LINK(S_OUT_LOANS, 198423991, "param", "total out loans disbursements", start_month, month_count, amount),
    "In Loans", BUDGET_TIMELINE_LINK(S_IN_LOANS, 1346380865, "param", "total in loans repayments", start_month, month_count, amount),
    "Land", BUDGET_TIMELINE_LINK(S_OBLIGATIONS, 1675897452, "_category", "Land", start_month, month_count, amount),
    "Setup", BUDGET_TIMELINE_LINK(S_SETUP_EXP_P, 1304432573, "_category", budget_line, start_month, month_count, amount),
    "Infra", BUDGET_TIMELINE_LINK(S_INFRA_EXP_P, 1754072773, "_category", budget_line, start_month, month_count, amount),
    "Construction", BUDGET_TIMELINE_LINK(S_CONSTRUCTION_EXP_P, 462995266, "param", "total costs", start_month, month_count, amount),
    "Operations", BUDGET_TIMELINE_LINK(S_OPERATIONAL_EXP, 1093991711, "param", LOWER(budget_line)&" expenses", start_month, month_count, amount),
    NA()
  )
)`,
  },
  {
    name: 'BUDGET_ROWS',
    arguments: [],
    description: 'Generates Budget headers, detail lines, subtotals, funding, carry-forward and statuses. Aggregates the BUDGET_ITEMS named range once; Company Tax remains sourced from Cash Flow.',
    formula: String.raw`=LET(
  headers, {${BUDGET_REPORT_HEADERS.map(header => '"' + header + '"').join(',')}},
  totals_headers, {"Year","Category","Line","Past Expenses","Future Expenses","Contingency"},
  layout, BUDGET_ROW_LAYOUT(),
  years, CHOOSECOLS(layout, 1),
  categories, CHOOSECOLS(layout, 2),
  lines, CHOOSECOLS(layout, 3),
  kinds, CHOOSECOLS(layout, 4),
  report_years, FILTER(years, kinds="Year"),
  cutoff, MONTH_START(TODAY()),
  month_years, ARRAYFORMULA(YEAR(R_MONTHS)),
  expense_totals, BUDGET_EXPENSE_TOTALS(BUDGET_ITEMS),
  tax, CASH_FLOW_PARAM_VALUES("company tax"),
  tax_for_period, LAMBDA(budget_year, basis,
    LET(matches, ARRAYFORMULA((month_years=budget_year)*IF(basis="Actual", R_MONTHS<cutoff, R_MONTHS>=cutoff)),
      IF(SUM(matches)=0, 0, SUM(FILTER(tax, matches))))),
  tax_past, MAP(report_years, LAMBDA(budget_year, tax_for_period(budget_year, "Actual"))),
  tax_future, MAP(report_years, LAMBDA(budget_year, tax_for_period(budget_year, "Planned"))),
  line_past, MAP(years, categories, lines, LAMBDA(budget_year, category, line,
    IF(line="", 0, IF(line="Company Tax", XLOOKUP(budget_year, report_years, tax_past),
      BUDGET_REPORT_AMOUNT(expense_totals, budget_year, category, line, "Past Expenses"))))),
  line_future, MAP(years, categories, lines, LAMBDA(budget_year, category, line,
    IF(line="", 0, IF(line="Company Tax", XLOOKUP(budget_year, report_years, tax_future),
      BUDGET_REPORT_AMOUNT(expense_totals, budget_year, category, line, "Future Expenses"))))),
  line_contingency, MAP(years, categories, lines, LAMBDA(budget_year, category, line,
    IF(OR(line="", line="Company Tax"), 0,
      BUDGET_REPORT_AMOUNT(expense_totals, budget_year, category, line, "Contingency")))),
  detail_totals, VSTACK(totals_headers, HSTACK(years, categories, lines, line_past, line_future, line_contingency)),
  past, MAP(years, categories, lines, line_past, LAMBDA(budget_year, category, line, amount,
    IF(line="", BUDGET_REPORT_AMOUNT(detail_totals, budget_year, category, "", "Past Expenses"), amount))),
  future, MAP(years, categories, lines, line_future, LAMBDA(budget_year, category, line, amount,
    IF(line="", BUDGET_REPORT_AMOUNT(detail_totals, budget_year, category, "", "Future Expenses"), amount))),
  total, ARRAYFORMULA(past+future),
  contingency, MAP(years, categories, lines, LAMBDA(budget_year, category, line,
    IF(line="", BUDGET_REPORT_AMOUNT(detail_totals, budget_year, category, "", "Contingency"), ""))),
  required, ARRAYFORMULA(IF(lines="", total+N(contingency), "")),
  annual_required, FILTER(required, kinds="Year"),
  funding, ARRAYFORMULA(IN_LOAN_PARAM_VALUES_BY_LOAN("totals", "total in loans received")+
    OUT_LOAN_PARAM_VALUES_BY_LOAN("totals", "total out loan repayments received")+REVENUE_PARAM_VALUES("total revenue")),
  annual_funding, MAP(report_years, LAMBDA(budget_year, SUM(FILTER(funding, month_years=budget_year)))),
  annual_headroom, SCAN(0, ARRAYFORMULA(annual_funding-annual_required), LAMBDA(balance, change, balance+change)),
  available, MAP(years, kinds, LAMBDA(budget_year, kind,
    IF(kind="Year", XLOOKUP(budget_year, report_years, ARRAYFORMULA(annual_headroom+annual_required)), ""))),
  headroom, MAP(years, kinds, LAMBDA(budget_year, kind,
    IF(kind="Year", XLOOKUP(budget_year, report_years, annual_headroom), ""))),
  statuses, MAP(kinds, past, required, available, LAMBDA(kind, paid, needed, funds,
    IF(kind<>"Year", "", IFS(paid>funds, "Over Budget", needed>funds, "Projected Over Budget", funds=0, "No Budget", TRUE, "Within Budget")))),
  report, VSTACK(headers, HSTACK(categories, years, ARRAYFORMULA(IF(kinds="Year", years, "")),
    ARRAYFORMULA(IF(kinds="Category", categories, "")), lines,
    past, future, total, contingency, required, available, headroom, statuses)),
  MAKEARRAY(ROWS(report), COLUMNS(report), LAMBDA(report_row, report_column,
    LET(amount, INDEX(report, report_row, report_column),
      IF(OR(report_row=1, report_column<6, report_column>8), amount,
        LET(detail_row, report_row-1,
          budget_year, INDEX(years, detail_row),
          category, INDEX(categories, detail_row),
          line, INDEX(lines, detail_row),
          IF(OR(line="", line="Company Tax"), amount,
            SWITCH(report_column,
              6, IF(category="Setup", amount,
                BUDGET_SOURCE_LINK(category, line, DATE(budget_year, 1, 1), MAX(0, MIN(12, (YEAR(cutoff)-budget_year)*12+MONTH(cutoff)-1)), amount)),
              7, BUDGET_SOURCE_LINK(category, line, MAX(DATE(budget_year, 1, 1), cutoff), MAX(0, MIN(12, (budget_year-YEAR(cutoff))*12+13-MONTH(cutoff))), amount),
              8, IF(INDEX(layout, detail_row, 5), BUDGET_SOURCE_LINK(category, line, DATE(budget_year, 1, 1), 12, amount), amount)
            )
          )
        )
      )
    )
  ))
)`,
  },
];
