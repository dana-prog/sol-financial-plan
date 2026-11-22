const ALL_EXPENSES_HEADERS = [
  'Type', 'Year', 'Quarter', 'Month', 'Param', 'Cost', 'Budget category',
  'Budget line', 'Source parameter', 'Source sheet', 'Basis', 'Contingency', 'Total cost', 'Source row',
];

// A2 is the public entry point; definitions below are maintained in Sheets as named functions.
const ALL_EXPENSES_FORMULA = '=ALL_EXPENSES(R_MONTHS)';

// Dependency order supports importing or reviewing the definitions from top to bottom.
const ALL_EXPENSES_NAMED_FUNCTIONS = [
  {
    name: 'EXPENSE_DETAIL_ROWS',
    arguments: ['details', 'eligible'],
    description: 'Returns eligible nonzero signed expense details. Rejects nonempty text amounts and preserves source errors; empty results return one blank six-column row.',
    formula: String.raw`=LET(
  empty_details, {"","","","","",""},
  IF(SUM(ARRAYFORMULA(N(eligible)))=0, empty_details,
    LET(
      selected, FILTER(details, eligible),
      costs, CHOOSECOLS(selected, 4),
      errors, ARRAYFORMULA(ISERROR(costs)),
      IF(SUM(ARRAYFORMULA(N(errors)))>0, INDEX(costs, MATCH(TRUE, errors, 0)),
        IF(SUM(ARRAYFORMULA(N((costs<>"")*NOT(ISNUMBER(costs)))))>0, NA(),
          LET(
            nonzero, ARRAYFORMULA(ISNUMBER(costs)*(costs<>0)),
            IF(SUM(nonzero)=0, empty_details, FILTER(selected, nonzero))
          )
        )
      )
    )
  )
)`,
  },
  {
    name: 'INFRA_EXPENSES_D',
    arguments: ['month'],
    description: 'Returns monthly infrastructure details: type, parameter, supplier, base cost, source row and category.',
    formula: String.raw`=LET(
  infra_params, COL_VALUES(S_INFRA_EXP_P, "param"),
  infra_categories, COL_VALUES(S_INFRA_EXP_P, "_category"),
  infra_suppliers, COL_VALUES(S_INFRA_EXP_P, "supplier"),
  infra_data, DATA_ROWS(S_INFRA_EXP_P),
  infra_headers, HEADER_ROW(S_INFRA_EXP_P),
  infra_rows, SEQUENCE(ROWS(infra_params), 1, ROW(INDEX(S_INFRA_EXP_P, 1, 1)) + 1),
  EXPENSE_DETAIL_ROWS(
    HSTACK(ARRAYFORMULA(IF(infra_params<>"", "Infrastructure Expenses", "")), infra_params,
      infra_suppliers, CHOOSECOLS(infra_data, MATCH(month, infra_headers, 0)), infra_rows, infra_categories),
    ARRAYFORMULA((infra_params<>"")*(LOWER(TRIM(infra_categories))<>"totals"))
  )
)`,
  },
  {
    name: 'FINANCING_EXPENSES_D',
    arguments: ['month', 'source', 'param_name', 'party_header', 'type_label'],
    description: 'Returns monthly scheduled loan details for the specified source and parameter, excluding total rows.',
    formula: String.raw`=LET(
  params, COL_VALUES(source, "param"),
  names, COL_VALUES(source, "_name"),
  parties, ARRAYFORMULA(IFNA(XLOOKUP(names, COL_VALUES(source, "name"), COL_VALUES(source, party_header)), "")),
  values, CHOOSECOLS(DATA_ROWS(source), MATCH(month, HEADER_ROW(source), 0)),
  source_rows, SEQUENCE(ROWS(params), 1, ROW(INDEX(source, 1, 1)) + 1),
  EXPENSE_DETAIL_ROWS(
    HSTACK(ARRAYFORMULA(IF(names<>"", type_label, "")), names, parties, values, source_rows, params),
    ARRAYFORMULA((params=param_name)*(names<>"")*(LOWER(TRIM(names))<>"totals"))
  )
)`,
  },
  {
    name: 'EXPENSE_RECORDS',
    arguments: ['expenses', 'month', 'source_name', 'category', 'basis', 'rate'],
    description: 'Adds month, Budget mapping, source, basis, contingency and total to six-column source details; returns twelve internal record columns.',
    formula: String.raw`=LET(
  types, CHOOSECOLS(expenses, 1),
  params, CHOOSECOLS(expenses, 2),
  costs, CHOOSECOLS(expenses, 4),
  repeat, LAMBDA(value, ARRAYFORMULA(IF(types<>"", value, ""))),
  lines, MAP(params, CHOOSECOLS(expenses, 6), LAMBDA(param, source_category,
      IF(param="", "", SWITCH(category,
          "Setup", source_category, "Infra", source_category, "Construction", "Construction",
          "Land", "Land", "In Loans", "In Loan Repayment", "Out Loans", "Out Loan Disbursements",
          "Operations", SWITCH(param, "total guest services salaries", "Salaries",
            "total operations support salaries", "Salaries", PROPER(REGEXREPLACE(param, " expenses$", ""))),
          NA()
      ))
  )),
  HSTACK(types, repeat(month), CHOOSECOLS(expenses, 2, 3, 4), repeat(category), lines,
    repeat(source_name), repeat(basis), ARRAYFORMULA(IF(types<>"", costs*rate, "")),
    ARRAYFORMULA(IF(types<>"", costs*(1+rate), "")), CHOOSECOLS(expenses, 5))
)`,
  },
  {
    name: 'EXPENSES_FOR_MONTH',
    arguments: ['month'],
    description: 'Combines financing, Setup, Land, Infrastructure, Construction and Operations for one month, with their existing basis and contingency rules.',
    formula: String.raw`=LET(
  actual_setup, month < MONTH_START(TODAY()),
  VSTACK(
    EXPENSE_RECORDS(FINANCING_EXPENSES_D(month, S_IN_LOANS, "in loans repayments", "lender", "In Loan Repayments"),
      month, "In Loans", "In Loans", "Scheduled", 0),
    EXPENSE_RECORDS(FINANCING_EXPENSES_D(month, S_OUT_LOANS, "out loans disbursements", "debtor", "Out Loan Disbursements"),
      month, "Out Loans", "Out Loans", "Scheduled", 0),
    EXPENSE_RECORDS(SETUP_EXPENSES_D(month), month, IF(actual_setup, "Setup Exp (A)", "Setup Exp (P)"),
      "Setup", IF(actual_setup, "Actual", "Planned"), IF(actual_setup, 0, PARAM("setup contingency"))),
    EXPENSE_RECORDS(LAND_EXPENSES_D(month), month, "Obligations (P)", "Land", "Scheduled", 0),
    EXPENSE_RECORDS(INFRA_EXPENSES_D(month), month, "Infra Exp (P)", "Infra", "Planned", PARAM("infra contingency")),
    EXPENSE_RECORDS(CONSTRUCTION_EXPENSES_D(month), month, "Construction Exp (P)", "Construction", "Planned", PARAM("construction contingency")),
    EXPENSE_RECORDS(OPERATIONAL_EXPENSES_D(month), month, "Operational Exp (P)", "Operations", "Planned", PARAM("operating contingency"))
  )
)`,
  },
  {
    name: 'EXPENSE_DESCRIPTION',
    arguments: ['type', 'month', 'param', 'party'],
    description: 'Returns the readable expense description, including construction quantities and singular/plural unit names.',
    formula: String.raw`=IF(type="Construction Expenses",
  LET(
    units, INDEX(CONSTRUCTION_PARAM_VALUES(REGEXREPLACE(param, " const[.] cost$", " const. count")), 1, MATCH(month, R_MONTHS, 0)),
    unit_name, REGEXREPLACE(param, " const[.] cost$", ""),
    singular, SWITCH(unit_name, "clinics", "clinic", "sala toilets", "sala toilet", "complex room", "complex home", unit_name),
    plural, SWITCH(unit_name, "lobby", "lobbies", "storage", "storage", "pool lounge seating", "pool lounge seating",
      "general resort storage", "general resort storage", "shower & toilet", "showers & toilets",
      "open air sala - big", "open air salas - big", "open air sala - medium", "open air salas - medium", singular&"s"),
    TEXT(units, "0")&" "&IF(units=1, singular, plural)
  ),
  param&IF(party<>"", " ("&party&")", "")
)`,
  },
  {
    name: 'ALL_EXPENSES',
    arguments: ['months'],
    description: 'Returns the complete 14-column expense dataset for numeric month dates in the supplied timeline, sorted by type, month, parameter and source row. Contains no header or total record.',
    formula: String.raw`=LET(
  selected_months, FILTER(months, ISNUMBER(months)),
  empty_record, {"","","","","","","","","","",""},
  records, REDUCE(empty_record, selected_months, LAMBDA(result, month,
      VSTACK(result, EXPENSES_FOR_MONTH(month))
  )),
  IF(SUM(ARRAYFORMULA(N(CHOOSECOLS(records, 1)<>"")))=0, {"","","","","","","","","","","","",""},
    LET(
      populated, FILTER(records, CHOOSECOLS(records, 1)<>""),
      type_order, {"In Loan Repayments";"Out Loan Disbursements";"Setup Expenses";"Land Expenses";"Infrastructure Expenses";"Construction Expenses";"Direct Operating Expenses";"Other Operating Expenses"},
      ranked, HSTACK(ARRAYFORMULA(MATCH(CHOOSECOLS(populated, 1), type_order, 0)), populated),
      details, CHOOSECOLS(SORT(ranked, 1, TRUE, 3, TRUE, 4, TRUE, 13, TRUE), 2,3,4,5,6,7,8,9,10,11,12,13),
      descriptions, MAP(CHOOSECOLS(details, 1), CHOOSECOLS(details, 2), CHOOSECOLS(details, 3), CHOOSECOLS(details, 4), LAMBDA(type, month, param, party, EXPENSE_DESCRIPTION(type, month, param, party))),
      source_params, MAP(CHOOSECOLS(details, 1), CHOOSECOLS(details, 3), LAMBDA(type, param,
          SWITCH(type, "In Loan Repayments", "in loans repayments", "Out Loan Disbursements", "out loans disbursements", param))),
      HSTACK(CHOOSECOLS(details, 1), ARRAYFORMULA(YEAR(CHOOSECOLS(details, 2))),
        ARRAYFORMULA("Q"&ROUNDUP(MONTH(CHOOSECOLS(details, 2))/3)), CHOOSECOLS(details, 2), descriptions,
        CHOOSECOLS(details, 5, 6, 7), source_params, CHOOSECOLS(details, 8, 9, 10, 11, 12))
    )
  )
)`,
  }
];
