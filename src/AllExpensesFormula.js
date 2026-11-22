const ALL_EXPENSES_HEADERS = [
  'Type', 'Year', 'Quarter', 'Month', 'Param', 'Cost', 'Budget category',
  'Budget line', 'Source parameter', 'Source sheet', 'Basis', 'Contingency', 'Total cost', 'Source row',
];

// Keep this formula live: report consumers must never depend on a manually refreshed snapshot.
const ALL_EXPENSES_FORMULA = String.raw`=LET(
  months, FILTER(R_MONTHS, ISNUMBER(R_MONTHS)),
  empty_details, {"","","","","",""},
  empty_record, {"","","","","","","","","","",""},
  numeric_details, LAMBDA(details, eligible,
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
  ),
  infra_params, COL_VALUES(S_INFRA_EXP_P, "param"),
  infra_categories, COL_VALUES(S_INFRA_EXP_P, "_category"),
  infra_suppliers, COL_VALUES(S_INFRA_EXP_P, "supplier"),
  infra_data, DATA_ROWS(S_INFRA_EXP_P),
  infra_headers, HEADER_ROW(S_INFRA_EXP_P),
  infra_rows, SEQUENCE(ROWS(infra_params), 1, ROW(INDEX(S_INFRA_EXP_P, 1, 1)) + 1),
  infra_details, LAMBDA(month,
    numeric_details(
      HSTACK(ARRAYFORMULA(IF(infra_params<>"", "Infrastructure Expenses", "")), infra_params,
        infra_suppliers, CHOOSECOLS(infra_data, MATCH(month, infra_headers, 0)), infra_rows, infra_categories),
      ARRAYFORMULA((infra_params<>"")*(LOWER(TRIM(infra_categories))<>"totals"))
    )
  ),
  financing_details, LAMBDA(month, source, param_name, party_header, type_label,
    LET(
      params, COL_VALUES(source, "param"),
      names, COL_VALUES(source, "_name"),
      parties, ARRAYFORMULA(IFNA(XLOOKUP(names, COL_VALUES(source, "name"), COL_VALUES(source, party_header)), "")),
      values, CHOOSECOLS(DATA_ROWS(source), MATCH(month, HEADER_ROW(source), 0)),
      source_rows, SEQUENCE(ROWS(params), 1, ROW(INDEX(source, 1, 1)) + 1),
      numeric_details(
        HSTACK(ARRAYFORMULA(IF(names<>"", type_label, "")), names, parties, values, source_rows, params),
        ARRAYFORMULA((params=param_name)*(names<>"")*(LOWER(TRIM(names))<>"totals"))
      )
    )
  ),
  decorate, LAMBDA(expenses, month, source_name, category, basis, rate,
    LET(
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
    )
  ),
  records, REDUCE(empty_record, months, LAMBDA(result, month,
    LET(
      actual_setup, month < MONTH_START(TODAY()),
      VSTACK(result,
        decorate(financing_details(month, S_IN_LOANS, "in loans repayments", "lender", "In Loan Repayments"),
          month, "In Loans", "In Loans", "Scheduled", 0),
        decorate(financing_details(month, S_OUT_LOANS, "out loans disbursements", "debtor", "Out Loan Disbursements"),
          month, "Out Loans", "Out Loans", "Scheduled", 0),
        decorate(SETUP_EXPENSES_D(month), month, IF(actual_setup, "Setup Exp (A)", "Setup Exp (P)"),
          "Setup", IF(actual_setup, "Actual", "Planned"), IF(actual_setup, 0, PARAM("setup contingency"))),
        decorate(LAND_EXPENSES_D(month), month, "Obligations (P)", "Land", "Scheduled", 0),
        decorate(infra_details(month), month, "Infra Exp (P)", "Infra", "Planned", PARAM("infra contingency")),
        decorate(CONSTRUCTION_EXPENSES_D(month), month, "Construction Exp (P)", "Construction", "Planned", PARAM("construction contingency")),
        decorate(OPERATIONAL_EXPENSES_D(month), month, "Operational Exp (P)", "Operations", "Planned", PARAM("operating contingency"))
      )
    )
  )),
  IF(SUM(ARRAYFORMULA(N(CHOOSECOLS(records, 1)<>"")))=0, {"","","","","","","","","","","","",""},
    LET(
      populated, FILTER(records, CHOOSECOLS(records, 1)<>""),
      type_order, {"In Loan Repayments";"Out Loan Disbursements";"Setup Expenses";"Land Expenses";"Infrastructure Expenses";"Construction Expenses";"Direct Operating Expenses";"Other Operating Expenses"},
      ranked, HSTACK(ARRAYFORMULA(MATCH(CHOOSECOLS(populated, 1), type_order, 0)), populated),
      details, CHOOSECOLS(SORT(ranked, 1, TRUE, 3, TRUE, 4, TRUE, 13, TRUE), 2,3,4,5,6,7,8,9,10,11,12,13),
      descriptions, MAP(CHOOSECOLS(details, 1), CHOOSECOLS(details, 2), CHOOSECOLS(details, 3), CHOOSECOLS(details, 4),
        LAMBDA(type, month, param, party,
          IF(type="Construction Expenses",
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
          )
        )
      ),
      source_params, MAP(CHOOSECOLS(details, 1), CHOOSECOLS(details, 3), LAMBDA(type, param,
        SWITCH(type, "In Loan Repayments", "in loans repayments", "Out Loan Disbursements", "out loans disbursements", param))),
      HSTACK(CHOOSECOLS(details, 1), ARRAYFORMULA(YEAR(CHOOSECOLS(details, 2))),
        ARRAYFORMULA("Q"&ROUNDUP(MONTH(CHOOSECOLS(details, 2))/3)), CHOOSECOLS(details, 2), descriptions,
        CHOOSECOLS(details, 5, 6, 7), source_params, CHOOSECOLS(details, 8, 9, 10, 11, 12))
    )
  )
)`;
