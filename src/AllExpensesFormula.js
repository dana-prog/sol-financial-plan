// Preserve the sheet's named-function calculation separately from its formatted display.
const ALL_EXPENSES_FORMULA = String.raw`=LET(
  months, FILTER(R_MONTHS, ISNUMBER(R_MONTHS)),
  setup_params, COL_VALUES(S_SETUP_EXP_P, "param"),
  setup_categories, COL_VALUES(S_SETUP_EXP_P, "_category"),
  construction_params, COL_VALUES(S_CONSTRUCTION_EXP_P, "param"),
  construction_categories, COL_VALUES(S_CONSTRUCTION_EXP_P, "_category"),
  operating_params, COL_VALUES(S_OPERATIONAL_EXP, "param"),
  operating_categories, COL_VALUES(S_OPERATIONAL_EXP, "_category"),
  infra_params, COL_VALUES(S_INFRA_EXP_P, "param"),
  infra_suppliers, COL_VALUES(S_INFRA_EXP_P, "supplier"),
  infra_categories, ARRAYFORMULA(LOWER(TRIM(COL_VALUES(S_INFRA_EXP_P, "_category")))),
  infra_data, DATA_ROWS(S_INFRA_EXP_P),
  infra_headers, HEADER_ROW(S_INFRA_EXP_P),
  borrowing_params, COL_VALUES(S_BORROWINGS, "param"),
  borrowing_names, COL_VALUES(S_BORROWINGS, "_name"),
  borrowing_parties, ARRAYFORMULA(IFNA(XLOOKUP(borrowing_names, COL_VALUES(S_BORROWINGS, "name"),
    COL_VALUES(S_BORROWINGS, "lender")), "")),
  borrowing_data, DATA_ROWS(S_BORROWINGS),
  borrowing_headers, HEADER_ROW(S_BORROWINGS),
  loan_params, COL_VALUES(S_LOANS, "param"),
  loan_names, COL_VALUES(S_LOANS, "_name"),
  loan_parties, ARRAYFORMULA(IFNA(XLOOKUP(loan_names, COL_VALUES(S_LOANS, "name"),
    COL_VALUES(S_LOANS, "debtor")), "")),
  loan_data, DATA_ROWS(S_LOANS),
  loan_headers, HEADER_ROW(S_LOANS),
  without_totals, LAMBDA(expenses, params, source_categories,
    LET(
      categories, ARRAYFORMULA(IFNA(XLOOKUP(CHOOSECOLS(expenses, 2),
        params, source_categories), "")),
      IFNA(FILTER(expenses, ARRAYFORMULA(LOWER(TRIM(categories))<>"totals")), {"","","",""})
    )
  ),
  infra_details, LAMBDA(month,
    LET(
      values, CHOOSECOLS(infra_data, MATCH(month, infra_headers, 0)),
      IFNA(FILTER(
        HSTACK(ARRAYFORMULA(IF(infra_params<>"", "Infrastructure", "")), infra_params, infra_suppliers, values),
        infra_categories<>"totals",
        ISNUMBER(values),
        values>0
      ), {"","","",""})
    )
  ),
  financing_details, LAMBDA(month, params, names, parties, source_data, headers, param_name, type_label,
    LET(
      values, CHOOSECOLS(source_data, MATCH(month, headers, 0)),
      IFNA(FILTER(
        HSTACK(ARRAYFORMULA(IF(names<>"", type_label, "")), names, parties, values),
        params=param_name,
        ARRAYFORMULA(LOWER(TRIM(names))<>"totals"),
        ISNUMBER(values),
        values>0
      ), {"","","",""})
    )
  ),
  monthly_details, LAMBDA(month,
    LET(
      operating, without_totals(OPERATIONAL_EXPENSES_D(month), operating_params, operating_categories),
      all_expenses, VSTACK(
        financing_details(month, borrowing_params, borrowing_names, borrowing_parties, borrowing_data, borrowing_headers, "borrowings repaid", "Borrowing Repayments"),
        financing_details(month, loan_params, loan_names, loan_parties, loan_data, loan_headers, "loan disbursed", "Loan Disbursements"),
        without_totals(SETUP_EXPENSES_D(month), setup_params, setup_categories),
        LAND_EXPENSES_D(month),
        infra_details(month),
        without_totals(CONSTRUCTION_EXPENSES_D(month), construction_params, construction_categories),
        IFNA(FILTER(operating, ARRAYFORMULA(
          (CHOOSECOLS(operating, 1)="Direct Operating Expenses")+
          (CHOOSECOLS(operating, 1)="Other Operating Expenses"))), {"","","",""})
      ),
      IFNA(FILTER(all_expenses, CHOOSECOLS(all_expenses, 1)<>""), {"","","",""})
    )
  ),
  expenses, REDUCE({"",0,"","",0}, months,
    LAMBDA(result, month,
      LET(
        monthly, monthly_details(month),
        VSTACK(result, HSTACK(
          CHOOSECOLS(monthly, 1),
          ARRAYFORMULA(IF(CHOOSECOLS(monthly, 1)<>"", month, "")),
          CHOOSECOLS(monthly, 2, 3, 4)
        ))
      )
    )),
  grouped, IFNA(QUERY(expenses,
    "select Col1, Col2, Col3, Col4, sum(Col5) where Col1 is not null group by Col1, Col2, Col3, Col4 label sum(Col5) ''", 0),
    {"","","","",0}),
  type_labels, MAP(CHOOSECOLS(grouped, 1), LAMBDA(type,
    SWITCH(type, "Construction", "Construction Expenses", "Infrastructure", "Infrastructure Expenses",
      "Setup", "Setup Expenses", "Land", "Land Expenses", type)
  )),
  details, HSTACK(type_labels, CHOOSECOLS(grouped, 2, 3, 4, 5)),
  item_values, MAP(CHOOSECOLS(details, 1), CHOOSECOLS(details, 2), CHOOSECOLS(details, 3), CHOOSECOLS(details, 4),
    LAMBDA(type, month, param, detail,
      IF(type="Construction Expenses",
        LET(
          units, INDEX(CONSTRUCTION_PARAM_VALUES(REGEXREPLACE(param, " const[.] cost$", " const. count")), 1, MATCH(month, R_MONTHS, 0)),
          unit_name, REGEXREPLACE(param, " const[.] cost$", ""),
          singular, SWITCH(unit_name, "clinics", "clinic", "sala toilets", "sala toilet", "complex room", "complex home", unit_name),
          plural, SWITCH(unit_name,
            "lobby", "lobbies", "storage", "storage", "pool lounge seating", "pool lounge seating",
            "general resort storage", "general resort storage", "shower & toilet", "showers & toilets",
            "open air sala - big", "open air salas - big", "open air sala - medium", "open air salas - medium",
            singular&"s"),
          TEXT(units, "0")&" "&IF(units=1, singular, plural)
        ),
        param&IF(detail<>"", " ("&detail&")", "")
      )
    )),
  IF(SUM(ARRAYFORMULA(N(type_labels<>"")))=0, {"__NO_EXPENSES__","","","","",""},
    HSTACK(CHOOSECOLS(details, 1),
      ARRAYFORMULA(YEAR(CHOOSECOLS(details, 2))),
      ARRAYFORMULA("Q"&ROUNDUP(MONTH(CHOOSECOLS(details, 2))/3)&" "&YEAR(CHOOSECOLS(details, 2))),
      CHOOSECOLS(details, 2),
      item_values, CHOOSECOLS(details, 5)))
)`;
