// Preserve the sheet's named-function calculation separately from its formatted display.
const ALL_EXPENSES_FORMULA = String.raw`=LET(
  months, FILTER(R_MONTHS, ISNUMBER(R_MONTHS)),
  without_totals, LAMBDA(expenses, source,
    LET(
      categories, ARRAYFORMULA(IFNA(XLOOKUP(CHOOSECOLS(expenses, 2),
        COL_VALUES(source, "param"), COL_VALUES(source, "_category")), "")),
      IFNA(FILTER(expenses, ARRAYFORMULA(LOWER(TRIM(categories))<>"totals")), {"","","",""})
    )
  ),
  infra_details, LAMBDA(month,
    LET(
      source, S_INFRA_EXP_P,
      params, COL_VALUES(source, "param"),
      suppliers, COL_VALUES(source, "supplier"),
      categories, COL_VALUES(source, "_category"),
      values, CHOOSECOLS(DATA_ROWS(source), MATCH(month, HEADER_ROW(source), 0)),
      IFNA(FILTER(
        HSTACK(ARRAYFORMULA(IF(params<>"", "Infrastructure", "")), params, suppliers, values),
        ARRAYFORMULA(LOWER(TRIM(categories))<>"totals"),
        ISNUMBER(values),
        values>0
      ), {"","","",""})
    )
  ),
  financing_details, LAMBDA(month, source, param_name, party_header, type_label,
    LET(
      params, COL_VALUES(source, "param"),
      names, COL_VALUES(source, "_name"),
      parties, ARRAYFORMULA(IFNA(XLOOKUP(names, COL_VALUES(source, "name"),
        COL_VALUES(source, party_header)), "")),
      values, CHOOSECOLS(DATA_ROWS(source), MATCH(month, HEADER_ROW(source), 0)),
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
      operating, without_totals(OPERATIONAL_EXPENSES_D(month), S_OPERATIONAL_EXP),
      all_expenses, VSTACK(
        financing_details(month, S_BORROWINGS, "borrowings repaid", "lender", "Borrowing Repayments"),
        financing_details(month, S_LOANS, "loan disbursed", "debtor", "Loan Disbursements"),
        without_totals(SETUP_EXPENSES_D(month), S_SETUP_EXP_P),
        LAND_EXPENSES_D(month),
        infra_details(month),
        without_totals(CONSTRUCTION_EXPENSES_D(month), S_CONSTRUCTION_EXP_P),
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
  type_order, {"Borrowing Repayments";"Loan Disbursements";"Setup Expenses";"Land Expenses";"Infrastructure Expenses";"Construction Expenses";"Direct Operating Expenses";"Other Operating Expenses"},
  ranked, HSTACK(ARRAYFORMULA(IFNA(MATCH(type_labels, type_order, 0), 9)), type_labels, CHOOSECOLS(grouped, 2, 3, 4, 5)),
  details, SORT(ranked, 1, TRUE, 3, TRUE, 4, TRUE),
  item_values, MAP(CHOOSECOLS(details, 2), CHOOSECOLS(details, 3), CHOOSECOLS(details, 4), CHOOSECOLS(details, 5),
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
    HSTACK(CHOOSECOLS(details, 2),
      ARRAYFORMULA(YEAR(CHOOSECOLS(details, 3))),
      ARRAYFORMULA("Q"&ROUNDUP(MONTH(CHOOSECOLS(details, 3))/3)&" "&YEAR(CHOOSECOLS(details, 3))),
      CHOOSECOLS(details, 3),
      item_values, CHOOSECOLS(details, 6)))
)`;
