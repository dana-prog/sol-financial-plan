const BUDGET_TIMELINE_LINK_DEFINITION = {
  name: 'BUDGET_TIMELINE_LINK',
  arguments: ['source_table', 'sheet_id', 'lookup_col', 'lookup_value', 'budget_year', 'period', 'amount'],
  description: 'Links an amount to source rows within the budget year. Period is Past (before the current month), Future (current month onward), or Full Year. Empty periods return zero.',
  formula: String.raw`=LET(
  cutoff, MONTH_START(TODAY()),
  year_start, DATE(budget_year, 1, 1),
  past_months, MAX(0, MIN(12, (YEAR(cutoff)-budget_year)*12+MONTH(cutoff)-1)),
  month_count, SWITCH(period, "Past", past_months, "Future", 12-past_months, "Full Year", 12, NA()),
  start_month, IF(period="Future", MAX(year_start, cutoff), year_start),
  IF(month_count=0, 0,
    LET(
      keys, COL_VALUES(source_table, lookup_col),
      first_row, ROW(INDEX(source_table, 1, 1)) + MATCH(lookup_value, keys, 0),
      last_row, ROW(INDEX(source_table, 1, 1)) + LOOKUP(2, ARRAYFORMULA(1 / (keys = lookup_value)), SEQUENCE(ROWS(keys))),
      first_col, COLUMN(INDEX(source_table, 1, 1)) - 1 + MATCH(start_month, INDEX(source_table, 1, 0), 0),
      HYPERLINK(
        "#gid=" & sheet_id & "&range=" & ADDRESS(first_row, first_col, 4) & ":" & ADDRESS(last_row, first_col + month_count - 1, 4),
        amount
      )
    )
  )
)`,
};
