// Execute these cases in native Google Sheets against the installed named functions.
const headers = ['Budget Category', 'Budget Line', 'Year', 'Status', 'Cost', 'Contingency'];
const records = [
  ['Setup', 'Legal', 2026, 'Actual', 100, 0],
  ['Setup', 'Legal', 2026, 'Actual', -25, 0],
  ['Setup', 'Legal', 2026, 'Planned', 200, 10],
  ['Setup', 'Accounting', 2026, 'Planned', 50, 5],
  ['Setup', 'Legal', 2027, 'Planned', 900, 90],
  ['Out Loans', 'Out Loan Disbursements', 2028, 'Planned', 300, 0],
];
const quote = value => typeof value === 'number' ? String(value) : '"' + value.replaceAll('"', '""') + '"';
const table = rows => '{' + [headers, ...rows].map(row => row.map(quote).join(',')).join(';') + '}';
function amount({rows = records, year = 2026, category = 'Setup', line = 'Legal', column = 'Past Expenses', error} = {}) {
  let source = table(rows);
  if (error) source = source.replace(',100,0', ',' + error + ',0');
  return `BUDGET_REPORT_AMOUNT(BUDGET_EXPENSE_TOTALS(${source}),${year},${quote(category)},${quote(line)},${quote(column)})`;
}
const blank = ['', '', '', '', '', ''];
const cases = [
  {name: 'Actual costs include duplicate rows and signed credits', formula: `=${amount()}=75`},
  {name: 'Planned costs exclude Actual and other years', formula: `=${amount({column: 'Future Expenses'})}=200`},
  {name: 'Category subtotal includes all its lines', formula: `=${amount({line: '', column: 'Future Expenses'})}=250`},
  {name: 'Year subtotal includes all categories', formula: `=${amount({category: '', line: '', column: 'Future Expenses'})}=250`},
  {name: 'Contingency stays separate from cost', formula: `=${amount({line: '', column: 'Contingency'})}=15`},
  {name: 'Missing line returns zero', formula: `=${amount({line: 'Missing'})}=0`},
  {name: 'Plural Out Loans alias matches canonical line', formula: `=${amount({year: 2028, category: 'Out Loans', line: 'out loans disbursements', column: 'Future Expenses'})}=300`},
  {name: 'Source NA propagates', formula: `=ISNA(${amount({error: 'NA()'})})`},
  {name: 'Source divide-by-zero propagates', formula: `=ISERROR(${amount({error: '1/0'})})`},
  {name: 'Empty table returns zero', formula: `=${amount({rows: [blank]})}=0`},
  {name: 'Gaps and later records are retained', formula: `=${amount({rows: [blank, records[0], blank, records[1]]})}=75`},
  {name: 'Invalid numeric amount remains visible', formula: `=ISNA(${amount({error: '"invalid"'})})`},
  {name: 'Report preserves zero-cost catalog lines and annual rows', formula: '=AND(ROWS(BUDGET_ROW_LAYOUT())=113,COUNTIF(CHOOSECOLS(BUDGET_ROW_LAYOUT(),4),"Year")=4,COUNTIF(CHOOSECOLS(BUDGET_ROW_LAYOUT(),3),"Road")=3)'},
];
module.exports = cases;
if (require.main === module) process.stdout.write(JSON.stringify(cases, null, 2));
