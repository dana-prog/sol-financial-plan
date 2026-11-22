// Run these generated formulas in native Sheets; Node does not emulate Sheets calculation.
const fs = require('node:fs');
const vm = require('node:vm');
const context = vm.createContext({});
for (const file of ['BudgetFormula.js', 'AllExpensesFormula.js']) {
  vm.runInContext(fs.readFileSync(`${__dirname}/../src/${file}`, 'utf8'), context);
}
const budget = vm.runInContext('BUDGET_EXPENSE_SUM_DEFINITION.formula', context).slice(1);
const expenses = vm.runInContext('ALL_EXPENSES_NAMED_FUNCTIONS.find(f => f.name === "ALL_EXPENSES_ROWS").formula', context);
const quote = value => typeof value === 'number' ? String(value) : '"' + value.replaceAll('"', '""') + '"';
const table = rows => '{' + rows.map(row => row.map(quote).join(',')).join(';') + '}';
const headers = ['Budget Category', 'Budget Line', 'Year', 'Basis', 'Cost', 'Contingency'];
const records = [
  ['Setup', 'Legal', 2026, 'Actual', 100, 0],
  ['Setup', 'Legal', 2026, 'Actual', -25, 0],
  ['Setup', 'Legal', 2026, 'Planned', 200, 10],
  ['Setup', 'Accounting', 2026, 'Planned', 50, 5],
  ['Setup', 'Legal', 2027, 'Planned', 900, 90],
  ['Out Loans', 'Out Loan Disbursements', 2028, 'Planned', 300, 0],
];

function aggregation({category = 'Setup', line = 'Legal', year = 2026,
  basis = 'Actual', amount = 'Cost', rows = records, error} = {}) {
  let source = table([headers, ...rows]);
  if (error) source = source.replace(',100,0', `,${error},0`);
  return budget
    .replaceAll('BUDGET_VALUE("_category")', quote(category))
    .replaceAll('BUDGET_VALUE("_year")', String(year))
    .replaceAll('BUDGET_VALUE("budget line")', quote(line))
    .replaceAll('ALL_EXPENSES', source)
    .replace(/\bbasis\b/g, quote(basis))
    .replace(/\bamount_column\b/g, quote(amount));
}

const cases = [
  {name: 'Actual costs include duplicate rows and signed credits', formula: `=${aggregation()}=75`},
  {name: 'Planned costs exclude Actual and other years', formula: `=${aggregation({basis: 'Planned'})}=200`},
  {name: 'Blank budget line aggregates the category', formula: `=${aggregation({basis: 'Planned', line: ''})}=250`},
  {name: 'Contingency stays separate from base cost', formula: `=${aggregation({basis: '', line: '', amount: 'Contingency'})}=15`},
  {name: 'Missing budget line returns zero', formula: `=${aggregation({line: 'Missing'})}=0`},
  {name: 'Legacy plural Out Loans label matches the canonical expense line', formula: `=${aggregation({category: 'Out Loans', line: 'out loans disbursements', year: 2028, basis: 'Planned'})}=300`},
  {name: 'Source NA errors remain visible', formula: `=ISNA(${aggregation({error: 'NA()'})})`},
  {name: 'Other source errors remain visible', formula: `=ISERROR(${aggregation({error: '1/0'})})`},
  {name: 'Header and empty capacity produce zero', formula: `=${aggregation({rows: [['', '', '', '', '', '']]})}=0`},
  {name: 'Invalid amount header is not silently ignored', formula: `=ISNA(${aggregation({amount: 'Missing'})})`},
];

// Exercise the exact eligibility expression from the production formula against
// both bases on each side of the current-month boundary, without editing source data.
const eligibility = expenses.split('  eligible, ')[1].split(',\n  IF(SUM(eligible)')[0];
const fixtureRows = [
  ['Actual', 'EDATE(MONTH_START(TODAY()), -1)', 100],
  ['Planned', 'EDATE(MONTH_START(TODAY()), -1)', 200],
  ['Actual', 'MONTH_START(TODAY())', 300],
  ['Planned', 'MONTH_START(TODAY())', 400],
  ['Planned', 'EDATE(MONTH_START(TODAY()), 1)', 500],
  ['Actual', 'EDATE(MONTH_START(TODAY()), 1)', 600],
].map(([basis, month, cost]) => `{"Expense",${month},"Param","",${cost},"Setup","Legal","Source",${quote(basis)},0,${cost},2}`);
cases.push({
  name: 'Past Actual and current/future Planned only; no fallback to past plans',
  formula: `=LET(records,VSTACK(${fixtureRows.join(',')}),eligible,${eligibility},AND(SUM(eligible)=3,SUM(FILTER(CHOOSECOLS(records,5),eligible))=1000,INDEX(eligible,2)=0))`,
});
cases.push({
  name: 'Native empty-month generator returns exactly the title row',
  formula: '=AND(ROWS(ALL_EXPENSES_ROWS(DATE(2026,1,1)))=1,COLUMNS(ALL_EXPENSES_ROWS(DATE(2026,1,1)))=13)',
});

module.exports = cases;
if (require.main === module) process.stdout.write(JSON.stringify(cases, null, 2));
