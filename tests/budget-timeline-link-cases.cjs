// Execute these generated boundary checks in native Sheets.
const fs = require('node:fs');
const vm = require('node:vm');
const context = vm.createContext({});
vm.runInContext(fs.readFileSync(`${__dirname}/../src/BudgetTimelineLinkFormula.js`, 'utf8'), context);
const formula = vm.runInContext('BUDGET_TIMELINE_LINK_DEFINITION.formula', context);
const setup = formula.slice(1, formula.indexOf('  IF(month_count=0'));
const cases = [
  [2026, 11, 2026, 'Past', 2026, 1, 10],
  [2026, 11, 2026, 'Future', 2026, 11, 2],
  [2026, 11, 2026, 'Full Year', 2026, 1, 12],
  [2026, 1, 2026, 'Past', 2026, 1, 0],
  [2026, 1, 2026, 'Future', 2026, 1, 12],
  [2026, 12, 2026, 'Past', 2026, 1, 11],
  [2026, 12, 2026, 'Future', 2026, 12, 1],
  [2026, 11, 2025, 'Past', 2025, 1, 12],
  [2026, 11, 2025, 'Future', 2026, 11, 0],
  [2026, 11, 2027, 'Past', 2027, 1, 0],
  [2026, 11, 2027, 'Future', 2027, 1, 12],
  [2027, 1, 2026, 'Past', 2026, 1, 12],
].map(([currentYear, currentMonth, year, period, startYear, startMonth, count]) => ({
  name: `${period} ${year} as of ${currentYear}-${currentMonth}`,
  formula: `=LET(budget_year,${year},period,"${period}",${setup.replace('MONTH_START(TODAY())', `DATE(${currentYear},${currentMonth},1)`)}AND(start_month=DATE(${startYear},${startMonth},1),month_count=${count})))`,
}));
cases.push({
  name: 'Invalid period remains an error',
  formula: `=ISNA(LET(budget_year,2026,period,"Invalid",${setup}month_count)))`,
});
module.exports = cases;
if (require.main === module) process.stdout.write(JSON.stringify(cases, null, 2));
