const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');
const ctx = vm.createContext({});
for (const file of ['AllExpenses.js', 'AllExpensesData.js']) vm.runInContext(fs.readFileSync(`${__dirname}/../src/${file}`, 'utf8'), ctx);
const plain = value => JSON.parse(JSON.stringify(value));
const months = [46023, 46054];
const headers = ['name', 'lender', 'debtor', 'param', '_name', 'supplier', 'creditor', '_category', ...months];
const row = values => headers.map(header => values[header] ?? '');
function sources() {
  return Object.fromEntries(['R_MONTHS', 'S_BORROWINGS', 'S_LOANS', 'S_SETUP_EXP_P', 'S_OBLIGATIONS',
    'S_INFRA_EXP_P', 'S_CONSTRUCTION_EXP_P', 'S_OPERATIONAL_EXP'].map(name => [name, [name === 'R_MONTHS' ? months : headers]]));
}

test('all eight expense categories retain their source rules, grouping and month labels', () => {
  const input = sources();
  input.S_BORROWINGS.push(row({name: 'Loan A', lender: 'Bank', param: 'start balance', _name: 'Loan A'}),
    row({param: 'borrowings repaid', _name: 'Loan A', 46023: 10}),
    row({param: 'borrowings repaid', _name: 'TOTALS', 46023: 999}));
  input.S_LOANS.push(row({name: 'Loan B', debtor: 'Borrower', param: 'loan disbursed', _name: 'Loan B', 46023: 40}));
  input.S_SETUP_EXP_P.push(row({param: 'License', supplier: 'Agent', _category: 'Legal', 46023: 5, 46054: -1}),
    row({param: 'License', supplier: 'Agent', _category: 'Legal', 46023: 5}),
    row({param: 'Summary', _category: ' TOTALS ', 46023: 999}),
    row({param: 'total costs', _category: 'Legal', 46023: 999}));
  input.S_OBLIGATIONS.push(row({name: 'Plot', creditor: 'Seller', _category: 'lAnD', 46023: 15}),
    row({name: 'Equipment', _category: 'Other', 46023: 999}));
  input.S_INFRA_EXP_P.push(row({param: 'Road', supplier: 'Builder', _category: 'Infra', 46023: 20}),
    row({param: 'Summary', _category: 'totals', 46023: 999}));
  input.S_CONSTRUCTION_EXP_P.push(row({param: 'clinics const. cost', _category: 'Clinic', 46023: 50, 46054: 60}),
    row({param: 'clinics const. count', _category: 'Clinic', 46023: 1, 46054: 2}),
    row({param: 'complex room const. cost', _category: 'Homes', 46023: 30}),
    row({param: 'complex room const. count', _category: 'Homes', 46023: 6}));
  input.S_OPERATIONAL_EXP.push(row({param: 'Staff', _category: 'Direct Operating Expenses', 46023: 25}),
    row({param: 'Admin', _category: 'Other Operating Expenses', 46054: 30}),
    row({param: 'Summary', _category: 'totals', 46023: 999}),
    row({param: 'Other category', _category: 'General', 46023: 999}));
  const output = plain(ctx._allExpensesLayout(ctx._allExpensesFromSources(input))).rows.slice(0, -1);
  assert.deepEqual(output.map(r => [r[0], r[3], r[4], r[5]]), [
    ['Borrowing Repayments', 46023, 'Loan A (Bank)', 10],
    ['Loan Disbursements', 46023, 'Loan B (Borrower)', 40],
    ['Setup Expenses', 46023, 'License (Agent)', 10],
    ['Land Expenses', 46023, 'Plot (Seller)', 15],
    ['Infrastructure Expenses', 46023, 'Road (Builder)', 20],
    ['Construction Expenses', 46023, '1 clinic', 50],
    ['Construction Expenses', 46023, '6 complex homes', 30],
    ['Construction Expenses', 46054, '2 clinics', 60],
    ['Direct Operating Expenses', 46023, 'Staff', 25],
    ['Other Operating Expenses', 46054, 'Admin', 30],
  ]);
  assert.ok(output.every(r => r[1] === 2026 && r[2] === 'Q1 2026'));
});

test('no positive expenses produce an empty detail list; numeric-only categories ignore nonnumeric cells', () => {
  const input = sources();
  input.S_INFRA_EXP_P.push(row({param: 'Road', _category: 'Infra', 46023: 0, 46054: '#REF!'}));
  input.S_OPERATIONAL_EXP.push(row({param: 'Staff', _category: 'Direct Operating Expenses', 46023: 'pending', 46054: -10}));
  assert.deepEqual(plain(ctx._allExpensesFromSources(input)), []);
});

test('source structure and construction unit counts are validated before any output is written', () => {
  const missingSource = sources();
  delete missingSource.S_LOANS;
  assert.throws(() => ctx._allExpensesFromSources(missingSource), /Missing expense source/);
  const missingMonth = sources();
  missingMonth.S_LOANS = [headers.slice(0, -1)];
  assert.throws(() => ctx._allExpensesFromSources(missingMonth), /Missing month/);
  const missingCount = sources();
  missingCount.S_CONSTRUCTION_EXP_P.push(row({param: 'clinics const. cost', _category: 'Clinic', 46023: 50}));
  assert.throws(() => ctx._allExpensesFromSources(missingCount), /Missing planned unit count/);
});

test('construction plurals and rounded counts retain the existing labels', () => {
  assert.equal(ctx._allExpensesConstructionItem('lobby const. cost', 2), '2 lobbies');
  assert.equal(ctx._allExpensesConstructionItem('sala toilets const. cost', 1), '1 sala toilet');
  assert.equal(ctx._allExpensesConstructionItem('storage const. cost', 2), '2 storage');
  assert.equal(ctx._allExpensesConstructionItem('clinics const. cost', 2.5), '3 clinics');
});

test('source validation failure leaves the displayed table untouched', () => {
  let writes = 0;
  const sheet = {properties: {sheetId: 1, title: 'All Expenses', gridProperties: {rowCount: 55}}};
  const runtime = vm.createContext({console: {log() {}},
    SpreadsheetApp: {getActiveSpreadsheet: () => ({getId: () => 'copy-id'})},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock() {}})},
    Sheets: {Spreadsheets: {
      get: () => ({sheets: [sheet]}),
      Values: {batchGet: () => ({valueRanges: []})},
      batchUpdate: () => { writes++; },
    }},
  });
  for (const file of ['AllExpenses.js', 'AllExpensesData.js']) vm.runInContext(fs.readFileSync(`${__dirname}/../src/${file}`, 'utf8'), runtime);
  assert.throws(() => runtime.refreshAllExpenses(), /timeline has no months/);
  assert.equal(writes, 0);
});
