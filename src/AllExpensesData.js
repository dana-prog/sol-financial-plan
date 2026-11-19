const ALL_EXPENSES_SOURCE_RANGES = [
  'R_MONTHS', 'S_BORROWINGS', 'S_LOANS', 'S_SETUP_EXP_P', 'S_OBLIGATIONS',
  'S_INFRA_EXP_P', 'S_CONSTRUCTION_EXP_P', 'S_OPERATIONAL_EXP', 'S_SETUP_EXP_A',
];

function _allExpensesFromSources(inputs, currentMonth) {
  if (!Number.isFinite(currentMonth)) throw new Error('Invalid current month. Existing rows were preserved.');
  const months = (inputs.R_MONTHS?.[0] || []).filter(value => typeof value === 'number' && Number.isFinite(value));
  if (!months.length) throw new Error('The expense timeline has no months. Existing rows were preserved.');
  const grouped = new Map();
  const sources = [
    ['S_BORROWINGS', 'Borrowing Repayments', 'financing'],
    ['S_LOANS', 'Loan Disbursements', 'financing'],
    ['S_SETUP_EXP_P', 'Setup Expenses', 'setup'],
    ['S_SETUP_EXP_A', 'Setup Expenses', 'setup'],
    ['S_OBLIGATIONS', 'Land Expenses', 'land'],
    ['S_INFRA_EXP_P', 'Infrastructure Expenses', 'infra'],
    ['S_CONSTRUCTION_EXP_P', 'Construction Expenses', 'construction'],
    ['S_OPERATIONAL_EXP', '', 'operating'],
  ];
  for (const [source, defaultType, kind] of sources) {
    const sourceMonths = kind === 'setup' ? months.filter(month =>
      source === 'S_SETUP_EXP_A' ? month < currentMonth : month >= currentMonth) : months;
    if (!sourceMonths.length) continue;
    const data = inputs[source];
    if (!data?.length) throw new Error(`Missing expense source: ${source}. Existing rows were preserved.`);
    const headers = data[0];
    const column = name => {
      const index = headers.findIndex(value => String(value).toLowerCase() === name);
      if (index < 0) throw new Error(`Missing ${name} column in ${source}. Existing rows were preserved.`);
      return index;
    };
    const monthColumns = sourceMonths.map(month => {
      const index = headers.indexOf(month);
      if (index < 0) throw new Error(`Missing month ${month} in ${source}. Existing rows were preserved.`);
      return index;
    });
    const rows = data.slice(1);
    const paramColumn = column(kind === 'land' ? 'name' : 'param');
    const inheritCategory = source === 'S_SETUP_EXP_A' && !headers.some(value => String(value).toLowerCase() === '_category');
    const categoryColumn = kind === 'financing' ? -1 : column(inheritCategory ? 'category' : '_category');
    const detailColumn = ['setup', 'infra'].includes(kind) ? column('supplier') : kind === 'land' ? column('creditor') : -1;
    const nameColumn = kind === 'financing' ? column('_name') : paramColumn;
    const parties = new Map();
    const params = new Map();
    if (kind === 'financing') {
      const loanNameColumn = column('name');
      const partyColumn = column(source === 'S_BORROWINGS' ? 'lender' : 'debtor');
      for (const row of rows) {
        const key = String(row[loanNameColumn] ?? '').toLowerCase();
        if (!parties.has(key)) parties.set(key, String(row[partyColumn] ?? ''));
      }
    }
    if (kind === 'construction') {
      for (const row of rows) {
        const key = String(row[paramColumn] ?? '').toLowerCase();
        if (!params.has(key)) params.set(key, row);
      }
    }
    let previousCategory = '';
    for (const row of rows) {
      const param = String(row[paramColumn] ?? '');
      const name = String(row[nameColumn] ?? '');
      const category = String(row[categoryColumn] || (inheritCategory ? previousCategory : ''));
      previousCategory = category;
      const normalizedCategory = category.trim().toLowerCase();
      if (!name || normalizedCategory === 'totals') continue;
      if (kind === 'financing' && (param.toLowerCase() !== (source === 'S_BORROWINGS' ? 'borrowings repaid' : 'loan disbursed') || name.trim().toLowerCase() === 'totals')) continue;
      if (kind === 'setup' && param.toLowerCase() === 'total costs') continue;
      if (kind === 'land' && normalizedCategory !== 'land') continue;
      if (kind === 'construction' && !/const\. cost$/.test(param)) continue;
      if (kind === 'operating' && !['Direct Operating Expenses', 'Other Operating Expenses'].includes(category)) continue;
      const type = kind === 'operating' ? category : defaultType;
      const detail = kind === 'financing' ? parties.get(name.toLowerCase()) || '' : String(row[detailColumn] ?? '');
      sourceMonths.forEach((month, index) => {
        const value = row[monthColumns[index]];
        if (['setup', 'construction'].includes(kind) && typeof value === 'string' && /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|ERROR!)/.test(value)) {
          throw new Error(`Expense source ${source}: ${name} contains ${value} in month ${month}. Existing rows were preserved.`);
        }
        if (typeof value !== 'number' || value <= 0) return;
        if (!Number.isFinite(value)) throw new Error(`Invalid cost for ${name} in ${source}. Existing rows were preserved.`);
        const key = JSON.stringify([type, month, name, detail]);
        const existing = grouped.get(key);
        if (existing) {
          existing[5] += value;
          return;
        }
        const date = new Date(Math.round((month - 25569) * 86400000));
        const year = date.getUTCFullYear();
        const quarter = `Q${Math.floor(date.getUTCMonth() / 3) + 1} ${year}`;
        let item = name + (detail ? ` (${detail})` : '');
        if (kind === 'construction') {
          const countParam = param.replace(/ const[.] cost$/, ' const. count').toLowerCase();
          const units = params.get(countParam)?.[monthColumns[index]];
          if (typeof units !== 'number' || !Number.isFinite(units)) {
            throw new Error(`Missing planned unit count for ${param} in month ${month}. Existing rows were preserved.`);
          }
          item = _allExpensesConstructionItem(param, units);
        }
        grouped.set(key, [type, year, quarter, month, item, value]);
      });
    }
  }
  return [...grouped.values()];
}

function _allExpensesConstructionItem(param, units) {
  const name = param.replace(/ const[.] cost$/, '');
  const singular = {'clinics': 'clinic', 'sala toilets': 'sala toilet', 'complex room': 'complex home'}[name] || name;
  const plural = {
    'lobby': 'lobbies', 'storage': 'storage', 'pool lounge seating': 'pool lounge seating',
    'general resort storage': 'general resort storage', 'shower & toilet': 'showers & toilets',
    'open air sala - big': 'open air salas - big', 'open air sala - medium': 'open air salas - medium',
  }[name] || singular + 's';
  const count = Math.sign(units) * Math.round(Math.abs(units));
  return `${count} ${units === 1 ? singular : plural}`;
}
