const WISE_API_BASE = 'https://api.wise.com/2026Q4';
const WISE_SPREADSHEET_ID = '1mxvjXTedpabDNPl-WOEie4vdpeIwCnXDlzvihZvUML4';
const WISE_TRANSACTION_SHEET_ID = 2089007372;
const WISE_DAY_MS = 86400000;
const WISE_TRANSACTION_HEADERS = ['Date', 'Currency', 'Amount', 'Fees', 'Fee currency', 'Direction', 'Description', 'Payment reference', 'Wise reference', 'Running balance', 'Balance ID', 'Profile ID'];

/** All Wise calls are GET requests; the saved credential never leaves the server. */
function wiseGet_(path) {
  const token = PropertiesService.getScriptProperties().getProperty('WISE_API_TOKEN');
  if (!token || !token.trim()) throw new Error('Save WISE_API_TOKEN in Apps Script Project Settings → Script Properties.');
  let response;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      response = UrlFetchApp.fetch(WISE_API_BASE + path, {
        method: 'get', headers: {Authorization: 'Bearer ' + token.trim(), Accept: 'application/json'},
        muteHttpExceptions: true, followRedirects: false,
      });
    } catch (_) {
      throw new Error('Could not reach Wise. Try syncing again.');
    }
    const status = response.getResponseCode();
    if (status === 200) {
      try { return JSON.parse(response.getContentText()); }
      catch (_) { throw new Error('Wise returned an invalid JSON response.'); }
    }
    if (status === 429 && attempt < 2) {
      const headers = response.getAllHeaders();
      const retryKey = Object.keys(headers).find(key => key.toLowerCase() === 'retry-after');
      const wait = retryKey ? Number(headers[retryKey]) : attempt + 1;
      if (!Number.isFinite(wait) || wait < 0 || wait > 10) break;
      Utilities.sleep(Math.max(1, wait) * 1000);
      continue;
    }
    break;
  }
  const status = response.getResponseCode();
  const reason = {
    401: 'The API token is invalid or expired. Update WISE_API_TOKEN in Script Properties.',
    403: 'Wise denied access. Check token permissions and whether Wise requires additional verification.',
    429: 'Wise rate limit reached. Try again later.',
  }[status] || 'Wise could not complete this read. Try again or check the account’s API access.';
  // Do not surface response bodies: upstream errors can contain credentials or account details.
  throw new Error('Wise HTTP ' + status + ': ' + reason);
}

function wiseId_(value) {
  const id = String(value);
  if (!/^\d+$/.test(id)) throw new Error('Wise returned an invalid account ID.');
  return id;
}

function wiseStart_(value, now) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error('WISE_START_DATE must be a valid YYYY-MM-DD date.');
  const date = new Date(value + 'T00:00:00.000Z');
  if (!Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== value || date >= now) {
    throw new Error('WISE_START_DATE must be a valid date before today.');
  }
  return date;
}

function wiseMoney_(money, currency, optional) {
  if (optional && money == null) return '';
  if (!money || typeof money.value !== 'number' || !Number.isFinite(money.value) || money.currency !== currency) {
    throw new Error('Wise returned a missing amount or mismatched currency.');
  }
  return money.value;
}

function wiseTransactionRows_(transactions, balance, profileId) {
  return transactions.map(transaction => {
    const date = new Date(transaction.date);
    if (!Number.isFinite(date.getTime()) || !['DEBIT', 'CREDIT'].includes(transaction.type) || !transaction.referenceNumber) {
      throw new Error('Wise returned an incomplete statement transaction.');
    }
    const fees = transaction.totalFees;
    const details = transaction.details || {};
    const amount = wiseMoney_(transaction.amount, balance.currency);
    return [date, balance.currency, transaction.type === 'DEBIT' ? -Math.abs(amount) : Math.abs(amount),
      fees == null ? '' : wiseMoney_(fees, fees.currency), fees ? fees.currency : '', transaction.type,
      details.description || '', details.paymentReference || '', String(transaction.referenceNumber),
      wiseMoney_(transaction.runningBalance, balance.currency, true), wiseId_(balance.id), profileId];
  });
}

function wiseTransactionKey_(row) {
  // References can be shared by refunds and the two sides of a currency conversion.
  return JSON.stringify([row[11], row[10], row[1], row[8], row[5]].map(String));
}

/** Existing transactions are user-maintained history, even if Wise later changes them. */
function wiseNewTransactions_(existing, imports) {
  const seen = new Set(existing.map(wiseTransactionKey_));
  const rows = [];
  for (const row of imports.flatMap(part => part.rows)) {
    const key = wiseTransactionKey_(row);
    if (seen.has(key)) continue;
    seen.add(key);
    rows.push(row);
  }
  return rows.sort((a, b) => new Date(b[0]).getTime() - new Date(a[0]).getTime() || String(a[8]).localeCompare(String(b[8])));
}

/** Shared by the daily trigger and the manual menu item. */
function syncWiseScheduled() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another Wise sync is running. Try again shortly.');
  try {
    const spreadsheet = SpreadsheetApp.openById(WISE_SPREADSHEET_ID);
    const count = wiseSync_(spreadsheet, new Date());
    spreadsheet.toast('Wise balance updated. ' + count + ' transactions in history.', 'Wise sync complete', 5);
    return count;
  } finally {
    lock.releaseLock();
  }
}

function wiseSync_(spreadsheet, now) {
  const properties = PropertiesService.getScriptProperties();
  const profileId = wiseId_(properties.getProperty('WISE_PROFILE_ID'));
  const historyStart = wiseStart_(properties.getProperty('WISE_START_DATE') || '2026-01-01', now);
  const overview = spreadsheet.getSheetById(190902611);
  if (!overview || overview.getRange('J2').getValue() !== 'Wise Balance') {
    throw new Error('The Wise Balance field could not be located in Project Overview.');
  }
  if (['J3', 'L3', 'J5'].some(cell => overview.getRange(cell).getFormula())) {
    throw new Error('The Wise balance or update-time field contains a formula. No formula was overwritten.');
  }
  const transactionSheet = wiseSheet_(spreadsheet);
  const headers = wiseHeaders_(transactionSheet);
  const previousRows = transactionSheet.getLastRow() > 1
    ? transactionSheet.getRange(2, 1, transactionSheet.getLastRow() - 1, headers.length).getValues()
      .map(row => WISE_TRANSACTION_HEADERS.map(header => row[headers.indexOf(header)])) : [];
  const balances = ['STANDARD', 'SAVINGS'].flatMap(type => {
    const result = wiseGet_('/profiles/' + profileId + '/balances?types=' + type);
    if (!Array.isArray(result)) throw new Error('Wise returned an invalid balance list.');
    return result;
  });
  const usd = balances.find(balance => balance.currency === 'USD' && balance.type === 'STANDARD');
  const available = wiseMoney_(usd && usd.amount, 'USD');
  const thb = wiseThbEquivalent_(available, wiseGet_('/rates?source=USD&target=THB'));
  const imports = wiseFetchTransactions_(balances, profileId, previousRows, historyStart, now);
  const transactions = wiseNewTransactions_(previousRows, imports);
  wiseWriteTransactions_(transactionSheet, headers, transactions);
  overview.getRange('L3').setValue(available);
  overview.getRange('J3').setValue(thb.amount).setNote('USD balance converted at 1 USD = ' + thb.rate
    + ' THB. Wise rate timestamp: ' + thb.time + '. Indicative equivalent before conversion fees.');
  overview.getRange('J5').setValue(now);
  SpreadsheetApp.flush();
  return previousRows.length + transactions.length;
}

function wiseThbEquivalent_(usd, rates) {
  const rate = Array.isArray(rates) && rates.length === 1 ? rates[0] : null;
  if (!rate || rate.source !== 'USD' || rate.target !== 'THB' || typeof rate.rate !== 'number'
    || !Number.isFinite(rate.rate) || rate.rate <= 0 || !Number.isFinite(Date.parse(rate.time))
    || !Number.isFinite(usd * rate.rate)) {
    throw new Error('Wise returned an invalid USD/THB exchange rate. Nothing was written.');
  }
  return {amount: usd * rate.rate, rate: rate.rate, time: rate.time};
}

function wiseImportStart_(rows, balanceId, profileId, historyStart, now) {
  let latest = historyStart.getTime();
  for (const row of rows) {
    if (String(row[10]) !== balanceId || String(row[11]) !== profileId) continue;
    const date = row[0];
    if (!(date instanceof Date) || !Number.isFinite(date.getTime()) || date > now) {
      throw new Error('Wise Transactions contains an invalid or future transaction date. Nothing was written.');
    }
    latest = Math.max(latest, date.getTime());
  }
  // Sheets can round stored dates by a millisecond. Include the boundary second and deduplicate it.
  return Math.floor(latest / 1000) * 1000;
}

function wiseFetchTransactions_(balances, profileId, previousRows, historyStart, now) {
  const imports = [];
  const started = Date.now();
  for (const balance of balances) {
    const balanceId = wiseId_(balance.id);
    const start = wiseImportStart_(previousRows, balanceId, profileId, historyStart, now);
    for (let cursor = start; cursor < now.getTime();) {
      // Keep below Wise's 469-day maximum, including leap years.
      const end = Math.min(cursor + 365 * WISE_DAY_MS, now.getTime());
      if (Date.now() - started > 240000) throw new Error('Wise import exceeded four minutes. Retry the sync. Nothing was written.');
      const query = {currency: balance.currency, intervalStart: new Date(cursor).toISOString(),
        intervalEnd: new Date(end).toISOString(), type: 'COMPACT'};
      const statement = wiseGet_('/profiles/' + profileId + '/balance-statements/' + balanceId + '/statement.json?'
        + Object.keys(query).map(key => key + '=' + encodeURIComponent(query[key])).join('&'));
      if (!statement || !Array.isArray(statement.transactions)) throw new Error('Wise returned an invalid statement.');
      const rows = wiseTransactionRows_(statement.transactions, balance, profileId)
        .filter(row => row[0].getTime() >= cursor && row[0].getTime() <= end);
      imports.push({balanceId, profileId, start: cursor, end, rows});
      cursor = end;
    }
  }
  return imports;
}

function wiseSheet_(spreadsheet) {
  // The tab may be renamed without changing its identity.
  const sheet = spreadsheet.getSheetById(WISE_TRANSACTION_SHEET_ID);
  if (!sheet) throw new Error('The Wise transaction tab could not be located. Nothing was written.');
  return sheet;
}

function wiseHeaders_(sheet) {
  const width = sheet.getLastColumn();
  const headers = width ? sheet.getRange(1, 1, 1, width).getValues()[0]
    .map(header => ['New Balance', 'Balance'].includes(header) ? 'Running balance' : header) : [];
  if (headers.some(header => typeof header !== 'string' || !header.trim())
    || new Set(headers).size !== headers.length
    || WISE_TRANSACTION_HEADERS.some(header => !headers.includes(header))) {
    throw new Error('The Wise transaction tab has missing or duplicate column headers. Nothing was written.');
  }
  return headers;
}

function wiseCell_(value) {
  // Bank descriptions and references are text, never spreadsheet formulas.
  return typeof value === 'string' && /^[=+\-@]/.test(value) ? "'" + value : value;
}

function wiseWriteTransactions_(sheet, headers, rows) {
  if (!rows.length) return;
  const firstRow = sheet.getLastRow() + 1;
  const rowCount = firstRow + rows.length - 1;
  if (sheet.getMaxRows() < rowCount) sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
  const columnRange = header => sheet.getRange(firstRow, headers.indexOf(header) + 1, rows.length, 1);
  // Preserve numeric-looking identifiers and leading zeroes before Sheets parses the values.
  ['Currency', 'Fee currency', 'Direction', 'Description', 'Payment reference', 'Wise reference', 'Balance ID', 'Profile ID']
    .forEach(header => columnRange(header).setNumberFormat('@'));
  columnRange('Date').setNumberFormat('yyyy-mm-dd hh:mm');
  ['Amount', 'Fees', 'Running balance'].forEach(header => columnRange(header).setNumberFormat('#,##0.00;[Red](#,##0.00)'));
  // Manual columns, including Counterparty and Item, start empty for every new transaction.
  const values = rows.map(row => headers.map(header => {
    const index = WISE_TRANSACTION_HEADERS.indexOf(header);
    return index < 0 ? '' : wiseCell_(row[index]);
  }));
  sheet.getRange(firstRow, 1, rows.length, headers.length).setValues(values);
}
