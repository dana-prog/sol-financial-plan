const WISE_API_BASE = 'https://api.wise.com/2026Q4';
const WISE_SPREADSHEET_ID = '1mxvjXTedpabDNPl-WOEie4vdpeIwCnXDlzvihZvUML4';
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

/** Replace fetched periods to pick up corrections and removals, while preserving older history. */
function wiseMergeTransactions_(existing, imports) {
  const retained = existing.filter(row => !imports.some(part =>
    String(row[10]) === part.balanceId && String(row[11]) === part.profileId &&
    new Date(row[0]).getTime() >= part.start && new Date(row[0]).getTime() <= part.end));
  const rows = retained.concat(imports.flatMap(part => part.rows));
  // Adjacent statement requests may both include the transaction at their shared boundary.
  const unique = new Map(rows.map(row => [JSON.stringify(row), row]));
  return [...unique.values()].sort((a, b) => new Date(b[0]).getTime() - new Date(a[0]).getTime() || String(a[8]).localeCompare(String(b[8])));
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
  if (['J3', 'M3', 'J5'].some(cell => overview.getRange(cell).getFormula())) {
    throw new Error('The Wise balance or update-time field contains a formula. No formula was overwritten.');
  }
  const transactionSheet = wiseSheet_(spreadsheet);
  const previousRows = transactionSheet && transactionSheet.getLastRow() > 1
    ? transactionSheet.getRange(2, 1, transactionSheet.getLastRow() - 1, WISE_TRANSACTION_HEADERS.length).getValues() : [];
  const balances = ['STANDARD', 'SAVINGS'].flatMap(type => {
    const result = wiseGet_('/profiles/' + profileId + '/balances?types=' + type);
    if (!Array.isArray(result)) throw new Error('Wise returned an invalid balance list.');
    return result;
  });
  const usd = balances.find(balance => balance.currency === 'USD' && balance.type === 'STANDARD');
  const available = wiseMoney_(usd && usd.amount, 'USD');
  const thb = wiseThbEquivalent_(available, wiseGet_('/rates?source=USD&target=THB'));
  const imports = wiseFetchTransactions_(balances, profileId, previousRows, historyStart, now);
  const transactions = wiseMergeTransactions_(previousRows, imports);
  wiseWriteTransactions_(spreadsheet, transactionSheet, transactions);
  overview.getRange('M3').setValue(available);
  overview.getRange('J3').setValue(thb.amount).setNote('USD balance converted at 1 USD = ' + thb.rate
    + ' THB. Wise rate timestamp: ' + thb.time + '. Indicative equivalent before conversion fees.');
  overview.getRange('J5').setValue(now);
  SpreadsheetApp.flush();
  return transactions.length;
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
  // Sheets can round stored dates by a millisecond. Include the boundary second and reconcile it.
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

/** Refuse to overwrite a pre-existing tab unless its schema belongs to this integration. */
function wiseSheet_(spreadsheet) {
  const name = 'Wise Transactions';
  const headers = WISE_TRANSACTION_HEADERS;
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) return null;
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (JSON.stringify(existing) !== JSON.stringify(headers) || sheet.getLastColumn() > headers.length) {
    throw new Error('The ' + name + ' tab has an unexpected layout. Nothing was overwritten.');
  }
  return sheet;
}

function wiseCell_(value) {
  // Bank descriptions and references are text, never spreadsheet formulas.
  return typeof value === 'string' && /^[=+\-@]/.test(value) ? "'" + value : value;
}

function wiseWriteTransactions_(spreadsheet, existing, rows) {
  const headers = WISE_TRANSACTION_HEADERS;
  const sheet = existing || spreadsheet.insertSheet('Wise Transactions');
  const rowCount = rows.length + 1;
  const oldLastRow = sheet.getLastRow();
  if (sheet.getMaxRows() < rowCount) sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  // Preserve numeric-looking identifiers and leading zeroes before Sheets parses the values.
  [2, 5, 6, 7, 8, 9, 11, 12].forEach(column => sheet.getRange(1, column, rowCount, 1).setNumberFormat('@'));
  sheet.getRange(1, 1, rowCount, headers.length).setValues([headers, ...rows].map(row => row.map(wiseCell_)));
  if (oldLastRow > rowCount) sheet.getRange(rowCount + 1, 1, oldLastRow - rowCount, headers.length).clearContent();
  if (!existing) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4').setWrap(true);
    sheet.setRowHeight(1, 42);
    sheet.setColumnWidths(1, headers.length, 140);
    sheet.setColumnWidth(1, 165);
    sheet.setColumnWidth(7, 380);
    sheet.setColumnWidth(8, 220);
    sheet.getRange('A1').setNote('Dates are displayed in the spreadsheet time zone. History start dates use UTC.');
    sheet.getRange('C1').setNote('Signed statement movement: credits positive, debits negative. Fees are shown separately as reported by Wise; do not subtract them again from the statement amount.');
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange(2, 3, rows.length, 2).setNumberFormat('#,##0.00;[Red](#,##0.00)');
    sheet.getRange(2, 10, rows.length, 1).setNumberFormat('#,##0.00;[Red](#,##0.00)');
    sheet.getRange(2, 7, rows.length, 3).setWrap(true);
    sheet.autoResizeRows(2, rows.length);
  }
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  if (rows.length) sheet.getRange(1, 1, rowCount, headers.length).createFilter();
}
