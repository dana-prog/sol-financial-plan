const WISE_API_BASE = 'https://api.wise.com/2026Q4';
const WISE_SPREADSHEET_ID = '1mxvjXTedpabDNPl-WOEie4vdpeIwCnXDlzvihZvUML4';
const WISE_DAY_MS = 86400000;
const WISE_BALANCE_HEADERS = ['Currency', 'Available balance', 'Reserved amount', 'Balance type', 'Name', 'Investment state', 'Updated at', 'Profile ID', 'Balance ID'];
const WISE_TRANSACTION_HEADERS = ['Date', 'Currency', 'Amount', 'Fees', 'Fee currency', 'Direction', 'Description', 'Payment reference', 'Wise reference', 'Running balance', 'Balance ID', 'Profile ID'];

function showWiseConnection() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('WiseConnection').setWidth(520).setHeight(530), 'Wise connection');
}

function getWiseConnectionState() {
  const properties = PropertiesService.getScriptProperties();
  const profiles = wiseBusinessProfiles_();
  return {
    profiles,
    profileId: properties.getProperty('WISE_PROFILE_ID') || (profiles.length === 1 ? profiles[0].id : ''),
    startDate: properties.getProperty('WISE_START_DATE') || '2026-01-01',
    lastSync: properties.getProperty('WISE_LAST_SYNC') || '',
    automatic: ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === 'syncWiseScheduled'),
  };
}

function wiseBusinessProfiles_() {
  const profiles = wiseGet_('/profiles');
  if (!Array.isArray(profiles)) throw new Error('Wise returned an invalid profile list.');
  const business = profiles.filter(profile => String(profile.type).toUpperCase() === 'BUSINESS');
  if (!business.length) throw new Error('This token does not provide access to a Wise business profile.');
  return business.map(profile => ({
    id: wiseId_(profile.id),
    name: profile.businessName || (profile.details && profile.details.name) || 'Business ' + profile.id,
  }));
}

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
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error('Choose a valid history start date.');
  const date = new Date(value + 'T00:00:00.000Z');
  if (!Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== value || date >= now) {
    throw new Error('Choose a valid history start date before today.');
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

function wiseBalanceRows_(balances, profileId, now) {
  return balances.map(balance => {
    if (!/^[A-Z]{3}$/.test(balance.currency)) throw new Error('Wise returned an invalid currency.');
    return [balance.currency, wiseMoney_(balance.amount, balance.currency),
      wiseMoney_(balance.reservedAmount, balance.currency, true), balance.type, balance.name || '',
      balance.investmentState || '', now, profileId, wiseId_(balance.id)];
  }).sort((a, b) => a[0].localeCompare(b[0]) || a[8].localeCompare(b[8]));
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

function syncWiseAccount(options) {
  if (!options || typeof options !== 'object') throw new Error('Choose a business profile and history start date in SOL → Wise connection.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another Wise sync is running. Try again shortly.');
  try { return wiseSync_(options); }
  finally { lock.releaseLock(); }
}

function wiseSync_(options) {
  const now = new Date();
  const historyStart = wiseStart_(options.startDate, now);
  const profileId = wiseId_(options.profileId);
  const profiles = wiseBusinessProfiles_();
  const profile = profiles.find(item => item.id === profileId);
  if (!profile) throw new Error('Select one of the business profiles available to this token.');
  const properties = PropertiesService.getScriptProperties();
  const previousProfile = properties.getProperty('WISE_PROFILE_ID');
  if (previousProfile && previousProfile !== profileId) throw new Error('This financial plan is already connected to a different Wise business profile.');
  const spreadsheet = SpreadsheetApp.openById(WISE_SPREADSHEET_ID);
  const balanceSheet = wiseSheet_(spreadsheet, 'Wise Balances', WISE_BALANCE_HEADERS);
  const transactionSheet = wiseSheet_(spreadsheet, 'Wise Transactions', WISE_TRANSACTION_HEADERS);
  const previousRows = transactionSheet && transactionSheet.getLastRow() > 1
    ? transactionSheet.getRange(2, 1, transactionSheet.getLastRow() - 1, WISE_TRANSACTION_HEADERS.length).getValues() : [];
  const balances = ['STANDARD', 'SAVINGS'].flatMap(type => {
    const result = wiseGet_('/profiles/' + profileId + '/balances?types=' + type);
    if (!Array.isArray(result)) throw new Error('Wise returned an invalid balance list.');
    return result;
  });
  const balanceRows = wiseBalanceRows_(balances, profileId, now);
  const imports = [];
  const started = Date.now();
  for (const balance of balances) {
    const balanceId = wiseId_(balance.id);
    for (let cursor = historyStart.getTime(); cursor < now.getTime();) {
      // Keep below Wise's 469-day maximum, including leap years.
      const end = Math.min(cursor + 365 * WISE_DAY_MS, now.getTime());
      if (Date.now() - started > 240000) throw new Error('History import exceeded four minutes. Choose a more recent start date and retry. No sync checkpoint was advanced.');
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
  const transactions = wiseMergeTransactions_(previousRows, imports);
  wiseWriteSheet_(spreadsheet, balanceSheet, 'Wise Balances', WISE_BALANCE_HEADERS, balanceRows);
  wiseWriteSheet_(spreadsheet, transactionSheet, 'Wise Transactions', WISE_TRANSACTION_HEADERS, transactions);
  SpreadsheetApp.flush();
  // Advance only after both outputs have been written; retrying is safe after a partial write.
  properties.setProperties({WISE_PROFILE_ID: profileId, WISE_START_DATE: options.startDate,
    WISE_LAST_SYNC: now.toISOString()});
  properties.deleteProperty('WISE_BALANCE_IDS');
  return 'Connected to ' + profile.name.replace(/\.$/, '') + '. Updated ' + balances.length + ' balances and ' + transactions.length
    + ' statement transactions. Last sync: ' + Utilities.formatDate(now, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss') + '.';
}

/** Refuse to overwrite a pre-existing tab unless its schema belongs to this integration. */
function wiseSheet_(spreadsheet, name, headers) {
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

function wiseWriteSheet_(spreadsheet, existing, name, headers, rows) {
  const sheet = existing || spreadsheet.insertSheet(name);
  const rowCount = rows.length + 1;
  const oldLastRow = sheet.getLastRow();
  if (sheet.getMaxRows() < rowCount) sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  // Preserve numeric-looking identifiers and leading zeroes before Sheets parses the values.
  const textColumns = name === 'Wise Transactions' ? [2, 5, 6, 7, 8, 9, 11, 12] : [1, 4, 5, 6, 8, 9];
  textColumns.forEach(column => sheet.getRange(1, column, rowCount, 1).setNumberFormat('@'));
  sheet.getRange(1, 1, rowCount, headers.length).setValues([headers, ...rows].map(row => row.map(wiseCell_)));
  if (oldLastRow > rowCount) sheet.getRange(rowCount + 1, 1, oldLastRow - rowCount, headers.length).clearContent();
  if (!existing) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4').setWrap(true);
    sheet.setRowHeight(1, 42);
    sheet.setColumnWidths(1, headers.length, 140);
    if (name === 'Wise Transactions') {
      sheet.setColumnWidth(1, 165);
      sheet.setColumnWidth(7, 380);
      sheet.setColumnWidth(8, 220);
      sheet.getRange('A1').setNote('Dates are displayed in the spreadsheet time zone. History start dates use UTC.');
      sheet.getRange('C1').setNote('Signed statement movement: credits positive, debits negative. Fees are shown separately as reported by Wise; do not subtract them again from the statement amount.');
    } else {
      sheet.setColumnWidth(5, 180);
      sheet.setColumnWidth(7, 165);
      sheet.getRange('B1').setNote('Available balance reported by Wise, in the currency shown. Reserved funds are shown separately.');
    }
  }
  if (rows.length) {
    if (name === 'Wise Transactions') {
      sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
      sheet.getRange(2, 3, rows.length, 2).setNumberFormat('#,##0.00;[Red](#,##0.00)');
      sheet.getRange(2, 10, rows.length, 1).setNumberFormat('#,##0.00;[Red](#,##0.00)');
      sheet.getRange(2, 7, rows.length, 3).setWrap(true);
      sheet.autoResizeRows(2, rows.length);
    } else {
      sheet.getRange(2, 2, rows.length, 2).setNumberFormat('#,##0.00;[Red](#,##0.00)');
      sheet.getRange(2, 7, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    }
  }
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  if (rows.length) sheet.getRange(1, 1, rowCount, headers.length).createFilter();
}

function enableWiseDailySync() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('WISE_LAST_SYNC')) throw new Error('Complete a successful sync before enabling daily refresh.');
  if (!ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === 'syncWiseScheduled')) {
    ScriptApp.newTrigger('syncWiseScheduled').timeBased().everyDays(1).atHour(7).inTimezone('Asia/Bangkok').create();
  }
  return 'Daily refresh enabled for approximately 07:00–08:00 Bangkok time.';
}

function disableWiseDailySync() {
  ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === 'syncWiseScheduled')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  return 'Daily refresh disabled.';
}

function syncWiseScheduled() {
  const properties = PropertiesService.getScriptProperties();
  return syncWiseAccount({profileId: properties.getProperty('WISE_PROFILE_ID'), startDate: properties.getProperty('WISE_START_DATE')});
}
