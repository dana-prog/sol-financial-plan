const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const test = require('node:test');
function load(globals = {}) {
  const context = {Date, Map, Set, ...globals};
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(require('node:path').join(__dirname, '../src/Wise.js'), 'utf8'), context);
  return context;
}
const balance = {id: 10, currency: 'USD', type: 'STANDARD', amount: {value: 400, currency: 'USD'}};
function transaction(overrides = {}) {
  return {date: '2026-11-01T12:00:00Z', type: 'DEBIT', referenceNumber: 'TRANSFER-1',
    amount: {value: -100, currency: 'USD'}, totalFees: {value: 2, currency: 'USD'},
    runningBalance: {value: 400, currency: 'USD'}, details: {description: '=malicious()', paymentReference: 'Invoice 1'}, ...overrides};
}
test('statement amounts retain debit and credit signs without subtracting fees twice', () => {
  const ctx = load();
  const debit = ctx.wiseTransactionRows_([transaction()], balance, '20')[0];
  assert.equal(debit[2], -100); assert.equal(debit[3], 2);
  assert.equal(ctx.wiseTransactionRows_([transaction({type: 'CREDIT'})], balance, '20')[0][2], 100);
  assert.equal(ctx.wiseCell_(debit[6]), "'=malicious()");
  assert.equal(ctx.wiseCell_(-100), -100);
});
test('malformed amounts, currencies, dates and account IDs fail before writing', () => {
  const ctx = load();
  assert.throws(() => ctx.wiseTransactionRows_([transaction({amount: {value: 3, currency: 'EUR'}})], balance, '20'), /currency/);
  assert.throws(() => ctx.wiseTransactionRows_([transaction({date: 'invalid'})], balance, '20'), /incomplete/);
  assert.throws(() => ctx.wiseBalanceRows_([{...balance, amount: {value: NaN, currency: 'USD'}}], '20', new Date()), /amount/);
  assert.throws(() => ctx.wiseId_('../profiles'), /ID/);
  assert.throws(() => ctx.wiseStart_('2026-02-30', new Date('2026-11-13')), /date/);
  assert.throws(() => ctx.wiseStart_('2027-01-01', new Date('2026-11-13')), /date/);
});
test('period replacement updates corrections, removes stale rows and preserves other balances and older history', () => {
  const ctx = load();
  const row = ctx.wiseTransactionRows_([transaction()], balance, '20')[0];
  const older = [...row]; older[0] = new Date('2026-01-01');
  const other = [...row]; other[10] = '11';
  const corrected = [...row]; corrected[2] = -120;
  const part = {balanceId:'10', profileId:'20', start:Date.parse('2026-11-01'), end:Date.parse('2026-11-03'), rows:[corrected]};
  const result = ctx.wiseMergeTransactions_([row,older,other], [part,part]);
  assert.equal(result.length, 3);
  assert.equal(result.filter(item => item[2] === -120).length, 1);
  assert.equal(ctx.wiseMergeTransactions_([row], [{...part,rows:[]}]).length, 0);
});
test('identical references on opposite sides of a conversion and refund remain separate', () => {
  const ctx = load();
  const debit = ctx.wiseTransactionRows_([transaction()], balance, '20')[0];
  const credit = ctx.wiseTransactionRows_([transaction({type:'CREDIT'})], balance, '20')[0];
  const otherBalance = [...debit]; otherBalance[10] = '11';
  assert.equal(ctx.wiseMergeTransactions_([], [{rows:[debit,credit,otherBalance]}]).length, 3);
});
test('API requests are read-only, use a header token, and never expose response bodies', () => {
  let request;
  const ctx = load({
    PropertiesService: {getScriptProperties: () => ({getProperty: () => 'secret-token'})},
    UrlFetchApp: {fetch: (url, options) => {
      request = {url, options};
      return {getResponseCode: () => 403, getContentText: () => 'secret-token and private details'};
    }},
  });
  assert.throws(() => ctx.wiseGet_('/profiles'), error => /HTTP 403/.test(error.message) && !error.message.includes('secret-token'));
  assert.equal(request.options.method, 'get');
  assert.equal(request.options.followRedirects, false);
  assert.equal(request.options.headers.Authorization, 'Bearer secret-token');
  assert.ok(!request.url.includes('secret-token'));
});
test('long rate limits are reported without an early retry', () => {
  let calls = 0;
  const ctx = load({
    PropertiesService: {getScriptProperties: () => ({getProperty: () => 'token'})},
    UrlFetchApp: {fetch: () => {calls++; return {getResponseCode: () => 429, getAllHeaders: () => ({'Retry-After': '60'})};}},
  });
  assert.throws(() => ctx.wiseGet_('/profiles'), /rate limit/);
  assert.equal(calls, 1);
});
test('failed statement fetch leaves outputs and checkpoint unchanged and releases lock', () => {
  let writes = 0, released = false;
  const props = {getProperty: () => null, setProperties: () => writes++};
  const ctx = load({
    PropertiesService: {getScriptProperties: () => props},
    LockService: {getScriptLock: () => ({tryLock: () => true, releaseLock: () => {released = true;}})},
    SpreadsheetApp: {openById: () => ({getSheetByName: () => null, insertSheet: () => writes++})},
  });
  ctx.wiseBusinessProfiles_ = () => [{id:'20', name:'Business'}];
  ctx.wiseGet_ = path => {
    if (path.includes('statement.json')) throw new Error('Wise HTTP 403');
    return path.endsWith('STANDARD') ? [balance] : [];
  };
  assert.throws(() => ctx.syncWiseAccount({profileId:'20', startDate:'2026-01-01'}), /403/);
  assert.equal(writes, 0); assert.equal(released, true);
});
test('history requests split into at most 365 days and save checkpoint after both writes', () => {
  const events = [], requests = [];
  const props = {getProperty: () => null, setProperties: values => events.push(['checkpoint',values]), deleteProperty() {}};
  const ctx = load({
    PropertiesService: {getScriptProperties: () => props},
    SpreadsheetApp: {openById: () => ({getSheetByName: () => null, getSpreadsheetTimeZone: () => 'Asia/Bangkok'}), flush: () => events.push(['flush'])},
    Utilities: {formatDate: () => 'now'},
  });
  ctx.wiseBusinessProfiles_ = () => [{id:'20', name:'Business'}];
  ctx.wiseGet_ = path => {
    if (!path.includes('statement.json')) return path.endsWith('STANDARD') ? [balance] : [];
    const query = new URL('https://example.test' + path).searchParams;
    const start = Date.parse(query.get('intervalStart')), end = Date.parse(query.get('intervalEnd'));
    requests.push({start,end});
    assert.ok(end - start <= 365 * 86400000);
    return {transactions:[]};
  };
  ctx.wiseWriteSheet_ = (_, __, name) => events.push(['write',name]);
  ctx.wiseSync_({profileId:'20', startDate:'2024-01-01'});
  assert.ok(requests.length > 1);
  for (let i = 1; i < requests.length; i++) assert.equal(requests[i].start, requests[i-1].end);
  assert.deepEqual(events.map(event => event[0]), ['write','write','flush','checkpoint']);
});
test('text formatting is applied before writing numeric-looking payment references', () => {
  const events = [];
  const range = {setValues: rows => events.push({write: rows}), setNumberFormat: format => events.push({format}),
    createFilter() {}, clearContent() {}, setWrap() {}};
  const sheet = {getLastRow: () => 1, getMaxRows: () => 1000, getMaxColumns: () => 26,
    getRange: (...args) => {events.push({range:args}); return range;}, getFilter: () => null, autoResizeRows() {}};
  const ctx = load();
  const row = ctx.wiseTransactionRows_([transaction({details:{paymentReference:'026073154018474'}})], balance, '20')[0];
  ctx.wiseWriteSheet_({}, sheet, 'Wise Transactions', Array(12).fill('header'), [row]);
  const writeIndex = events.findIndex(event => event.write);
  assert.ok(events.slice(0, writeIndex).some(event => event.range?.[1] === 8));
  assert.equal(events[writeIndex].write[1][7], '026073154018474');
  assert.equal(events.slice(0,writeIndex).filter(event => event.format === '@').length, 8);
});
test('every refresh reconciles the selected history for existing and new balances', () => {
  const saved = {WISE_PROFILE_ID:'20', WISE_START_DATE:'2026-01-01',
    WISE_LAST_SYNC:'2026-11-12T10:00:00Z', WISE_BALANCE_IDS:'["10"]'};
  const requests = [];
  const ctx = load({
    PropertiesService: {getScriptProperties: () => ({getProperty: key => saved[key], setProperties() {}, deleteProperty() {}})},
    SpreadsheetApp: {openById: () => ({getSpreadsheetTimeZone: () => 'Asia/Bangkok'}), flush() {}},
    Utilities: {formatDate: () => 'now'},
  });
  ctx.wiseBusinessProfiles_ = () => [{id:'20', name:'Business'}];
  ctx.wiseSheet_ = () => ({getLastRow: () => 1});
  ctx.wiseWriteSheet_ = () => {};
  ctx.wiseGet_ = path => {
    if (!path.includes('statement.json')) return path.endsWith('STANDARD') ? [balance,{...balance,id:11}] : [];
    const url = new URL('https://example.test' + path);
    requests.push([url.pathname, url.searchParams.get('intervalStart')]);
    return {transactions:[]};
  };
  ctx.wiseSync_({profileId:'20', startDate:'2026-01-01'});
  assert.equal(requests.find(([path]) => path.includes('/10/'))[1], '2026-01-01T00:00:00.000Z');
  assert.equal(requests.find(([path]) => path.includes('/11/'))[1], '2026-01-01T00:00:00.000Z');
});
test('daily schedule is idempotent and disabling it preserves unrelated triggers', () => {
  const triggers = [{getHandlerFunction: () => 'onInstallableEdit'}];
  let created = 0;
  const builder = {timeBased() {return this;}, everyDays() {return this;}, atHour() {return this;}, inTimezone() {return this;},
    create() {created++; triggers.push({getHandlerFunction: () => 'syncWiseScheduled'});}};
  const ctx = load({
    PropertiesService: {getScriptProperties: () => ({getProperty: () => 'last sync'})},
    ScriptApp: {getProjectTriggers: () => triggers, newTrigger: () => builder,
      deleteTrigger: trigger => triggers.splice(triggers.indexOf(trigger),1)},
  });
  ctx.enableWiseDailySync(); ctx.enableWiseDailySync();
  assert.equal(created,1);
  ctx.disableWiseDailySync();
  assert.equal(triggers.length,1); assert.equal(triggers[0].getHandlerFunction(),'onInstallableEdit');
});
