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
  assert.throws(() => ctx.wiseMoney_({value: NaN, currency: 'USD'}, 'USD'), /amount/);
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
function syncFixture() {
  const events = [];
  const saved = {WISE_PROFILE_ID:'20', WISE_START_DATE:'2026-01-01'};
  const overview = {getRange: cell => ({
    getValue: () => cell === 'J2' ? 'Wise Balance' : '', getFormula: () => '',
    setValue: value => {events.push([cell,value]); return {setNote() {}};},
  })};
  const spreadsheet = {getSheetById: () => overview, getSheetByName: () => null, toast() {}};
  const ctx = load({
    PropertiesService: {getScriptProperties: () => ({getProperty: key => saved[key]})},
    SpreadsheetApp: {openById: () => spreadsheet, flush: () => events.push(['flush'])},
    LockService: {getScriptLock: () => ({tryLock: () => true, releaseLock: () => events.push(['release'])})},
  });
  ctx.wiseWriteTransactions_ = (_, __, rows) => events.push(['transactions',rows]);
  ctx.wiseGet_ = path => path.startsWith('/rates?') ? [{source:'USD',target:'THB',rate:35,time:'2026-11-13T10:00:00Z'}]
    : path.includes('statement.json') ? {transactions:[]} : path.endsWith('STANDARD') ? [balance] : [];
  return {ctx,events,overview,spreadsheet};
}
test('failed statement fetch leaves balance, timestamp and transaction history unchanged', () => {
  const {ctx,events} = syncFixture();
  ctx.wiseGet_ = path => {
    if (path.includes('statement.json')) throw new Error('Wise HTTP 403');
    if (path.startsWith('/rates?')) return [{source:'USD',target:'THB',rate:35,time:'2026-11-13T10:00:00Z'}];
    return path.endsWith('STANDARD') ? [balance] : [];
  };
  assert.throws(() => ctx.syncWiseScheduled(), /403/);
  assert.deepEqual(events,[['release']]);
});
test('successful daily/manual sync updates the overview only after importing transactions', () => {
  const {ctx,events,spreadsheet} = syncFixture();
  const now = new Date('2026-11-13T10:00:00Z');
  ctx.wiseSync_(spreadsheet,now);
  assert.deepEqual(events.map(event => event[0]), ['transactions','N3','J3','J5','flush']);
  assert.equal(events[1][1],400);
  assert.equal(events[2][1],14000);
  assert.equal(events[3][1],now);
});
test('overview formulas are preserved and failed transaction writes do not advance the update time', () => {
  const {ctx,events,overview,spreadsheet} = syncFixture();
  ctx.wiseWriteTransactions_ = () => {throw new Error('write failed');};
  assert.throws(() => ctx.wiseSync_(spreadsheet,new Date('2026-11-13')), /write failed/);
  assert.equal(events.length,0);
  for (const protectedCell of ['J3', 'N3', 'J5']) {
    overview.getRange = cell => ({getValue: () => 'Wise Balance', getFormula: () => cell === protectedCell ? '=NAMED_BALANCE()' : ''});
    assert.throws(() => ctx.wiseSync_(spreadsheet,new Date('2026-11-13')), /No formula was overwritten/);
  }
  assert.equal(events.length,0);
});
test('each balance resumes from its latest imported transaction, independently of row order', () => {
  const ctx = load();
  const row = ctx.wiseTransactionRows_([transaction()], balance, '20')[0];
  const latest = [...row]; latest[0] = new Date('2026-11-05T10:23:45.678Z');
  const other = [...row]; other[10] = '11'; other[0] = new Date('2026-11-10T10:00:00Z');
  const otherProfile = [...latest]; otherProfile[11] = '21'; otherProfile[0] = new Date('2026-11-12');
  const start = new Date('2026-01-01'), now = new Date('2026-11-13');
  assert.equal(ctx.wiseImportStart_([latest,row,other,otherProfile],'10','20',start,now), Date.parse('2026-11-05T10:23:45Z'));
  assert.equal(ctx.wiseImportStart_([latest,row,other],'11','20',start,now), other[0].getTime());
  assert.equal(ctx.wiseImportStart_([latest,row,other],'12','20',start,now), start.getTime());
  const invalid = [...row]; invalid[0] = new Date('2027-01-01');
  assert.throws(() => ctx.wiseImportStart_([invalid],'10','20',start,now), /invalid or future/);
});
test('API history requests start at the latest transaction and split only long gaps', () => {
  const ctx = load(), requests = [];
  const row = ctx.wiseTransactionRows_([transaction()], balance, '20')[0];
  ctx.wiseGet_ = path => {
    const query = new URL('https://example.test' + path).searchParams;
    requests.push({start:Date.parse(query.get('intervalStart')), end:Date.parse(query.get('intervalEnd'))});
    return {transactions:[]};
  };
  const now = new Date('2026-11-13T10:00:00Z');
  ctx.wiseFetchTransactions_([balance],'20',[row],new Date('2024-01-01'),now);
  assert.equal(requests.length,1);
  assert.equal(requests[0].start,row[0].getTime());
  assert.equal(requests[0].end,now.getTime());
  requests.length = 0;
  ctx.wiseFetchTransactions_([balance],'20',[],new Date('2024-01-01'),now);
  assert.ok(requests.length > 1);
  requests.forEach(request => assert.ok(request.end - request.start <= 365 * 86400000));
  for (let i = 1; i < requests.length; i++) assert.equal(requests[i].start,requests[i-1].end);
});
test('reimporting the boundary replaces it without duplicating history and keeps same-time transactions', () => {
  const ctx = load();
  const boundary = ctx.wiseTransactionRows_([transaction()], balance, '20')[0];
  const sameTime = ctx.wiseTransactionRows_([transaction({referenceNumber:'TRANSFER-2'})], balance, '20')[0];
  const older = [...boundary]; older[0] = new Date('2026-10-01');
  const part = {balanceId:'10', profileId:'20', start:boundary[0].getTime(),end:Date.parse('2026-11-13'),rows:[boundary,sameTime]};
  const result = ctx.wiseMergeTransactions_([older,boundary],[part]);
  assert.equal(result.length,3);
  assert.equal(ctx.wiseMergeTransactions_(result,[part]).length,3);
});
test('text formatting is applied before writing numeric-looking payment references', () => {
  const events = [];
  const range = {setValues: rows => events.push({write: rows}), setNumberFormat: format => events.push({format}),
    createFilter() {}, clearContent() {}, setWrap() {}};
  const sheet = {getLastRow: () => 1, getMaxRows: () => 1000, getMaxColumns: () => 26,
    getRange: (...args) => {events.push({range:args}); return range;}, getFilter: () => null, autoResizeRows() {}};
  const ctx = load();
  const row = ctx.wiseTransactionRows_([transaction({details:{paymentReference:'026073154018474'}})], balance, '20')[0];
  ctx.wiseWriteTransactions_({}, sheet, [row]);
  const writeIndex = events.findIndex(event => event.write);
  assert.equal(events[writeIndex].write[1][7], '026073154018474');
  assert.equal(events.slice(0,writeIndex).filter(event => event.format === '@').length, 8);
});
test('THB equivalent uses the requested rate direction and rejects invalid rate responses', () => {
  const ctx = load();
  const rate = {source:'USD',target:'THB',rate:33.5,time:'2026-11-13T10:00:00Z'};
  assert.equal(ctx.wiseThbEquivalent_(100,[rate]).amount,3350);
  assert.equal(ctx.wiseThbEquivalent_(0,[rate]).amount,0);
  for (const invalid of [[],[{...rate,source:'THB',target:'USD'}],[{...rate,rate:0}],[{...rate,rate:-1}],
    [{...rate,rate:'33.5'}],[{...rate,time:'invalid'}]]) {
    assert.throws(() => ctx.wiseThbEquivalent_(100,invalid), /invalid USD\/THB/);
  }
});
test('failed exchange-rate retrieval leaves both balances and the update time untouched', () => {
  const {ctx,events,spreadsheet} = syncFixture();
  ctx.wiseGet_ = path => path.startsWith('/rates?') ? [] : path.endsWith('STANDARD') ? [balance] : [];
  assert.throws(() => ctx.wiseSync_(spreadsheet,new Date('2026-11-13')), /invalid USD\/THB/);
  assert.equal(events.length,0);
});
