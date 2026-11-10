const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../server.js'), 'utf8');
const context = { require, process, Buffer, console, __dirname: require('node:path').resolve(__dirname, '..') };
vm.createContext(context);
vm.runInContext(source.slice(0, source.lastIndexOf('app.listen(')), context);
const realSearchAddressOnce = context.searchAddressOnce;
test('ignores latter-half building names as dong candidates', () => {
  for (const address of [
    '경기도 시흥시 서울대학로264번길 50 배곧헤리움 어반크로스2차 제에이동 제702호',
    '부산 북구 가람로 64 902호, 웅진아름누리',
    '서울 중구 청구로19길 9-7 리버파크뷰신당 나동 603호',
  ]) assert.deepEqual(Array.from(context.extractDongCandidates(address)), []);
});
test('uses token start and trimmed full-address midpoint', () => {
  assert.deepEqual(Array.from(context.extractDongCandidates('  당하동 1318-6  ')), ['당하동']);
  // Starts exactly at the midpoint, and is therefore excluded.
  assert.deepEqual(Array.from(context.extractDongCandidates('123 나동')), []);
  // Starts before the midpoint but extends beyond it, and is included.
  assert.deepEqual(Array.from(context.extractDongCandidates('1 가나동')), ['가나동']);
});
test('keeps original midpoint across gu removal and successive backoff queries', async () => {
  const saved = context.searchAddressOnce;
  try {
    context.searchAddressOnce = async keyword => ({
      success: keyword === '인천 1234567890', totalCount: 1, emdNm: '다른동',
    });
    const result = await context.searchAddress('인천 서구 1234567890 나동', {mode: 'precise'});
    assert.equal(result.success, true);
  } finally { context.searchAddressOnce = saved; }
});
test('retains attached parentheses as a single token', () => {
  assert.deepEqual(Array.from(context.tokenizeForBackoff('인천 당하동 1318-6(아파트 이름(별칭))')),
    ['인천', '당하동', '1318-6', '(아파트 이름(별칭))']);
});
test('keeps address preprocessing', () => {
  assert.equal(context.preprocessAddress('서울 잠원동60-3'), '서울 잠원동 60-3');
});
test('backoff reaches parcel without 번지 and accepts old-style API result', async () => {
  const saved = context.searchAddressOnce;
  const keywords = [];
  try {
    context.searchAddressOnce = async keyword => {
      keywords.push(keyword);
      return keyword === '서울 서초구 방배2동 946-8'
        ? { success: true, totalCount: 1, jibunAddr: '서울특별시 서초구 방배동 946-8' }
        : { success: false };
    };
    const result = await context.searchAddress('서울 서초구 방배2동 946-8번지 일대 디에이치방배 104동 1402호', {mode: 'precise'});
    assert.equal(result.success, true);
    assert.equal(keywords.at(-1), '서울 서초구 방배2동 946-8');
  } finally { context.searchAddressOnce = saved; }
});
test('retains old dong plausibility checks and unique-candidate guard', async () => {
  assert.equal(context.isPlausibleBackoffMatch(['당하동'], '인천', {emdNm: '다른동'}), false);
  assert.equal(context.isPlausibleBackoffMatch(['당하동'], '인천', {emdNm: '당하동'}), true);
  assert.equal(context.isPlausibleBackoffMatch(['당하동'], '인천 당하동', {emdNm: '다른동'}), true);
  const saved = context.searchAddressOnce;
  try {
    context.searchAddressOnce = async () => ({success: true, totalCount: 2, emdNm: '당하동'});
    const result = await context.runBackoffChain('인천 당하동 1318-6', ['당하동'], {}, 1, {requireUnique: true});
    assert.equal(result.success, false);
  } finally { context.searchAddressOnce = saved; }
});
test('does not generate the verification column', () => {
  const row = context.buildResultRow({}, {success: true});
  assert.equal(Object.hasOwn(row, '약간은 확인 필요(ex. 구 이름이 일치하지 않는다)'), false);
});
test('running status does not infer a proxy error', () => {
  assert.equal(context.inferLikelyCause({ status: 'running', elapsedMs: 90000 }), '주소 조회를 진행 중입니다.');
});
test('records API outcomes without logging request secrets', async () => {
  const axios = require('axios');
  const original = axios.get;
  const originalLog = console.log;
  const logs = [];
  const diagnostic = { enabled: true, verbose: true, apiStats: {callCount: 0, totalElapsedMs: 0, maxElapsedMs: 0, timeoutCount: 0, failureCount: 0} };
  try {
    console.log = (...args) => logs.push(args);
    for (const outcome of [
      {status: 200, data: {results: {common: {errorCode: '0'}, juso: [{}]}}},
      {status: 200, data: {results: {common: {errorCode: 'E0001'}}}},
      Object.assign(new Error('secret-address'), {code: 'ECONNABORTED'}),
      Object.assign(new Error('secret-address'), {response: {status: 503}})
    ]) {
      axios.get = async () => { if (outcome instanceof Error) throw outcome; return outcome; };
      await realSearchAddressOnce('secret-address', 1, diagnostic);
    }
    assert.equal(diagnostic.apiStats.callCount, 4);
    assert.equal(diagnostic.apiStats.timeoutCount, 1);
    assert.equal(diagnostic.apiStats.failureCount, 3);
    assert.equal(JSON.stringify(logs).includes('secret-address'), false);
    assert.equal(logs[3][1].httpStatus, 503);
  } finally {
    axios.get = original;
    console.log = originalLog;
  }
});
