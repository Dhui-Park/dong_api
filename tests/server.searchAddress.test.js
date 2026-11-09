const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require.resolve('../server.js'), 'utf8');
const context = { require, process, Buffer, console, __dirname: require('node:path').resolve(__dirname, '..') };
vm.createContext(context);
vm.runInContext(source.slice(0, source.lastIndexOf('app.listen(')), context);
test('normalizes attached parcel numbers', () => {
  assert.equal(context.preprocessAddress('서울 잠원동60-3'), '서울 잠원동 60-3');
  assert.equal(context.preprocessAddress('서울 잠원동 60-3'), '서울 잠원동 60-3');
});
test('rejects the observed wrong buildings', () => {
  assert.equal(context.matchesOriginalAddress('서울 서초구 잠원동60-3 메이플자이 205동', { jibunAddr: '서울특별시 서초구 서초동 1337-27 토요코인' }), false);
  for (const number of [3145, 3148]) assert.equal(context.matchesOriginalAddress(`충북 진천군 덕산면 두촌리 ${number}`, { jibunAddr: '충청북도 진천군 덕산읍 두촌리 2456' }), false);
});
test('accepts matching parcel and road identities with different details', () => {
  assert.equal(context.matchesOriginalAddress('강원 동해시 이도동 751 제101동 제401호', { jibunAddr: '강원특별자치도 동해시 이도동 751 동해아이파크' }), true);
  assert.equal(context.matchesOriginalAddress('서울 서초구 고무래로 94 301동', { roadAddr: '서울특별시 서초구 고무래로 94 (서초동)' }), true);
  assert.equal(context.matchesOriginalAddress('서울 서초구 고무래로 94', { roadAddr: '서울특별시 서초구 고무래로 95' }), false);
});
test('does not search below the parcel boundary', async () => {
  const keywords = [];
  context.searchAddressOnce = async keyword => { keywords.push(keyword); return { success: false }; };
  await context.runBackoffChain('서울 서초구 잠원동 60-3 메이플자이 205동 2004호', [], null, 1);
  assert.equal(keywords.at(-1), '서울 서초구 잠원동 60-3');
  assert.ok(keywords.every(k => k.includes('잠원동 60-3')));
});
