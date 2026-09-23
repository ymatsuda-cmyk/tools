const assert = require('assert');

(async () => {
  const api = await import('./customer9.api.js');

  // normalizeConfig: 書き漏らし（provider・url）を埋める
  assert.deepStrictEqual(api.normalizeConfig({}), { provider: 'json', url: '' });
  assert.deepStrictEqual(
    api.normalizeConfig({ provider: 'gas', url: 'https://script.google.com/macros/s/xxx/exec' }),
    { provider: 'gas', url: 'https://script.google.com/macros/s/xxx/exec' }
  );

  // json プロバイダ: 未設定でも組み込みサンプルが返る
  const demo = await api.list(api.normalizeConfig({}));
  assert.strictEqual(demo.provider, 'json');
  assert.ok(Array.isArray(demo.customers) && demo.customers.length > 0);
  assert.ok(demo.customers.every(c => typeof c.code === 'string' && typeof c.name === 'string'));

  // json プロバイダ: config.customers を渡すとそれを使う
  const custom = await api.list(api.normalizeConfig({ customers: [{ code: 'X1', name: 'テスト太郎' }] }));
  assert.deepStrictEqual(custom.customers, [
    { code: 'X1', name: 'テスト太郎', email: '', phone: '', status: '' }
  ]);

  // gas プロバイダ: URL未設定ならエラーになる
  await assert.rejects(
    () => api.list(api.normalizeConfig({ provider: 'gas' })),
    /URL/
  );

  console.log('customer9.api tests: OK');
})().catch(err => {
  console.error(err);
  process.exit(1);
});
