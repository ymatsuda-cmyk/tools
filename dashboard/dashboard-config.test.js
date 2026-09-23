const assert = require('assert');
const config = require('./dashboard-config.js');

const tabs = config.customTabs([
  { id: 'custom', label: 'カスタム' },
  { id: 'links', label: '予約済み' },
  { id: 'custom', label: '重複' },
  { id: 'invalid id', label: '不正' }
]);

assert.deepStrictEqual(tabs, [{ id: 'custom', label: 'カスタム' }]);
assert.strictEqual(config.cardTab({}, 'links', tabs), 'links');
assert.strictEqual(config.cardTab({ tab: 'custom' }, 'links', tabs), 'custom');
assert.strictEqual(config.cardTab({ tab: 'unknown' }, 'info', tabs), 'info');
assert.strictEqual(config.isVisible({ hidden: true }), false);
assert.strictEqual(config.isVisible({ hidden: false }), true);
assert.strictEqual(config.isVisible({}), true);

console.log('dashboard-config tests: OK');