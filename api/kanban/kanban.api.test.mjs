// node --test api/kanban/kanban.api.test.mjs
import test from 'node:test'
import assert from 'node:assert/strict'
import * as kanban from './kanban.api.js'

const serial = d => Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000) + 25569
const day = (base, n) => { const d = new Date(base); d.setDate(d.getDate() + n); return d }

test('暗号化して復号すると元に戻る。パスフレーズ違いは badpass', async () => {
  const src = { schema: kanban.SCHEMA, updatedAt: '2026-09-29T00:00:00Z', tasks: [{ id: 1, title: '秘密のタスク' }] }
  const env = await kanban.encrypt(src, '合言葉', 1000)
  assert.ok(kanban.isEnvelope(env))
  assert.ok(!JSON.stringify(env).includes('秘密'))
  assert.deepEqual(await kanban.decrypt(env, '合言葉'), src)
  await assert.rejects(kanban.decrypt(env, 'ちがう'), e => e.code === 'badpass')
})

test('Excelのシリアル値は時差に関係なくその日の0:00になる', () => {
  const d = kanban.toDate(46294)          // 2026-09-29
  assert.equal(`${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`, '2026-9-29')
  assert.equal(d.getHours(), 0)
})

test('遅延・本日〆・今週・来週・TODOの振り分けと総数', async () => {
  const now = new Date(2026, 8, 29)       // 火曜
  const t = (id, s, e, extra = {}) => ({ id, title: 't' + id, start: serial(day(now, s)), end: serial(day(now, e)), ...extra })
  const raw = {
    schema: kanban.SCHEMA,
    tasks: [
      t(1, -5, -2),                                         // 遅延
      t(2, -1, 0, { actualStart: serial(day(now, -1)) }),   // 本日〆・対応中
      t(3, -1, 0, { actualEnd: serial(now) }),              // 本日〆・完了
      t(4, 1, 3, { note: '☆▲' }),                           // 今週・保留
      t(5, 7, 9),                                           // 来週
      { id: 6, title: 't6' },                               // TODO（開始日・終了日とも未設定）
      { id: 7, title: 't7', actualEnd: serial(now) }        // 日付は無いが完了済みなのでTODOには出ない
    ]
  }
  globalThis.fetch = async () => ({ ok: true, status: 200, json: async () => raw })
  const data = await kanban.load({ binId: 'x' })
  const r = kanban.classify(data.tasks, { now })
  assert.equal(r.late.rest.length, 1)
  assert.deepEqual([r.today.rest.length, r.today.total], [1, 2])
  assert.deepEqual([r.week.rest.length, r.week.total], [2, 3])   // 9/28〜10/4 に重なるのは 2・3・4
  assert.equal(r.week.rest.find(x => x.id === '4').status, 'held')
  assert.deepEqual([r.next.rest.length, r.next.total], [1, 1])
  assert.deepEqual(r.todo.rest.map(x => x.id), ['6'])   // 開始日・終了日がどちらも未設定で未完了なのは6だけ
})


test('options.category / options.classification で大分類・小分類を絞り込める', async () => {
  const now = new Date(2026, 8, 29)       // 火曜
  const t = (id, s, e, extra = {}) => ({ id, title: 't' + id, start: serial(day(now, s)), end: serial(day(now, e)), ...extra })
  const raw = {
    schema: kanban.SCHEMA,
    tasks: [
      t(1, -1, 0, { category: '受注', classification: 'A社' }),
      t(2, -1, 0, { category: '受注', classification: 'B社' }),
      t(3, -1, 0, { category: '保守', classification: 'A社' })
    ]
  }
  globalThis.fetch = async () => ({ ok: true, status: 200, json: async () => raw })
  const data = await kanban.load({ binId: 'x' })
  assert.equal(kanban.classify(data.tasks, { now, category: '受注' }).today.total, 2)
  assert.equal(kanban.classify(data.tasks, { now, classification: 'A社' }).today.total, 2)
  assert.equal(kanban.classify(data.tasks, { now, category: '受注', classification: 'A社' }).today.total, 1)
})

test('暗号文をパスフレーズなしで読むと locked', async () => {
  const env = await kanban.encrypt({ schema: kanban.SCHEMA, tasks: [] }, 'p', 1000)
  globalThis.fetch = async () => ({ ok: true, status: 200, json: async () => env })
  await assert.rejects(kanban.load({ binId: 'x' }), e => e.code === 'locked')
  assert.deepEqual((await kanban.load({ binId: 'x', passphrase: 'p' })).tasks, [])
})

test('binId も url も省略すると data/kanban/wbs-tasks.enc.json の既定URLを使う', async () => {
  const raw = { schema: kanban.SCHEMA, tasks: [{ id: 1, title: 't1' }] }
  let calledUrl = ''
  globalThis.fetch = async (url) => { calledUrl = url; return { ok: true, status: 200, json: async () => raw } }
  const data = await kanban.load({})
  assert.match(calledUrl, /data\/kanban\/wbs-tasks\.enc\.json/)
  assert.equal(data.tasks[0].title, 't1')
})

test('config.url を指定すると、そのURLからそのまま読む', async () => {
  const raw = { schema: kanban.SCHEMA, tasks: [] }
  let calledUrl = ''
  globalThis.fetch = async (url) => { calledUrl = url; return { ok: true, status: 200, json: async () => raw } }
  await kanban.load({ url: 'https://example.test/mirror.json' })
  assert.match(calledUrl, /^https:\/\/example\.test\/mirror\.json\?t=\d+$/)
})

