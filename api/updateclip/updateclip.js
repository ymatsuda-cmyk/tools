/**
 * clipstock の「完了ぶんをまとめて生成」を、clipstock の画面の外から呼べる形にしたもの。
 *
 * 生成そのものとNotionへの書き戻しは clipstock 側の実装をそのまま呼ぶ。ここが持つのは
 * 「どのスペースの・どのAIで・どれを対象にするか」を引数で受け取る入口だけ。
 * 同じオリジンに置いてあるので、GASのURL・共有トークン・AI接続先は clipstock と同じ
 * localStorage を共有する。設定を二重に持たせない。
 *
 * 使い方:
 *   const api = await import('/api/updateclip/updateclip.js')
 *   const spaces = await api.spaces()
 *   const models = api.models()
 *   const plan = await api.targets(spaces[0].id)
 *   await api.run({ spaceId, modelKey, onProgress, shouldCancel })
 */

import { listVideos, listVideosFromNotion } from '../../beta/clipstock/src/lib/store.js'
import { fetchTranscript, saveGenerated, statusCounts } from '../../beta/clipstock/src/lib/gas.js'
import { generateAll, stagesOf } from '../../beta/clipstock/src/lib/generate.js'
import { initSpaces, selectSpace, spaceList, spaceMode, sourcesOf, defaultSpaceId } from '../../beta/clipstock/src/lib/spaces.js'
import { loadSettings, connectionOf } from '../../beta/clipstock/src/lib/llm-settings.js'
import { loadConfig, isConfigured } from '../../beta/clipstock/src/lib/videos-config.js'
import { excludeExcluded, STATUS_DONE, STATUS_SUMMARIZED } from '../../beta/clipstock/src/lib/filters.js'
import { knownTagsOf } from '../../beta/clipstock/src/lib/tags.js'

/** モデルは接続をまたいで同名がありうるので、接続IDと組にして一意にする */
const MODEL_SEPARATOR = '::'

let spacesReady = null

function ensureSpaces() {
  if (!spacesReady) spacesReady = initSpaces()
  return spacesReady
}

/** 選べるスペース(パターン)の一覧。spaces.json の中身がそのまま出る */
export async function spaces() {
  await ensureSpaces()
  return spaceList().map((s) => ({ id: s.id, label: s.label, mode: s.mode || '' }))
}

/** 選べるAIの一覧。clipstock の設定に登録されている接続×モデルを平らに並べる */
export function models() {
  const settings = loadSettings()
  return settings.connections.flatMap((c) =>
    (c.models || []).map((model) => ({
      key: c.id + MODEL_SEPARATOR + model,
      model,
      connectionLabel: c.label,
      label: c.label ? `${c.label} / ${model}` : model,
      active: c.id === settings.activeConnectionId && model === settings.activeModel,
    }))
  )
}

/** key に対応する接続。見つからなければ clipstock で選択中のものに落とす */
function connectionFor(key) {
  const settings = loadSettings()
  const fallback = connectionOf(settings)
  if (!key) return fallback

  const at = String(key).indexOf(MODEL_SEPARATOR)
  if (at === -1) return fallback
  const connectionId = String(key).slice(0, at)
  const model = String(key).slice(at + MODEL_SEPARATOR.length)

  const c = settings.connections.find((x) => x.id === connectionId)
  if (!c || !model) return fallback
  return {
    baseUrl: c.baseUrl,
    apiKey: c.apiKey,
    model,
    numCtx: c.numCtx,
    temperature: settings.temperature,
    think: c.think ?? null,
  }
}

/**
 * その段までひと通り揃っているか。
 * サマリだけは一覧に本文が乗っているので、has ではなく summary を見る。
 */
function isComplete(item, mode) {
  return stagesOf(mode).every((stage) =>
    stage.id === 'summary' ? Boolean(item.summary) : Boolean(item.has?.[stage.id])
  )
}

/**
 * 一覧の取得。件数は Notion を正とする。
 * 静的JSONは cron の書き出し待ちで遅れるので、Notion に届かないときだけそちらに落とす。
 */
async function loadItems(fromNotion) {
  if (fromNotion) {
    try {
      return await listVideosFromNotion()
    } catch (err) {
      console.warn('Notion から読めないため一覧JSONを使います:', err)
    }
  }
  return listVideos()
}

/**
 * 生成の対象。
 * 未生成(完了)だけでなく、途中で失敗して一部だけ欠けているものも拾う。
 * 拾わないと、1段だけ落ちた動画が永久に取り残される。
 */
export async function targets(spaceId, { fromNotion = true } = {}) {
  await ensureSpaces()
  const space = selectSpace(spaceId)
  const mode = spaceMode(space)

  const { items, source } = await loadItems(fromNotion)
  const visible = excludeExcluded(items)

  const fresh = visible.filter((i) => i.status === STATUS_DONE)
  const partial = visible.filter((i) => i.status === STATUS_SUMMARIZED && !isComplete(i, mode))

  // 状態ごとの件数。ダッシュボードが「完了が何件たまっているか」を出すのに使う
  const counts = {}
  visible.forEach((i) => {
    const key = i.status || ''
    counts[key] = (counts[key] || 0) + 1
  })

  return {
    spaceId: space.id,
    spaceLabel: space.label,
    mode,
    source,
    items: visible,
    total: visible.length,
    counts,
    fresh,
    partial,
  }
}

/**
 * そのパターンが読む集計のかたまりを選ぶ。
 * 分類で絞る取り込み元はその分類ぶんだけ、分類なしの取り込み元は受け皿なので、
 * 同じパターンの他の取り込み元が取る分類を除いた残り全部を見る(store.js の割り当てと同じ規則)。
 */
function bucketsOf(space, counts) {
  const sources = sourcesOf(space)
  const claimed = new Map()
  sources.forEach((s) => {
    if (!s.category) return
    const db = s.db || 'video'
    if (!claimed.has(db)) claimed.set(db, new Set())
    claimed.get(db).add(s.category)
  })

  return sources.flatMap((s) => {
    const byCategory = counts[s.db || 'video'] || {}
    if (s.category) return byCategory[s.category] ? [byCategory[s.category]] : []
    const taken = claimed.get(s.db || 'video') || new Set()
    return Object.entries(byCategory)
      .filter(([category]) => !taken.has(category))
      .map(([, bucket]) => bucket)
  })
}

/** パターン1つぶんの件数 */
function summarize(space, counts) {
  const mode = spaceMode(space)
  const byStatus = {}
  let total = 0
  let partial = 0

  bucketsOf(space, counts).forEach((bucket) => {
    total += bucket.total || 0
    partial += (bucket.partial && bucket.partial[mode]) || 0
    Object.entries(bucket.status || {}).forEach(([status, n]) => {
      byStatus[status] = (byStatus[status] || 0) + n
    })
  })

  return {
    id: space.id,
    label: space.label,
    mode,
    total,
    counts: byStatus,
    fresh: byStatus[STATUS_DONE] || 0,
    partial,
  }
}

/**
 * 件数だけを取る軽い入口。targets() と違って一覧そのものは取らない。
 * すべてのパターンぶんを1回でまとめて返すので、パターンを切り替えても取り直しは要らない。
 */
export async function countsAll({ refresh = false } = {}) {
  await ensureSpaces()
  const data = await statusCounts(refresh)
  const counts = data.counts || {}
  return {
    updatedAt: data.updatedAt || null,
    defaultSpaceId: defaultSpaceId(),
    spaces: spaceList().map((space) => summarize(space, counts)),
  }
}

/**
 * 対象をまとめて生成する。1件ごと・1段ごとにNotionへ保存するので、
 * 途中で止めても・失敗しても、そこまでの結果は残る。
 *
 * @param {object}   options
 * @param {string}   options.spaceId         対象のスペース(パターン)
 * @param {string}   [options.modelKey]      models() の key。省略時は clipstock で選択中のAI
 * @param {boolean}  [options.includePartial] 項目が欠けているものも対象にする(既定 true)
 * @param {(p: {index:number,total:number,title:string,stage:string,text:string}) => void} [options.onProgress]
 * @param {() => boolean} [options.shouldCancel] true を返した時点で次の1件に進まない
 */
export async function run({ spaceId, modelKey, includePartial = true, onProgress, shouldCancel } = {}) {
  if (!isConfigured(loadConfig())) {
    throw new Error('clipstock の接続設定が未入力です。clipstock の設定でGASのURLと共有トークンを入れてください')
  }

  const connection = connectionFor(modelKey)
  if (!connection) {
    throw new Error('AI接続が未設定です。clipstock の設定で接続先とモデルを追加してください')
  }

  const plan = await targets(spaceId)
  const list = includePartial ? [...plan.fresh, ...plan.partial] : plan.fresh
  const total = list.length

  const result = { total, done: 0, failed: 0, failures: [], cancelled: false, stopped: null, model: connection.model }
  if (!total) return result

  // 同じ回の中で似た動画がそれぞれ別の新語を作らないよう、生まれたタグも語彙に足していく
  let vocabulary = knownTagsOf(plan.items)

  for (let index = 0; index < total; index++) {
    if (shouldCancel?.()) {
      result.cancelled = true
      break
    }
    const item = list[index]
    const report = (stage, text) => onProgress?.({ index, total, title: item.title, stage, text: text || '' })

    try {
      report('原文を読み込み中', '')
      const { text: transcript } = await fetchTranscript(item.key)
      if (!transcript) throw new Error('原文が空です')

      await generateAll(
        {
          title: item.title,
          transcript,
          summary: '',
          fields: '',
          mode: plan.mode,
          knownTags: vocabulary,
          connection,
        },
        {
          onStageStart: (stage) => report(stage.label, ''),
          onProgress: (stage, text) => report(stage.label, text),
          onStage: async (stageId, stageDetail, model) => {
            await saveGenerated(item.key, stageDetail, model, transcript.length)
            if (Array.isArray(stageDetail.tags)) {
              stageDetail.tags.forEach((t) => {
                if (!vocabulary.includes(t)) vocabulary = [...vocabulary, t]
              })
            }
          },
        }
      )
      result.done++
    } catch (err) {
      const message = String(err?.message ?? err)
      result.failed++
      result.failures.push(`${item.title}: ${message}`)
      // 1日の利用上限に達したら、残りを叩いても全滅するだけなのでここで打ち切る
      if (message.includes('1日の利用上限')) {
        result.stopped = 'quota'
        break
      }
    }
  }

  return result
}
