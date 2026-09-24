/**
 * Issue状態ダッシュボード 一覧取得API
 *
 * utility/agent.html が読んでいるのと同じ data/agent/dashboard.json を、
 * ダッシュボード（稼働状況タブ）など他の画面からも使えるようにする薄い入口。
 * GitHub Pages に公開された静的JSONを読むだけで、GASも認証も要らない
 * （議事録カードの api/minutes/minutes.api.js と同じ考え方）。
 *
 * 使い方:
 *   const agent = await import('/api/agent/agent.api.js')
 *   const { generatedAt, issues, phases } = await agent.recent()
 *   // phases[i] = { key, label, items: issues のうち該当フェーズのもの }
 *
 * 日付の相対表示（「16時間前」など）や「放置」判定は、更新のたびに時間が進むため
 * ここでは計算しない。呼び出し側（ダッシュボード）が描画のたびに算出すること。
 * 判定に使う閾値だけ STUCK_THRESHOLD_SECONDS として公開する。
 */

const DATA_URL = 'https://ymatsuda-cmyk.github.io/tools-beta/data/agent/dashboard.json'

/* utility/agent.html の PHASES と同じ並び・グルーピング */
const PHASES = [
  ['waiting',      '待機中',   ['created', 'queued', 'rework']],
  ['implementing', '実装中',   ['implementing', 'waiting_decision', 'preview_deploying']],
  ['approval',     '承認待ち', ['waiting_approval', 'approved']],
  ['done',         '終了',     ['completed', 'rejected', 'failed']]
]

/* utility/agent.html の STUCK_THRESHOLD_SECONDS と同じ値 */
export const STUCK_THRESHOLD_SECONDS = {
  waiting_approval: 2 * 3600,
  waiting_decision: 1 * 3600
}

/**
 * dashboard.json を取得し、フェーズごとにグルーピングして返す。
 * config.dataUrl を指定すると、既定のURLの代わりにそこから読む。
 */
export async function recent(config = {}) {
  const url = (config && config.dataUrl) || DATA_URL
  const sep = url.includes('?') ? '&' : '?'

  const res = await fetch(url + sep + 't=' + Date.now(), { cache: 'no-store' })
  if (!res.ok) throw new Error(`Issue一覧を取得できません（HTTP ${res.status}）`)

  const raw = await res.json()
  if (!raw || !Array.isArray(raw.issues)) throw new Error('Issue一覧の形式が正しくありません')

  const issues = raw.issues
    .map(normalizeIssue)
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))

  const phases = PHASES.map(([key, label, statuses]) => ({
    key,
    label,
    items: issues.filter(i => statuses.includes(i.status))
  }))

  // 未知のstatusはagent.htmlのphaseOf()と同じく最後のフェーズ（終了）に寄せる
  const known = new Set(PHASES.flatMap(([, , statuses]) => statuses))
  const orphans = issues.filter(i => !known.has(i.status))
  if (orphans.length) phases[phases.length - 1].items.push(...orphans)

  return {
    generatedAt: raw.generatedAt || null,
    issues,
    phases
  }
}

function normalizeIssue(raw) {
  return {
    number: raw.number,
    title: raw.title || '(タイトルなし)',
    status: String(raw.status || '').toLowerCase(),
    updatedAt: raw.updatedAt || null,
    issueUrl: raw.issueUrl || null,
    pullRequestUrl: raw.pullRequestUrl || null,
    previewUrl: raw.previewUrl || null
  }
}

/**
 * 放置判定。閾値を超えて同じstatusのまま更新されていない項目に立つ。
 * now を渡すとテストしやすい（既定は呼び出し時刻）。
 */
export function isStuck(issue, now = Date.now()) {
  const threshold = STUCK_THRESHOLD_SECONDS[issue.status]
  if (threshold == null || !issue.updatedAt) return false
  const seconds = (now - new Date(issue.updatedAt).getTime()) / 1000
  return seconds >= threshold
}
