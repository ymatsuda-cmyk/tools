/**
 * Issue状態ダッシュボード 一覧取得API
 *
 * utility/agent.html（agent-dashboard-version: 4）が読んでいるのと同じ
 * data/agent/dashboard.json を、ダッシュボード（稼働状況タブ）など他の画面からも
 * 使えるようにする薄い入口。GitHub Pages に公開された静的JSONを読むだけで、
 * GASも認証も要らない（議事録カードの api/minutes/minutes.api.js と同じ考え方）。
 *
 * 分類ロジック（classify）は utility/agent.html のものと完全に同じにしてある。
 * 片方だけ直すと表示件数がずれるので、担当の振り分けやしきい値を変えるときは
 * 必ず両方を直すこと。
 *
 * 使い方:
 *   const agent = await import('/api/agent/agent.api.js')
 *   const { issues } = await agent.recent()
 *   const groups = agent.groupByActor(issues)   // { human, cline, pa, system, done }
 *
 * note/warn は「いま」からの経過時間で変わるため、recent() では計算しない。
 * classify() / groupByActor() は呼び出し側が描画のたびに（=最新の Date.now() で）
 * 呼ぶこと。elapsed表示も同様に formatElapsed() を描画のたびに呼ぶこと。
 */

const DATA_URL = 'https://ymatsuda-cmyk.github.io/tools-beta/data/agent/dashboard.json'

/* reply/ の通知がこの秒数を超えて未送信なら「Power Automate待ち」とみなす */
export const PA_STALE_SECONDS = 180

/* utility/agent.html の ACTORS と同じ並び・色 */
export const ACTORS = [
  { key:'human',  label:'人の対応待ち',   sub:'Teamsでの回答・承認・修正指示', color:['#FAEEDA', '#BA7517', '#633806'] },
  { key:'cline',  label:'Cline作業中',    sub:'VS Code上で実装中',             color:['#E6F1FB', '#378ADD', '#0C447C'] },
  { key:'pa',     label:'自動連携待ち',   sub:'Power AutomateのTeams送信待ち', color:['#EEEDFE', '#7F77DD', '#3C3489'] },
  { key:'system', label:'システム処理中', sub:'Pythonの着手・公開・マージ',     color:['#F1EFE8', '#888780', '#444441'] }
]

/* この状態になったら「終了」扱い。archived な項目も終了に回る（下のclassify参照） */
export const TERMINAL = ['completed', 'rejected']

/* status -> [担当, 説明, 警告を出す秒数(null=出さない), 警告文]。utility/agent.html の STATUS_MAP と同じ */
export const STATUS_MAP = {
  created:           ['system', '着手待ち（空き枠待ち）', 1800, '停滞'],
  queued:            ['system', '着手準備中',             1800, '停滞'],
  rework:            ['system', '再実装の着手待ち',       1800, '停滞'],
  preview_deploying: ['system', '検証サイトへ公開中',     1800, '停滞'],
  approved:          ['system', 'マージ処理中',           1800, '停滞'],
  implementing:      ['cline',  'Clineが実装中',          3600, '長時間・停止の可能性'],
  waiting_decision:  ['human',  'Teamsで質問の回答待ち',  3600, '放置'],
  waiting_approval:  ['human',  'Teamsで承認待ち',        7200, '放置'],
  failed:            ['human',  '停止・要対応',           null, '']
}

/**
 * dashboard.json を取得する。フェーズ分けや経過時間の計算はしない
 * （時間とともに変わるため、描画のたびに classify()/groupByActor() を呼ぶこと）。
 * config.dataUrl を指定すると、既定のURLの代わりにそこから読む。
 */
export async function recent(config = {}) {
  const url = (config && config.dataUrl) || DATA_URL
  const sep = url.includes('?') ? '&' : '?'

  const res = await fetch(url + sep + 't=' + Date.now(), { cache: 'no-store' })
  if (!res.ok) throw new Error(`Issue一覧を取得できません（HTTP ${res.status}）`)

  const raw = await res.json()
  if (!raw || !Array.isArray(raw.issues)) throw new Error('Issue一覧の形式が正しくありません')

  const repository = raw.repository || ''
  const issues = raw.issues
    .map(i => normalizeIssue(i, repository))
    .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))

  return { generatedAt: raw.generatedAt || null, repository, issues }
}

function normalizeIssue(raw, repository) {
  const repoBase = repository ? 'https://github.com/' + repository : ''
  const issueUrl = raw.issueUrl || (repoBase && raw.number != null ? repoBase + '/issues/' + raw.number : null)

  // PRのURLが状態ファイルに無い場合は、ブランチ名でPRを検索するページへ飛ぶ（agent.htmlと同じ）
  let prUrl = raw.pullRequestUrl || null
  if (!prUrl && raw.branch && repoBase) {
    prUrl = repoBase + '/pulls?q=' + encodeURIComponent('is:pr head:' + raw.branch)
  }

  return {
    number: raw.number,
    title: raw.title || '(タイトルなし)',
    status: String(raw.status || '').toLowerCase(),
    statusLabel: raw.statusLabel || raw.status || '',
    error: raw.error || '',
    updatedAt: raw.updatedAt || null,
    issueUrl,
    prUrl,
    previewUrl: raw.previewUrl || null,
    branch: raw.branch || null,
    pendingNotifications: raw.pendingNotifications || { count: 0, oldestAt: null },
    archived: !!raw.archived
  }
}

export function secondsSince(iso, now = Date.now()) {
  if (!iso) return null
  const t = new Date(iso).getTime()
  if (isNaN(t)) return null
  return Math.max(0, Math.floor((now - t) / 1000))
}

/** 「N秒前」「N分前」…utility/agent.html の formatElapsed と同じ計算 */
export function formatElapsed(seconds) {
  if (seconds == null) return ''
  if (seconds < 60) return seconds + '秒前'
  const minutes = Math.floor(seconds / 60)
  if (minutes < 60) return minutes + '分前'
  const hours = Math.floor(minutes / 60)
  if (hours < 24) return hours + '時間' + (minutes % 60) + '分前'
  return Math.floor(hours / 24) + '日前'
}

/**
 * 1件を担当(actor)に振り分ける。utility/agent.html の classify() と完全に同じロジック。
 * 戻り値: { actor: 'human'|'cline'|'pa'|'system'|'done', note?, warn? }
 * archived または terminal(completed/rejected) は、途中の状態がどうであれ「終了」に回す
 * （通知の取り残しで対応待ち列が埋まるのを防ぐための挙動。agent.htmlの実装意図をそのまま踏襲）。
 */
export function classify(issue, now = Date.now()) {
  const status = (issue.status || '').toLowerCase()
  const pending = issue.pendingNotifications || {}
  const pendingAge = secondsSince(pending.oldestAt, now)

  if (issue.archived || TERMINAL.includes(status)) return { actor: 'done' }

  if (pending.count > 0 && pendingAge != null && pendingAge >= PA_STALE_SECONDS) {
    return {
      actor: 'pa',
      note: `Teams通知 ${pending.count}件が未送信`,
      warn: formatElapsed(pendingAge).replace('前', '') + ' 滞留'
    }
  }

  const entry = STATUS_MAP[status] || ['system', issue.statusLabel || status, null, '']
  const age = secondsSince(issue.updatedAt, now)
  let note = entry[1]
  if (status === 'failed' && issue.error) note += '：' + issue.error
  const warn = entry[2] != null && age != null && age >= entry[2] ? entry[3] : ''
  return { actor: entry[0], note, warn }
}

/**
 * issues を担当ごとにグルーピングする。更新日時の新しい順は recent() の時点で済んでいる前提。
 * 戻り値: { human:[{issue,result}], cline:[...], pa:[...], system:[...], done:[...] }
 */
export function groupByActor(issues, now = Date.now()) {
  const groups = { human: [], cline: [], pa: [], system: [], done: [] }
  for (const issue of issues) {
    const result = classify(issue, now)
    groups[result.actor].push({ issue, result })
  }
  return groups
}

/** 終了に振り分けられた1件のラベル。utility/agent.html の done セクションと同じ表記 */
export function doneLabel(issue) {
  const orphan = !TERMINAL.includes(issue.status)
  return orphan ? `退避済み（${issue.statusLabel || issue.status}のまま）` : (issue.statusLabel || issue.status)
}
