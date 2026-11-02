// src/lib/summary-text.js — 既にある要約テキストを、議事録の要約構造に変換する
//
// 対応する形式:
//   タイプ1: 「## 会議情報 / ## 議事メモ / ## 次の段取り / ## AIからの提案」の Markdown 形式
//   タイプ2: 「会議のメモ:」で始まり、議題ごとに概要とインデント付きの要点が並ぶ形式
//            (末尾に「フォローアップ タスク:」)
//
// 戻り値は summarize.js の generateSummary と同じ形(model を除く):
//   { cardSummary, agenda: [{topic, points, outcome}], decisions, todos: [{text, done}], topics }
// meetingAt は日時が読み取れたときだけ 'YYYY-MM-DDTHH:mm'(datetime-local 用)で付く。

const CARD_MAX = 100

/** 'type1' | 'type2' | null */
export function detectSummaryType(text) {
  const t = String(text || '')
  if (/^##\s*(会議情報|議事メモ)/m.test(t)) return 'type1'
  if (/^\s*会議の?メモ\s*[:：]/m.test(t) || /^\s*フォローアップ\s*タスク\s*[:：]/m.test(t)) return 'type2'
  return null
}

export function parseSummaryText(text) {
  const src = String(text || '').replace(/\r\n?/g, '\n')
  const type = detectSummaryType(src)
  if (!type) throw new Error('形式を判別できません。タイプ1(## 議事メモ)かタイプ2(会議のメモ:)の要約を貼り付けてください。')
  const parsed = type === 'type1' ? parseType1(src) : parseType2(src)
  if (!parsed.agenda.length && !parsed.todos.length) {
    throw new Error('議題もToDoも読み取れませんでした。貼り付けた内容を確認してください。')
  }
  return { type, ...parsed }
}

/** 議題名を並べた一覧カード用の短い要約。長すぎる場合は切る */
function cardSummaryOf(agenda) {
  const s = '議題: ' + agenda.map((a) => a.topic).join('／')
  return s.length > CARD_MAX ? s.slice(0, CARD_MAX - 1) + '…' : s
}

/** 「[挿入を追加]」のようなテンプレートの未記入欄 */
function isPlaceholder(s) {
  return /^\[[^\]]*(挿入|追加)[^\]]*\]$/.test(s.trim())
}

// ------------------------------------------------------------ タイプ1

function parseType1(src) {
  const sections = {}
  let cur = null
  src.split('\n').forEach((line) => {
    const m = line.match(/^##\s*(.+?)\s*$/)
    if (m) { cur = m[1]; sections[cur] = []; return }
    if (cur) sections[cur].push(line)
  })
  const sec = (name) => sections[Object.keys(sections).find((k) => k.includes(name))] || []

  // 会議情報: 「> 日付: 2026-10-09 12:52:12」
  let meetingAt = ''
  sec('会議情報').forEach((l) => {
    const m = l.match(/日付\s*[:：]\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/)
    if (m) {
      const p = (n) => String(n).padStart(2, '0')
      meetingAt = `${m[1]}-${p(m[2])}-${p(m[3])}T${p(m[4] || 0)}:${m[5] || '00'}`
    }
  })

  // 議事メモ: 「- トピック: X」の下に、字下げした要点が並ぶ
  const agenda = []
  const decisions = []
  let item = null
  sec('議事メモ').forEach((line) => {
    if (!line.trim()) return
    const indented = /^\s+/.test(line)
    const body = line.replace(/^\s*[-*・]\s*/, '').trim()
    const t = body.match(/^トピック\s*[:：]\s*(.+)$/)
    if (!indented && t) {
      item = { topic: t[1].trim(), points: [], outcome: '' }
      agenda.push(item)
      return
    }
    if (!item || !body) return
    const c = body.match(/^結論\s*[:：]\s*(.+)$/)
    if (c) {
      item.outcome = c[1].trim()
      decisions.push(`${item.topic}: ${item.outcome}`)
    } else {
      item.points.push(body)
    }
  })
  agenda.forEach((a) => { if (!a.outcome) a.outcome = '継続検討' })

  // 次の段取り: 「- [ ] 内容」「- [x] 内容」
  const todos = []
  sec('次の段取り').forEach((line) => {
    const m = line.match(/^\s*[-*]\s*\[([ xX])\]\s*(.*)$/)
    if (!m) return
    const text = m[2].trim()
    if (!text || isPlaceholder(text)) return
    todos.push({ text, done: m[1].toLowerCase() === 'x' })
  })

  // AIからの提案: 「> 1. 内容」の番号付きの行を未解決の論点にする
  const topics = []
  sec('AIからの提案').forEach((line) => {
    const m = line.match(/^>?\s*\d+[.．)）]\s*(.+)$/)
    if (m) topics.push(m[1].trim())
  })

  return { cardSummary: cardSummaryOf(agenda), agenda, decisions, todos, topics, meetingAt }
}

// ------------------------------------------------------------ タイプ2

const FOLLOWUP_RE = /^\s*フォローアップ\s*タスク\s*[:：]?\s*$/
const isIndented = (line) => /^(\t|\s{2,})\S/.test(line) || /^[\t　]+\S/.test(line)
const isHeader = (line) => /[:：]\s*$/.test(line.trim()) && !isIndented(line)

function parseType2(src) {
  const lines = src.split('\n')
  const cut = lines.findIndex((l) => FOLLOWUP_RE.test(l))
  const main = cut === -1 ? lines : lines.slice(0, cut)
  const follow = cut === -1 ? [] : lines.slice(cut + 1)

  // 議題: 「議題名:」の次の字下げなしの行が概要、字下げした「見出し: 本文」が要点
  const agenda = []
  let item = null
  main.forEach((raw) => {
    const line = raw.replace(/\s+$/, '')
    const t = line.trim()
    if (!t) return
    if (/^AI\s*によって生成/.test(t) || /^会議の?メモ\s*[:：]?$/.test(t)) return
    if (isIndented(line)) {
      if (item) item.points.push(t)
    } else if (isHeader(line)) {
      item = { topic: t.replace(/[:：]\s*$/, '').trim(), points: [], outcome: '' }
      agenda.push(item)
    } else if (item) {
      item.outcome = item.outcome ? `${item.outcome} ${t}` : t // 概要(複数行なら連結)
    }
  })

  // フォローアップ: 「題名:」の次の行が「内容 (担当者)」
  const todos = []
  let title = ''
  follow.forEach((raw) => {
    const t = raw.trim()
    if (!t) return
    if (isHeader(raw)) {
      title = t.replace(/[:：]\s*$/, '').trim()
    } else {
      todos.push({ text: title ? `${title}: ${t}` : t, done: false })
      title = ''
    }
  })

  return { cardSummary: cardSummaryOf(agenda), agenda, decisions: [], todos, topics: [], meetingAt: '' }
}
