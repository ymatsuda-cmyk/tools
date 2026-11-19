// src/lib/summary-text.js — 既にある要約テキストを、議事録の要約構造に変換する
//
// 対応する形式:
//   タイプ1: 「## 会議情報 / ## 議事メモ / ## 次の段取り / ## AIからの提案」の Markdown 形式
//   タイプ2: 「会議のメモ:」で始まり、議題ごとに概要とインデント付きの要点が並ぶ形式
//            (末尾に「フォローアップ タスク:」)
//   タイプ3: 「### 会議情報 / ### 会議メモ / ### 次の段取り / ### AIによる提案」の Markdown 形式。
//            議題は「#### 議題名」、要点は「- 」の箇条書き(字下げした子の箇条書きも可)
//
// 戻り値は summarize.js の generateSummary と同じ形(model を除く):
//   { cardSummary, agenda: [{topic, points, outcome}], decisions, todos: [{text, done}], topics }
// meetingAt は日時が読み取れたときだけ 'YYYY-MM-DDTHH:mm'(datetime-local 用)で付く。

const CARD_MAX = 100

const TYPE_LABELS = { type1: 'タイプ1', type2: 'タイプ2', type3: 'タイプ3' }

/** 画面に出す形式名(「タイプ1」など) */
export function summaryTypeLabel(type) {
  return TYPE_LABELS[type] || type
}

/** 'type1' | 'type2' | 'type3' | null */
export function detectSummaryType(text) {
  const t = String(text || '')
  // タイプ3 は「###」。タイプ1 の「##」の判定は「###」には当たらない(# は空白ではないため)
  if (/^###\s*(会議情報|会議メモ|議事メモ)/m.test(t)) return 'type3'
  if (/^##\s*(会議情報|議事メモ)/m.test(t)) return 'type1'
  if (/^\s*会議の?メモ\s*[:：]/m.test(t) || /^\s*フォローアップ\s*タスク\s*[:：]/m.test(t)) return 'type2'
  return null
}

export function parseSummaryText(text) {
  const src = String(text || '').replace(/\r\n?/g, '\n')
  const type = detectSummaryType(src)
  if (!type) throw new Error('形式を判別できません。タイプ1(## 議事メモ)・タイプ2(会議のメモ:)・タイプ3(### 会議メモ)のいずれかの要約を貼り付けてください。')
  const parsed = type === 'type1' ? parseType1(src) : type === 'type2' ? parseType2(src) : parseType3(src)
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

// ------------------------------------------------------------ タイプ3

/** 箇条書きの親が「担当体制」のような短い見出し語か(句点で終わる文ではない) */
const isLabelLike = (s) => s.length <= 20 && !/[。．.]$/.test(s) && !/[:：]/.test(s)

function parseType3(src) {
  // 「##」「###」で区切る(「####」は議題なので区切りにしない)
  const sections = {}
  let cur = null
  src.split('\n').forEach((line) => {
    const m = line.match(/^#{2,3}(?!#)\s*(.+?)\s*$/)
    if (m) { cur = m[1]; sections[cur] = []; return }
    if (cur) sections[cur].push(line)
  })
  const sec = (...names) => sections[Object.keys(sections).find((k) => names.some((n) => k.includes(n)))] || []

  // 会議情報: 「> 日付: 2026-06-25 16:01:48 / 2026-06-25 16:01:48」(先頭の日時を使う)
  let meetingAt = ''
  sec('会議情報').forEach((l) => {
    if (meetingAt) return
    const m = l.match(/日付\s*[:：]\s*(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/)
    if (m) {
      const p = (n) => String(n).padStart(2, '0')
      meetingAt = `${m[1]}-${p(m[2])}-${p(m[3])}T${p(m[4] || 0)}:${m[5] || '00'}`
    }
  })

  // 会議メモ: 「#### 議題」の下に「- 要点」。字下げした子の箇条書きは、
  // 親が「担当体制」のような見出し語なら「【担当体制】子」、文なら親の次に「└ 子」として並べる
  const agenda = []
  let item = null
  let parent = null // { text, indent, label, emitted }
  const indentOf = (line) => line.match(/^[ \t　]*/)[0].replace(/\t/g, '    ').length
  const flushParent = () => {
    if (parent && !parent.emitted) { item.points.push(parent.text); parent.emitted = true }
  }
  sec('会議メモ', '議事メモ').forEach((line) => {
    const t = line.trim()
    if (!t || /^-{3,}$/.test(t)) return
    const h = line.match(/^####+\s*(.+?)\s*$/)
    if (h) {
      if (item) flushParent()
      item = { topic: h[1].trim(), points: [], outcome: '' }
      agenda.push(item)
      parent = null
      return
    }
    if (!item) { // 議題見出しのない要点は「その他」にまとめる
      item = { topic: 'その他', points: [], outcome: '' }
      agenda.push(item)
    }
    const body = t.replace(/^(?:[-*・]|\d+[.．)）])\s*/, '').trim()
    if (!body) return
    const indent = indentOf(line)
    if (parent && indent > parent.indent) {
      if (parent.label) {
        item.points.push(`【${parent.text}】${body}`)
        parent.emitted = true // 見出し語は子に付けたので、単独の要点にはしない
      } else {
        flushParent()
        item.points.push(`└ ${body}`)
      }
      parent.hasChild = true
      return
    }
    // 子を持たなかった見出し語の親は、そのまま要点にする
    flushParent()
    parent = { text: body, indent, label: isLabelLike(body), emitted: false, hasChild: false }
    if (!parent.label) { item.points.push(body); parent.emitted = true }
  })
  if (item) flushParent()

  // 次の段取り: 「- [ ] 内容」「- [x] 内容」
  const todos = []
  sec('次の段取り').forEach((line) => {
    const m = line.match(/^\s*[-*]\s*\[([ xX])\]\s*(.*)$/)
    if (!m) return
    const text = m[2].trim()
    if (!text || isPlaceholder(text)) return
    todos.push({ text, done: m[1].toLowerCase() === 'x' })
  })

  // AIによる提案: 「- 内容」または「1. 内容」を未解決の論点にする
  const topics = []
  sec('AIによる提案', 'AIからの提案').forEach((line) => {
    const m = line.match(/^>?\s*(?:[-*・]|\d+[.．)）])\s*(.+)$/)
    if (m) topics.push(m[1].trim())
  })

  return { cardSummary: cardSummaryOf(agenda), agenda, decisions: [], todos, topics, meetingAt }
}
