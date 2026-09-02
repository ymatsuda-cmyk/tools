/**
 * 分野別要約・応用・活用アイデアの保存フォーマット。
 *
 * JSONではなく「Notionでそのまま読める見出し+箇条書き」で保存する。
 * 理由: このアプリを開かずにNotion上で直接読めることを優先したい。
 * JSONにすると人が読めず、Notionのプロパティ欄が実質使えなくなる。
 *
 *   ## セクション名
 *   本文(任意・複数行可)
 *   - 箇条書き
 *   - 箇条書き
 *   ## 次のセクション名
 *   ...
 *
 * 見出しの前に本文が来た場合は「見出し無しの先頭セクション」として扱う。
 */

/** @returns {{heading: string, body: string, points: string[]}[]} */
export function parseSections(text) {
  const lines = String(text ?? '').split('\n')
  const sections = []
  let current = null

  const ensure = () => {
    if (!current) {
      current = { heading: '', body: '', points: [] }
      sections.push(current)
    }
    return current
  }

  for (const raw of lines) {
    const line = raw.trimEnd()
    const heading = line.match(/^##\s+(.*)$/)
    if (heading) {
      current = { heading: heading[1].trim(), body: '', points: [] }
      sections.push(current)
      continue
    }
    const bullet = line.match(/^[-*]\s+(.*)$/)
    if (bullet) {
      ensure().points.push(bullet[1].trim())
      continue
    }
    if (!line.trim()) continue
    const s = ensure()
    s.body = s.body ? `${s.body}\n${line.trim()}` : line.trim()
  }

  return sections.filter((s) => s.heading || s.body || s.points.length)
}

/** parseSections の逆。保存前に必ずこれを通してフォーマットを揃える */
export function serializeSections(sections) {
  if (!Array.isArray(sections)) return ''
  return sections
    .map((s) => {
      const out = []
      if (s.heading) out.push(`## ${String(s.heading).trim()}`)
      if (s.body) out.push(String(s.body).trim())
      ;(s.points || []).forEach((p) => {
        const t = String(p ?? '').trim()
        if (t) out.push(`- ${t}`)
      })
      return out.join('\n')
    })
    .filter(Boolean)
    .join('\n')
}

/** セクション全体の文字数。生成量が上限に収まっているかの目安表示に使う */
export function sectionsCharCount(text) {
  return String(text ?? '').length
}
