// src/edit.js — 本文（要約・文字起こし）の手直し
import { parseT, fmtT, segmentsFromText } from './summary.js'
import { todoKey, newTodoKey } from './todos.js'

const attr = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

/** 要約 → 編集用の下書き */
export function draftFromSummary(s) {
  return {
    card: s.cardSummary || '',
    summary: s.summary || s.cardSummary || '',
    decisions: (s.decisions || []).map((d) => ({ text: d.text, t: d.t || '' })),
    agenda: (s.agenda || []).map((a) => ({
      topic: a.topic, t: a.t || '', outcome: a.outcome || '',
      points: (a.points || []).map((p) => ({ text: p.text, t: p.t || '' })),
    })),
    topics: [...(s.topics || [])],
    todos: (s.todos || []).map((t, i) => ({ key: todoKey(t, i), text: t.text, owner: t.owner || '', due: t.due || '', orphan: !!t.orphan })),
  }
}

export const blank = {
  decisions: () => ({ text: '', t: '' }),
  points: () => ({ text: '', t: '' }),
  agenda: () => ({ topic: '', t: '', outcome: '', points: [{ text: '', t: '' }] }),
  topics: () => '',
  todos: () => ({ key: '', text: '', owner: '', due: '', orphan: false }),
}

const tin = (v, extra = '') => `<input class="etm" data-f="t" value="${attr(v)}" placeholder="時刻" inputmode="numeric" aria-label="時刻（例 12:30）" ${extra}>`
const del = (list, i, ai) => `<button type="button" class="ib edel" data-act="e-del" data-list="${list}" data-i="${i}" ${ai != null ? `data-ai="${ai}"` : ''} aria-label="この行を消す">×</button>`
const add = (list, label, ai) => `<button type="button" class="btn sm" data-act="e-add" data-list="${list}" ${ai != null ? `data-ai="${ai}"` : ''}>＋ ${label}</button>`

export function editorHTML(d) {
  const rows = (list, arr, ai) => arr.map((x, i) => `
    <div class="erow" data-list="${list}" data-i="${i}" ${ai != null ? `data-ai="${ai}"` : ''}>
      <textarea class="ein" data-f="text" rows="2" aria-label="本文">${attr(x.text)}</textarea>${tin(x.t)}${del(list, i, ai)}
    </div>`).join('')
  return `<form class="editor" id="editor" data-edit="sum">
    <div class="esec" data-sec="summary">
      <div class="efield"><label for="edCard">一覧に出す短い要約</label><input id="edCard" data-k="card" value="${attr(d.card)}"></div>
      <div class="efield"><label for="edSum">サマリ</label><textarea id="edSum" data-k="summary" rows="5">${attr(d.summary)}</textarea></div>
    </div>
    <section class="esec" data-sec="agenda"><h2>論点別</h2>
      ${d.agenda.map((a, ai) => `
        <div class="eagenda" data-ai="${ai}">
          <div class="erow top"><input class="ein" data-f="topic" value="${attr(a.topic)}" placeholder="議題" aria-label="議題">${tin(a.t)}${del('agenda', ai)}</div>
          ${rows('points', a.points, ai)}
          ${add('points', '要点を追加', ai)}
          <div class="efield"><label>結論</label><input class="ein" data-f="outcome" value="${attr(a.outcome)}" placeholder="結論（無ければ空）"></div>
        </div>`).join('')}
      ${add('agenda', '議題を追加')}
      <h3 class="esub">未解決の論点</h3>
      ${d.topics.map((t, i) => `<div class="erow" data-list="topics" data-i="${i}"><input class="ein" data-f="text" value="${attr(t)}" aria-label="論点">${del('topics', i)}</div>`).join('')}
      ${add('topics', '論点を追加')}
    </section>
    <section class="esec" data-sec="decisions"><h2>決定事項</h2>${rows('decisions', d.decisions)}${add('decisions', '決定事項を追加')}</section>
    <section class="esec" data-sec="todo"><h2>ToDo</h2>
      ${d.todos.map((t, i) => `
        <div class="erow todo" data-list="todos" data-i="${i}" data-key="${attr(t.key)}" data-orphan="${t.orphan ? 1 : 0}">
          <textarea class="ein" data-f="text" rows="2" aria-label="ToDo">${attr(t.text)}</textarea>
          <input class="eow" data-f="owner" value="${attr(t.owner)}" placeholder="担当" aria-label="担当">
          <input class="eow" data-f="due" value="${attr(t.due)}" placeholder="期限" aria-label="期限">${del('todos', i)}
        </div>`).join('')}
      ${add('todos', 'ToDo を追加')}
    </section>
    <p class="muted sm">時刻は 1:23 や 1:02:03 の形で入れます（空でも構いません）。マーカーは、同じ文字が残っていれば、そのまま引き継がれます。</p>
    <div class="acts"><button type="button" class="btn main" data-act="e-save">保存</button><button type="button" class="btn" data-act="e-cancel">キャンセル</button><span class="msg" id="edMsg" role="alert"></span></div>
  </form>`
}

/** 編集画面の入力 → 下書き（行の追加・削除のとき、入力済みの内容を失わないために使う） */
export function readDraft(root) {
  const val = (el, f) => (el.querySelector(`[data-f="${f}"]`)?.value ?? '')
  const list = (name, sel = root) => [...sel.querySelectorAll(`.erow[data-list="${name}"]`)]
  return {
    card: root.querySelector('[data-k="card"]')?.value ?? '',
    summary: root.querySelector('[data-k="summary"]')?.value ?? '',
    decisions: list('decisions').map((r) => ({ text: val(r, 'text'), t: val(r, 't') })),
    agenda: [...root.querySelectorAll('.eagenda')].map((b) => {
      const top = b.querySelector('.erow.top')
      return {
        topic: val(top, 'topic'), t: val(top, 't'), outcome: b.querySelector('[data-f="outcome"]')?.value ?? '',
        points: list('points', b).map((r) => ({ text: val(r, 'text'), t: val(r, 't') })),
      }
    }),
    topics: list('topics').map((r) => val(r, 'text')),
    todos: list('todos').map((r) => ({ key: r.dataset.key || '', text: val(r, 'text'), owner: val(r, 'owner'), due: val(r, 'due'), orphan: r.dataset.orphan === '1' })),
  }
}

const bad = (t) => (t && parseT(t.trim()) == null ? t : null)

/** 下書き → 要約。時刻の書き間違いがあれば { err } を返す */
export function summaryFromDraft(d, prev, now = Date.now()) {
  const times = [...d.decisions.map((x) => x.t), ...d.agenda.flatMap((a) => [a.t, ...a.points.map((p) => p.t)])]
  const wrong = times.map(bad).find(Boolean)
  if (wrong) return { err: `時刻「${wrong}」は読み取れません。1:23 や 1:02:03 の形で入れてください。` }
  const clean = (t) => (t ? t.trim() : '')
  const s = {
    ...prev,
    cardSummary: d.card.trim(),
    summary: d.summary.trim(),
    decisions: d.decisions.filter((x) => x.text.trim()).map((x) => ({ text: x.text.trim(), t: clean(x.t) })),
    agenda: d.agenda.filter((a) => a.topic.trim() || a.points.some((p) => p.text.trim())).map((a) => ({
      topic: a.topic.trim() || '（議題）', t: clean(a.t), outcome: a.outcome.trim(),
      points: a.points.filter((p) => p.text.trim()).map((p) => ({ text: p.text.trim(), t: clean(p.t) })),
    })),
    topics: d.topics.map((x) => x.trim()).filter(Boolean),
    todos: d.todos.filter((t) => t.text.trim()).map((t) => ({
      id: t.key || newTodoKey(), text: t.text.trim(), owner: t.owner.trim(), due: t.due.trim(), ...(t.orphan ? { orphan: true } : {}),
    })),
    edited: now,
  }
  if (!s.summary && !s.cardSummary && !s.agenda.length) return { err: 'サマリか論点別のどちらかは入れてください。' }
  return { s }
}

/** 文字起こし → 編集用のテキスト（1行1発言。行頭の [m:ss] が時刻） */
export const trToText = (segments) => segments.map((x) => (x.t != null ? `[${fmtT(x.t)}] ` : '') + x.s).join('\n')

/** 編集したテキスト → 文字起こし。直していない行の時刻（小数つき）は、元のまま残す */
export function textToSegments(text, prev) {
  const segs = segmentsFromText(text)
  return segs.map((x, i) => {
    const p = prev && prev[i]
    if (x.t != null && p && p.t != null && Math.floor(p.t) === x.t) return { ...x, t: p.t }
    return x
  })
}
