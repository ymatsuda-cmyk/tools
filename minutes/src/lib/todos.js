import { plainTextOf } from './markers.js'

// 突き合わせとみなす類似度のしきい値(0〜1)。
// 再生成でAIが言い回しを少し変える程度は拾い、別のToDoを誤って結び付けない値にしている。
const MATCH_THRESHOLD = 0.5

/** 比較用に、マーカータグ・空白・記号を取り除いて正規化する */
function normalize(text) {
  return plainTextOf(text || '')
    .replace(/[\s　、。,.・()()「」【】\-—/:：]/g, '')
    // どのToDoにも出てくる定型の言い回しは、似ていると誤判定する原因になるため除く
    .replace(/(を確認する|を確認|を実施する|を実施|を検討する|を依頼する|を行う|について|に関して|に対して|します|する|した|して)/g, '')
    .replace(/[のをがにはでとへもや]/g, '')
    .toLowerCase()
}

/** 文字の2-gramによるDice係数。日本語でも分かち書き無しで近さを測れる */
function similarity(a, b) {
  const x = normalize(a)
  const y = normalize(b)
  if (!x || !y) return 0
  if (x === y) return 1
  const grams = (s) => {
    const m = new Map()
    for (let i = 0; i < s.length - 1; i++) {
      const g = s.slice(i, i + 2)
      m.set(g, (m.get(g) || 0) + 1)
    }
    return m
  }
  const gx = grams(x)
  const gy = grams(y)
  let overlap = 0
  gx.forEach((n, g) => { overlap += Math.min(n, gy.get(g) || 0) })
  const nx = Math.max(x.length - 1, 0)
  const ny = Math.max(y.length - 1, 0)
  if (!nx || !ny) return 0
  const dice = (2 * overlap) / (nx + ny)
  // AIが言い換えて文が長くなると、Dice係数は両方の長さで割るぶん低く出すぎる。
  // 短い方を基準にした重なり率も見て、ただし極端に短い一致だけで結び付かないよう
  // Dice係数にも最低ラインを設ける。
  const overlapRatio = overlap / Math.min(nx, ny)
  return dice >= 0.25 ? Math.max(dice, overlapRatio * 0.9) : dice
}

/**
 * 修正前のToDoと修正後のToDoを突き合わせ、チェック状態と経過ログを引き継ぐ。
 *
 * - 修正後の各ToDoに、最も近い修正前ToDoを1対1で割り当てる(しきい値未満は割り当てない)
 * - keepOrphans=true のとき、どこにも割り当てられなかった修正前ToDoのうち
 *   経過ログを持つものは捨てずに末尾へ残す(orphan: true)
 *   要約の再生成ではAIがToDoを落とすことがあるため true、
 *   人が意図して削除する手編集では false にする
 *
 * @returns {{ todos: object[], lostLogs: object[] }} lostLogs は引き継げず残さなかった経過付きToDo
 */
export function mergeTodos(prevTodos, nextTodos, { keepOrphans }) {
  const prev = (prevTodos || []).map((t) => (typeof t === 'string' ? { text: t, done: false, logs: [] } : t))
  const next = (nextTodos || []).map((t) => (typeof t === 'string' ? { text: t, done: false } : t))

  // 全組み合わせの類似度を計算し、高い順に貪欲に1対1で割り当てる
  const pairs = []
  next.forEach((n, ni) => {
    prev.forEach((p, pi) => {
      const score = similarity(n.text, p.text)
      if (score >= MATCH_THRESHOLD) pairs.push({ ni, pi, score })
    })
  })
  pairs.sort((a, b) => b.score - a.score)

  const nextToPrev = new Map()
  const usedPrev = new Set()
  pairs.forEach(({ ni, pi }) => {
    if (nextToPrev.has(ni) || usedPrev.has(pi)) return
    nextToPrev.set(ni, pi)
    usedPrev.add(pi)
  })

  const todos = next.map((n, ni) => {
    const pi = nextToPrev.get(ni)
    if (pi === undefined) return { text: n.text, done: !!n.done, logs: n.logs || [] }
    const p = prev[pi]
    return {
      text: n.text,
      // 手編集で明示的に付けたチェックは尊重し、再生成(常にfalse)では前回の状態を引き継ぐ
      done: !!(n.done || p.done),
      logs: [...(p.logs || []), ...(n.logs || [])],
    }
  })

  const unmatched = prev.filter((p, pi) => !usedPrev.has(pi) && (p.logs || []).length)
  if (keepOrphans) {
    unmatched.forEach((p) => todos.push({ ...p, orphan: true }))
    return { todos, lostLogs: [] }
  }
  return { todos, lostLogs: unmatched }
}
