import { escapeHtml } from './render.js'
import { tagStats, vocabSummary, mergeCandidates, TAIL_MAX } from '../lib/vocab.js'

/**
 * タグの整理パネル。
 *
 * 見て終わる画面にしないため、候補の隣に必ず操作(まとめる / 残す)を置く。
 * 度数の棒は「そのタグで絞り込めるか」を見るためのもので、
 * 下位(TAIL_MAX回以下)は区切りから下にまとめて、掃除待ちの列として見せる。
 */
export function openVocabPanel(items, handlers) {
  const root = document.getElementById('modal-root')

  function paint(state = {}) {
    const stats = tagStats(items)
    const summary = vocabSummary(items)
    const candidates = mergeCandidates(items, handlers.dismissed())
    const head = stats.filter((s) => !s.tail)
    const tail = stats.filter((s) => s.tail)
    const max = stats[0]?.count || 1

    const bar = (s) => `
      <div class="vb-row">
        <span class="vb-label ${s.tail ? 'tail' : ''}">${escapeHtml(s.tag)}</span>
        <span class="vb-track"><span class="vb-fill ${s.tail ? 'tail' : ''}" style="width:${Math.max(3, (s.count / max) * 100)}%"></span></span>
        <span class="vb-count">${s.count}</span>
      </div>`

    const candidateCard = (c) => `
      <div class="vc-card">
        <div class="vc-pair">
          <span class="tag">${escapeHtml(c.from)} <span class="muted">${c.fromCount}</span></span>
          <i class="ti ti-arrow-right" aria-hidden="true"></i>
          <span class="chip on">${escapeHtml(c.to)} <span style="opacity:.7">${c.toCount}</span></span>
        </div>
        <p class="vc-why">共起 ${Math.round(c.co * 100)}%${
          c.sameWord
            ? ' ・ 綴りが違うだけの同じ語'
            : c.fromCount === 1
              ? ' ・ 1件だけなので根拠は弱い'
              : c.co >= 1
                ? ` ・ ${c.fromCount}件すべてに ${escapeHtml(c.to)} が付いている`
                : ' ・ 判断が要る'
        }</p>
        <div class="row">
          <button class="btn vc-merge" data-from="${escapeHtml(c.from)}" data-to="${escapeHtml(c.to)}">まとめる</button>
          <button class="btn vc-keep" data-key="${escapeHtml(c.key)}">別物として残す</button>
        </div>
      </div>`

    root.innerHTML = `
      <div class="modal-overlay">
        <div class="modal modal-wide">
          <div class="modal-head"><span>タグの整理</span><button class="btn-ghost btn-close" aria-label="閉じる"><i class="ti ti-x"></i></button></div>
          <div class="modal-body">
            <p class="foot-note">全${summary.contents}件 ／ タグ${summary.tags}語${summary.tailTags ? ` ／ ${TAIL_MAX}回以下 ${summary.tailTags}語` : ''}${summary.untagged ? ` ／ タグなし ${summary.untagged}件` : ''}</p>
            ${state.busy ? `<p class="foot-note">${escapeHtml(state.busy)}</p>` : ''}
            ${state.error ? `<p class="error-text">${escapeHtml(state.error)}</p>` : ''}

            <h4>タグの使われ方</h4>
            ${stats.length
              ? `<div class="vb-list">
                   ${head.map(bar).join('')}
                   ${tail.length
                     ? `<div class="vb-divider"><span>ここから下は絞り込みに使えていない</span></div>${tail.map(bar).join('')}`
                     : ''}
                 </div>`
              : '<p class="muted">まだタグがありません</p>'}

            <h4>統合の候補</h4>
            <p class="foot-note">いつも一緒に付いている組は、別の観点ではなく言い換えの疑いがあります</p>
            ${candidates.length
              ? `<div class="vc-list">${candidates.map(candidateCard).join('')}</div>`
              : '<p class="muted">候補はありません</p>'}
          </div>
          <div class="modal-foot">
            <button class="btn btn-reset">残すと決めた組を戻す</button>
            <button class="btn btn-primary btn-done">閉じる</button>
          </div>
        </div>
      </div>
    `

    const close = () => (root.innerHTML = '')
    root.querySelector('.btn-close').addEventListener('click', close)
    root.querySelector('.btn-done').addEventListener('click', close)
    root.querySelector('.btn-reset').addEventListener('click', () => {
      handlers.onResetDismissed()
      paint()
    })

    root.querySelectorAll('.vc-keep').forEach((btn) =>
      btn.addEventListener('click', () => {
        handlers.onKeep(btn.dataset.key)
        paint()
      })
    )
    root.querySelectorAll('.vc-merge').forEach((btn) =>
      btn.addEventListener('click', async () => {
        const { from, to } = btn.dataset
        if (!confirm(`「${from}」が付いているコンテンツすべてを「${to}」に置き換えます。元に戻すには手作業が必要です。`)) return
        paint({ busy: `「${from}」を「${to}」にまとめています...` })
        try {
          const res = await handlers.onMerge(from, to)
          paint({ busy: `${res.updated}件を更新しました${res.failed ? `(${res.failed}件失敗)` : ''}` })
        } catch (err) {
          paint({ error: `まとめられませんでした: ${err.message || err}` })
        }
      })
    )
  }

  paint()
}
