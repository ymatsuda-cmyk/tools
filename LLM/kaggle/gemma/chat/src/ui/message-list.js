import { h, clear } from '../lib/dom.js'
import { attachmentTag } from '../lib/parsers.js'
import { formatTokens } from '../lib/tokens.js'
import { renderMarkdown } from '../lib/markdown.js'
import { openModal } from './modal.js'

export function attachmentCard(att, { over = false, onRemove = null } = {}) {
  return h(
    'div',
    {
      class: over ? 'att over' : 'att',
      title: att.name,
      onClick: () => previewAttachment(att),
    },
    h('span', { class: 'tag', text: attachmentTag(att.kind) }),
    h(
      'div',
      { style: { minWidth: 0 } },
      h('div', { class: 'name', text: att.name }),
      h('div', {
        class: 'meta',
        text: `${att.chars.toLocaleString()}字 · 約${formatTokens(att.tokens)} tok`,
      }),
    ),
    onRemove &&
      h('button', {
        class: 'icon',
        'aria-label': '削除',
        text: '×',
        onClick: (e) => {
          e.stopPropagation()
          onRemove()
        },
      }),
  )
}

function previewAttachment(att) {
  openModal(({ close }) =>
    h(
      'div',
      { class: 'modal', style: { maxWidth: '720px' } },
      h('h2', { text: att.name }),
      h('div', {
        class: 'hint',
        style: { marginBottom: '10px' },
        text: `${att.chars.toLocaleString()}字 · 推定 ${att.tokens.toLocaleString()} トークン`,
      }),
      h('pre', { class: 'preview', text: att.text }),
      h(
        'div',
        { style: { marginTop: '12px', textAlign: 'right' } },
        h('button', { text: '閉じる', onClick: close }),
      ),
    ),
  )
}

function messageEl(m) {
  const wrap = h('div', { class: m.role === 'user' ? 'msg user' : 'msg' })

  if (m.attachments?.length) {
    wrap.append(h('div', { class: 'atts' }, m.attachments.map((a) => attachmentCard(a))))
  }

  if (m.role === 'user') {
    wrap.append(h('div', { class: 'bubble', text: m.content }))
  } else {
    const body = h('div', { class: 'assistant-body' })
    renderMarkdown(body, m.content)
    wrap.append(body)

    if (m.truncated) {
      wrap.append(
        h('div', {
          class: 'warn',
          style: { marginTop: '8px' },
          text: `入力の一部が処理されていない可能性があります。実際に読み込まれたのは ${m.promptTokens?.toLocaleString()} トークンでした。`,
        }),
      )
    }
    if (m.error) wrap.append(h('div', { class: 'warn', text: m.error }))
  }

  return wrap
}

export function createMessageList(root) {
  let stick = true
  root.addEventListener('scroll', () => {
    stick = root.scrollHeight - root.scrollTop - root.clientHeight < 80
  })

  const scrollToEnd = () => {
    if (stick) root.scrollTop = root.scrollHeight
  }

  return {
    /** 確定済みメッセージを描き直す */
    render(messages) {
      clear(root)
      if (!messages.length) {
        root.append(h('div', { class: 'empty', text: 'ファイルを添付するか、質問を入力してください' }))
        return
      }
      for (const m of messages) root.append(messageEl(m))
      stick = true
      scrollToEnd()
    },
    /** ストリーミング用の空要素を末尾に追加して返す */
    beginStream() {
      const body = h('div', { class: 'assistant-body' })
      root.append(h('div', { class: 'msg' }, body))
      stick = true
      scrollToEnd()
      return body
    },
    /** 一時的な案内を末尾に出す（DB には保存しない） */
    showNotice(text) {
      root.append(h('div', { class: 'msg' }, h('div', { class: 'warn', text })))
      stick = true
      scrollToEnd()
    },
    scrollToEnd,
  }
}
