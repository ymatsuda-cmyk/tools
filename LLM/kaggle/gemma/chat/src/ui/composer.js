import { h, clear } from '../lib/dom.js'
import { PASTE_THRESHOLD, packPaste, parseFile } from '../lib/parsers.js'
import {
  MAX_INPUT_TOKENS,
  estimateTokens,
  estimateWaitSec,
  formatTokens,
  formatWait,
} from '../lib/tokens.js'
import { attachmentCard } from './message-list.js'

/**
 * @param {HTMLElement} root
 * @param {{ onSend: Function, onStop: Function, onCompress: Function, onNewChat: Function }} handlers
 */
export function createComposer(root, handlers) {
  let atts = []
  let historyTokens = 0
  let busy = false
  let error = null

  const textarea = h('textarea', { placeholder: '続けて質問する…' })
  const fileInput = h('input', {
    type: 'file',
    multiple: true,
    hidden: true,
    accept:
      '.pdf,.docx,.xlsx,.xlsm,.xls,.txt,.md,.csv,.tsv,.json,.log,.ts,.tsx,.js,.py,.sql,.html,.css,.xml,.yaml,.yml',
  })

  const totals = () => {
    const attTokens = atts.reduce((s, a) => s + a.tokens, 0)
    const total = historyTokens + attTokens + estimateTokens(textarea.value)
    return { total, over: total > MAX_INPUT_TOKENS }
  }

  const worst = () =>
    atts.length ? atts.reduce((a, b) => (a.tokens >= b.tokens ? a : b)) : null

  async function addFiles(files) {
    error = null
    render()
    for (const f of Array.from(files ?? [])) {
      try {
        atts.push(await parseFile(f))
      } catch (e) {
        error = e.message
      }
      render()
    }
  }

  function send() {
    const { over } = totals()
    if (busy || over) return
    if (!textarea.value.trim() && atts.length === 0) return
    const text = textarea.value
    const sending = atts
    textarea.value = ''
    atts = []
    render()
    handlers.onSend(text, sending)
  }

  textarea.addEventListener('input', render)
  textarea.addEventListener('paste', (e) => {
    const pasted = e.clipboardData.getData('text')
    if (pasted.length <= PASTE_THRESHOLD) return
    e.preventDefault()
    atts.push(packPaste(pasted, atts.filter((a) => a.kind === 'paste').length + 1))
    render()
  })
  textarea.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
      e.preventDefault()
      send()
    }
  })
  fileInput.addEventListener('change', () => {
    addFiles(fileInput.files)
    fileInput.value = ''
  })

  function render() {
    const { total, over } = totals()
    const ratio = Math.min(100, Math.round((total / MAX_INPUT_TOKENS) * 100))
    const w = worst()

    clear(root)

    root.append(
      h(
        'div',
        { class: over ? 'meter over' : 'meter' },
        h('span', { text: 'コンテキスト' }),
        h('div', { class: 'track' }, h('div', { class: 'fill', style: { width: `${ratio}%` } })),
        h('span', {
          class: 'val',
          text: `${formatTokens(total)} / ${formatTokens(MAX_INPUT_TOKENS)}`,
        }),
      ),
    )

    if (over) {
      root.append(
        h(
          'div',
          { class: 'warn' },
          `上限を ${(total - MAX_INPUT_TOKENS).toLocaleString()} トークン超えています。このまま送ると古い内容から無言で切り捨てられます。`,
          h(
            'div',
            { class: 'acts' },
            w &&
              h('button', {
                text: '最大の添付を外す',
                onClick: () => {
                  atts = atts.filter((a) => a.id !== w.id)
                  render()
                },
              }),
            h('button', { text: '履歴を要約して圧縮', onClick: handlers.onCompress }),
            h('button', { text: '新しいチャットに分ける', onClick: handlers.onNewChat }),
          ),
        ),
      )
    }

    if (error) root.append(h('div', { class: 'warn', text: error }))

    if (atts.length) {
      root.append(
        h(
          'div',
          { class: 'atts' },
          atts.map((a) =>
            attachmentCard(a, {
              over: over && w && a.id === w.id,
              onRemove: () => {
                atts = atts.filter((x) => x.id !== a.id)
                render()
              },
            }),
          ),
        ),
      )
    }

    const actions = busy
      ? [h('button', { text: '停止', onClick: handlers.onStop })]
      : [
          h('span', {
            class: 'hint',
            text: over ? '送信できません' : `応答まで${formatWait(estimateWaitSec(total))}`,
          }),
          h('button', { class: 'primary', text: '送信', disabled: over, onClick: send }),
        ]

    root.append(
      h(
        'div',
        { class: 'inputbox' },
        textarea,
        h(
          'div',
          { class: 'inputrow' },
          h('button', { class: 'icon', text: '添付', onClick: () => fileInput.click() }),
          fileInput,
          h('span', { class: 'spacer' }),
          ...actions,
        ),
      ),
    )
  }

  render()

  return {
    setHistoryTokens(n) {
      historyTokens = n
      render()
    },
    setBusy(v) {
      busy = v
      render()
    },
    focus() {
      textarea.focus()
    },
  }
}
