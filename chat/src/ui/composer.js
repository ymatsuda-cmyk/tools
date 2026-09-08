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
 * 入力欄。
 *
 * 重要: 構造は一度だけ組み立て、以降は値の書き換えだけを行う。
 * 入力のたびに DOM を作り直すと textarea が一瞬 DOM から外れ、
 * 1文字ごとにフォーカスが飛ぶ。
 *
 * @param {HTMLElement} root
 * @param {{ onSend: Function, onStop: Function, onCompress: Function, onNewChat: Function }} handlers
 */
export function createComposer(root, handlers) {
  let atts = []
  let historyTokens = 0
  let busy = false
  let error = null

  const meterFill = h('div', { class: 'fill' })
  const meterVal = h('span', { class: 'val' })
  const meter = h(
    'div',
    { class: 'meter' },
    h('span', { text: 'コンテキスト' }),
    h('div', { class: 'track' }, meterFill),
    meterVal,
  )

  const warnSlot = h('div')
  const errSlot = h('div')
  const attsSlot = h('div', { class: 'atts' })

  const textarea = h('textarea', { placeholder: '続けて質問する…' })
  const fileInput = h('input', {
    type: 'file',
    multiple: true,
    hidden: true,
    accept:
      '.pdf,.docx,.xlsx,.xlsm,.xls,.txt,.md,.csv,.tsv,.json,.log,.ts,.tsx,.js,.py,.sql,.html,.css,.xml,.yaml,.yml',
  })

  const hint = h('span', { class: 'hint' })
  const sendBtn = h('button', { class: 'primary', text: '送信' })
  const stopBtn = h('button', { text: '停止' })
  const actions = h('span', { class: 'actions' }, hint, sendBtn)

  root.append(
    meter,
    warnSlot,
    errSlot,
    attsSlot,
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
        actions,
      ),
    ),
  )

  const totals = () => {
    const attTokens = atts.reduce((s, a) => s + a.tokens, 0)
    const total = historyTokens + attTokens + estimateTokens(textarea.value)
    return { total, over: total > MAX_INPUT_TOKENS }
  }

  const worst = () => (atts.length ? atts.reduce((a, b) => (a.tokens >= b.tokens ? a : b)) : null)

  function update() {
    const { total, over } = totals()
    const w = worst()

    meter.className = over ? 'meter over' : 'meter'
    meterFill.style.width = `${Math.min(100, Math.round((total / MAX_INPUT_TOKENS) * 100))}%`
    meterVal.textContent = `${formatTokens(total)} / ${formatTokens(MAX_INPUT_TOKENS)}`

    clear(warnSlot)
    if (over) {
      warnSlot.append(
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
                  renderAtts()
                },
              }),
            h('button', { text: '履歴を要約して圧縮', onClick: handlers.onCompress }),
            h('button', { text: '新しいチャットに分ける', onClick: handlers.onNewChat }),
          ),
        ),
      )
    }

    clear(errSlot)
    if (error) errSlot.append(h('div', { class: 'warn', text: error }))

    clear(actions)
    if (busy) {
      actions.append(stopBtn)
    } else {
      hint.textContent = over ? '送信できません' : `応答まで${formatWait(estimateWaitSec(total))}`
      sendBtn.disabled = over
      actions.append(hint, sendBtn)
    }

    for (const el of attsSlot.children) {
      const hit = over && w && el.dataset.attId === w.id
      el.classList.toggle('over', Boolean(hit))
    }
  }

  function renderAtts() {
    clear(attsSlot)
    for (const a of atts) {
      const card = attachmentCard(a, {
        onRemove: () => {
          atts = atts.filter((x) => x.id !== a.id)
          renderAtts()
        },
      })
      card.dataset.attId = a.id
      attsSlot.append(card)
    }
    update()
  }

  async function addFiles(files) {
    error = null
    update()
    for (const f of Array.from(files ?? [])) {
      try {
        atts.push(await parseFile(f))
      } catch (e) {
        error = e.message
      }
      renderAtts()
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
    renderAtts()
    handlers.onSend(text, sending)
  }

  textarea.addEventListener('input', update)
  textarea.addEventListener('paste', (e) => {
    // スクリーンショットやファイルのペーストを優先して拾う
    const files = Array.from(e.clipboardData.files ?? [])
    if (files.length) {
      e.preventDefault()
      addFiles(files)
      return
    }
    const pasted = e.clipboardData.getData('text')
    if (pasted.length <= PASTE_THRESHOLD) return
    e.preventDefault()
    atts.push(packPaste(pasted, atts.filter((a) => a.kind === 'paste').length + 1))
    renderAtts()
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
  sendBtn.addEventListener('click', send)
  stopBtn.addEventListener('click', () => handlers.onStop())

  update()

  return {
    addFiles,
    setHistoryTokens(n) {
      historyTokens = n
      update()
    },
    setBusy(v) {
      busy = v
      update()
    },
    focus() {
      textarea.focus()
    },
  }
}
