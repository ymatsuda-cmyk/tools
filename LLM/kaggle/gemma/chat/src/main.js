import { h, clear } from './lib/dom.js'
import { loadSettings, saveSettings } from './lib/settings.js'
import {
  createConversation,
  db,
  deleteConversation,
  listConversations,
  listMessages,
  touchConversation,
} from './lib/db.js'
import { streamChat } from './lib/client.js'
import { estimateTokens } from './lib/tokens.js'
import { createThrottledRenderer } from './lib/markdown.js'
import { createMessageList } from './ui/message-list.js'
import { createComposer } from './ui/composer.js'
import { openSettings } from './ui/settings-dialog.js'
import { createStatusBar } from './ui/status-bar.js'

const $ = (id) => document.getElementById(id)

let settings = loadSettings()
let convId = null
let controller = null

const statusBar = createStatusBar($('status-bar'), { getSettings: () => settings })
const list = createMessageList($('messages'))
const composer = createComposer($('composer'), { onSend, onStop, onCompress, onNewChat })

/** 添付を本文に展開してモデルに渡す形にする */
function toWire(m) {
  if (!m.attachments?.length) return { role: m.role, content: m.content }
  const blocks = m.attachments.map((a) => `<${a.kind} name="${a.name}">\n${a.text}\n</${a.kind}>`)
  return { role: 'user', content: `${blocks.join('\n\n')}\n\n${m.content}` }
}

function renderTopbar() {
  $('model-name').textContent = settings.model
}

async function renderSidebar() {
  const root = $('conv-list')
  const convs = await listConversations()
  clear(root)
  for (const c of convs) {
    root.append(
      h(
        'div',
        {
          class: c.id === convId ? 'conv active' : 'conv',
          onClick: () => selectConversation(c.id),
        },
        h('span', { text: c.title }),
        h('button', {
          class: 'icon',
          'aria-label': '削除',
          text: '×',
          onClick: async (e) => {
            e.stopPropagation()
            await deleteConversation(c.id)
            if (c.id === convId) convId = null
            await refresh()
          },
        }),
      ),
    )
  }
}

async function refresh() {
  await renderSidebar()
  const messages = convId ? await listMessages(convId) : []
  list.render(messages)
  composer.setHistoryTokens(
    estimateTokens(settings.systemPrompt) +
      messages.reduce(
        (s, m) =>
          s + estimateTokens(m.content) + (m.attachments ?? []).reduce((t, a) => t + a.tokens, 0),
        0,
      ),
  )
}

async function selectConversation(id) {
  convId = id
  await refresh()
}

async function onNewChat() {
  convId = await createConversation()
  await refresh()
  composer.focus()
}

function onStop() {
  controller?.abort()
}

async function onSend(text, atts) {
  if (!settings.apiKey) {
    openSettings(settings, applySettings)
    return
  }
  // 起動制御を設定している場合、停止中なら先に知らせる
  if (settings.gasUrl && statusBar.getState().key === 'stopped') {
    list.showNotice('バックエンドが停止しています。上部の「起動」を押してから、3〜5分待って再送してください。')
    return
  }
  if (convId === null) convId = await createConversation()

  await db.messages.add({
    convId,
    role: 'user',
    content: text,
    attachments: atts,
    createdAt: Date.now(),
  })

  const count = await db.messages.where('convId').equals(convId).count()
  const title = (text.trim() || atts[0]?.name || '新規チャット').slice(0, 30)
  await touchConversation(convId, count === 1 ? title : undefined)
  await refresh()

  const prior = await listMessages(convId)
  const wire = [{ role: 'system', content: settings.systemPrompt }, ...prior.map(toWire)]
  const estimated = wire.reduce((s, w) => s + estimateTokens(w.content), 0)

  controller = new AbortController()
  composer.setBusy(true)

  const body = list.beginStream()
  const renderer = createThrottledRenderer(body)

  let acc = ''
  let promptTokens
  let completionTokens
  let error

  try {
    for await (const ev of streamChat(settings, wire, controller.signal)) {
      if (ev.delta) {
        acc += ev.delta
        renderer.update(acc)
        list.scrollToEnd()
      }
      if (ev.usage) {
        promptTokens = ev.usage.prompt_tokens
        completionTokens = ev.usage.completion_tokens
      }
    }
  } catch (e) {
    if (e.name !== 'AbortError') error = e.message
  }

  renderer.finish(acc)

  // 実測が推定の 7 割を下回っていたら切り捨てを疑う
  const truncated = promptTokens !== undefined && promptTokens < estimated * 0.7

  await db.messages.add({
    convId,
    role: 'assistant',
    content: acc,
    attachments: [],
    createdAt: Date.now(),
    promptTokens,
    completionTokens,
    truncated,
    error,
  })
  await touchConversation(convId)

  controller = null
  composer.setBusy(false)
  await refresh()
}

async function onCompress() {
  if (!convId) return
  const prior = await listMessages(convId)
  if (prior.length < 3) return
  const keep = prior.slice(-2)
  const drop = prior.slice(0, -2)
  const summary = drop
    .map((m) => `${m.role === 'user' ? 'Q' : 'A'}: ${m.content.slice(0, 300)}`)
    .join('\n')

  await db.transaction('rw', db.messages, async () => {
    await db.messages.bulkDelete(drop.map((m) => m.id))
    await db.messages.add({
      convId,
      role: 'user',
      content: `【これまでの要約】\n${summary}`,
      attachments: [],
      createdAt: keep[0].createdAt - 1,
    })
  })
  await refresh()
}

function applySettings(s) {
  settings = s
  saveSettings(s)
  renderTopbar()
  statusBar.refresh()
  refresh()
}

$('new-chat').addEventListener('click', onNewChat)
$('open-settings').addEventListener('click', () => openSettings(settings, applySettings))

renderTopbar()
statusBar.refresh()
const convs = await listConversations()
if (convs.length) convId = convs[0].id
await refresh()
if (!settings.apiKey) openSettings(settings, applySettings)
