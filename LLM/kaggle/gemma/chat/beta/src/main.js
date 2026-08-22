import { h, clear } from './lib/dom.js'
import { activeProfile, connectionOf, loadSettings, saveSettings } from './lib/settings.js'
import {
  createConversation,
  db,
  deleteConversation,
  listConversations,
  listMessages,
  touchConversation,
} from './lib/db.js'
import { streamChat } from './lib/client.js'
import { estimateTokens, recordCalibration } from './lib/tokens.js'
import { createThrottledRenderer } from './lib/markdown.js'
import { createMessageList } from './ui/message-list.js'
import { createComposer } from './ui/composer.js'
import { openSettings } from './ui/settings-dialog.js'
import { createEndpointPanel } from './ui/endpoint-panel.js'
import { setupDropzone } from './ui/dropzone.js'

const $ = (id) => document.getElementById(id)

let settings = loadSettings()
let convId = null
let controller = null

const endpointPanel = createEndpointPanel($('endpoint-panel'), {
  getSettings: () => settings,
  onSelect: (id) => {
    settings.activeId = id
    saveSettings(settings)
    renderTopbar()
    refresh()
  },
})
const list = createMessageList($('messages'))
const composer = createComposer($('composer'), { onSend, onStop, onCompress, onNewChat })

/**
 * 添付をモデルに渡す形へ変換する。
 * テキスト系はタグで囲んで本文に展開し、画像は OpenAI 形式の
 * image_url パートとして送る（プロキシ側で Ollama の images に変換）。
 */
function toWire(m) {
  const atts = m.attachments ?? []
  if (!atts.length) return { role: m.role, content: m.content }

  const docs = atts.filter((a) => a.kind !== 'image')
  const images = atts.filter((a) => a.kind === 'image' && a.dataUrl)

  const blocks = docs.map((a) => `<${a.kind} name="${a.name}">\n${a.text}\n</${a.kind}>`)
  const text = blocks.length ? `${blocks.join('\n\n')}\n\n${m.content}` : m.content

  if (!images.length) return { role: 'user', content: text }

  return {
    role: 'user',
    content: [
      { type: 'text', text },
      ...images.map((a) => ({ type: 'image_url', image_url: { url: a.dataUrl } })),
    ],
  }
}

function renderTopbar() {
  const p = activeProfile(settings)
  $('model-name').textContent = p?.model || '—'
}

let renamingId = null

async function renderSidebar() {
  const root = $('conv-list')
  const convs = await listConversations()
  clear(root)

  for (const c of convs) {
    if (c.id === renamingId) {
      const input = h('input', { value: c.title })
      const commit = async () => {
        const title = input.value.trim()
        renamingId = null
        if (title && title !== c.title) await touchConversation(c.id, title)
        await renderSidebar()
      }
      root.append(
        h(
          'div',
          { class: 'conv renaming' },
          input,
          h('button', {
            class: 'icon',
            'aria-label': '確定',
            onClick: commit,
          }, h('i', { class: 'ti ti-check', 'aria-hidden': 'true' })),
          h('button', {
            class: 'icon',
            'aria-label': 'キャンセル',
            onClick: () => {
              renamingId = null
              renderSidebar()
            },
          }, h('i', { class: 'ti ti-x', 'aria-hidden': 'true' })),
        ),
      )
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') commit()
        if (e.key === 'Escape') {
          renamingId = null
          renderSidebar()
        }
      })
      input.focus()
      input.select()
      continue
    }

    root.append(
      h(
        'div',
        {
          class: c.id === convId ? 'conv active' : 'conv',
          onClick: () => selectConversation(c.id),
        },
        h('span', { text: c.title }),
        h(
          'div',
          { class: 'conv-tools' },
          h('button', {
            class: 'icon',
            'aria-label': '名前を変更',
            onClick: (e) => {
              e.stopPropagation()
              renamingId = c.id
              renderSidebar()
            },
          }, h('i', { class: 'ti ti-pencil', 'aria-hidden': 'true' })),
          h('button', {
            class: 'icon',
            'aria-label': '削除',
            onClick: async (e) => {
              e.stopPropagation()
              if (!confirm(`「${c.title}」を削除しますか？`)) return
              await deleteConversation(c.id)
              if (c.id === convId) convId = null
              await refresh()
            },
          }, h('i', { class: 'ti ti-trash', 'aria-hidden': 'true' })),
        ),
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
  const conn = connectionOf(settings)
  if (!conn || !conn.apiKey || !conn.baseUrl) {
    openSettings(settings, applySettings)
    return
  }
  // 起動制御を設定している場合、停止中なら先に知らせる
  if (conn && activeProfile(settings)?.gasUrl && endpointPanel.activeIsStopped()) {
    list.showNotice('バックエンドが停止しています。サイドバーの「起動」を押してから、3〜5分待って再送してください。')
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
  const estimated = wire.reduce(
    (sum, w) =>
      sum +
      (typeof w.content === 'string'
        ? estimateTokens(w.content)
        : w.content.reduce((t, part) => t + (part.type === 'text' ? estimateTokens(part.text) : 300), 0)),
    0,
  )

  controller = new AbortController()
  composer.setBusy(true)

  const body = list.beginStream()
  const renderer = createThrottledRenderer(body)

  let acc = ''
  let promptTokens
  let completionTokens
  let error

  try {
    for await (const ev of streamChat(conn, wire, controller.signal)) {
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

  // 画像を含まない場合のみ、実測値で推定係数を較正する
  const hadImage = prior.some((m) => (m.attachments ?? []).some((a) => a.kind === 'image'))
  if (!hadImage) recordCalibration(estimated, promptTokens)

  // 推定は必ず誤差を持つので、割合と絶対量の両方を満たしたときだけ警告する
  const truncated =
    promptTokens !== undefined &&
    promptTokens < estimated * 0.6 &&
    estimated - promptTokens > 500

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
  endpointPanel.rerender()
  endpointPanel.refresh()
  refresh()
}

setupDropzone(document.querySelector('.main'), (files) => composer.addFiles(files))

$('new-chat').addEventListener('click', onNewChat)
$('open-settings').addEventListener('click', () => openSettings(settings, applySettings))

renderTopbar()
endpointPanel.refresh()
const convs = await listConversations()
if (convs.length) convId = convs[0].id
await refresh()
if (!connectionOf(settings)?.apiKey) openSettings(settings, applySettings)
