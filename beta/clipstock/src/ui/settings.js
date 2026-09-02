import { escapeHtml } from './render.js'
import { loadConfig, saveConfig } from '../lib/videos-config.js'
import { loadSettings, saveSettings, newConnection } from '../lib/llm-settings.js'
import { verifyCode } from '../lib/gas.js'

function roleLabel(role) {
  if (!role) return '<span class="muted">未確認</span>'
  if (role === 'err') return '<span class="error-text">権限がありません</span>'
  if (role === 'xYz') return '<span class="ok-text">管理者 — 全機能が使えます</span>'
  return `<span class="ok-text">権限: ${escapeHtml(role)}</span>`
}

/**
 * 設定モーダル。
 * AI接続の設定は localStorage キー 'gemma-chat.settings' を
 * 議事録アプリ・gemma-chat と共有しているので、どれかで足せば全部で使える。
 */
export function openSettings(onSaved) {
  const config = loadConfig()
  const settings = loadSettings()
  // 下書き。ここで編集し、保存時にまとめて反映する(キャンセル時は破棄)
  const draft = settings.connections.map((c) => ({ ...c, models: [...(c.models || [])] }))
  let activeId = settings.activeConnectionId
  let activeModel = settings.activeModel
  let verifiedRole = config.role

  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div class="overlay">
      <div class="modal">
        <h2 class="modal-title">設定</h2>

        <label class="field-label">GAS URL</label>
        <input id="cfg-gas" class="input" value="${escapeHtml(config.gasUrl)}" placeholder="https://script.google.com/macros/s/.../exec" />

        <label class="field-label">共有トークン</label>
        <input id="cfg-token" class="input" value="${escapeHtml(config.accessToken)}" placeholder="GASの ACCESS_TOKEN と同じ値" />

        <label class="field-label">コード</label>
        <div class="row">
          <input id="cfg-code" class="input grow" value="${escapeHtml(config.code)}" />
          <button id="cfg-verify" class="btn">確認</button>
        </div>
        <div id="cfg-role" class="foot-note">${roleLabel(config.role)}</div>

        <label class="field-label">AI接続</label>
        <div id="cfg-conns"></div>
        <button id="cfg-conn-add" class="btn btn-wide"><i class="ti ti-plus" aria-hidden="true"></i>接続を追加</button>

        <details class="json-block">
          <summary>JSONで一括設定</summary>
          <textarea id="cfg-json" rows="8" class="input mono"></textarea>
          <div class="row">
            <button id="cfg-json-export" class="btn">今の設定を書き出す</button>
            <button id="cfg-json-import" class="btn">この内容を反映</button>
          </div>
        </details>

        <div class="modal-foot">
          <button id="cfg-cancel" class="btn">キャンセル</button>
          <button id="cfg-save" class="btn btn-primary">保存</button>
        </div>
      </div>
    </div>
  `

  const $ = (id) => document.getElementById(id)

  function paintConns() {
    const el = $('cfg-conns')
    el.innerHTML = draft
      .map(
        (c) => `
      <div class="conn">
        <div class="row">
          <input class="input grow conn-label" data-id="${c.id}" value="${escapeHtml(c.label)}" placeholder="表示名(例: Gemini)" />
          ${draft.length > 1 ? `<button class="btn conn-del" data-id="${c.id}" aria-label="この接続を削除"><i class="ti ti-trash" aria-hidden="true"></i></button>` : ''}
        </div>
        <input class="input sm conn-base" data-id="${c.id}" value="${escapeHtml(c.baseUrl)}" placeholder="baseUrl (例: https://generativelanguage.googleapis.com/v1beta/openai)" />
        <input class="input sm conn-key" data-id="${c.id}" value="${escapeHtml(c.apiKey)}" placeholder="APIキー" />
        <div class="conn-models">
          ${(c.models || [])
            .map(
              (m) => `<span class="chip ${c.id === activeId && m === activeModel ? 'on' : ''} model-chip" data-conn="${c.id}" data-model="${escapeHtml(m)}">
                ${escapeHtml(m)}<i class="ti ti-x model-del" data-conn="${c.id}" data-model="${escapeHtml(m)}" aria-hidden="true"></i>
              </span>`
            )
            .join('') || '<span class="muted">モデル未登録</span>'}
        </div>
        <div class="row">
          <input class="input sm grow conn-new" data-id="${c.id}" placeholder="モデル名を追加(例: gemini-2.5-flash)" />
          <button class="btn conn-add" data-id="${c.id}">追加</button>
        </div>
      </div>`
      )
      .join('')

    el.querySelectorAll('.conn-label, .conn-base, .conn-key').forEach((input) => {
      input.addEventListener('input', () => {
        const c = draft.find((x) => x.id === input.dataset.id)
        if (!c) return
        if (input.classList.contains('conn-label')) c.label = input.value
        if (input.classList.contains('conn-base')) c.baseUrl = input.value
        if (input.classList.contains('conn-key')) c.apiKey = input.value
      })
    })
    el.querySelectorAll('.conn-del').forEach((btn) => {
      btn.addEventListener('click', () => {
        const i = draft.findIndex((x) => x.id === btn.dataset.id)
        if (i === -1) return
        draft.splice(i, 1)
        if (activeId === btn.dataset.id) {
          activeId = draft[0]?.id ?? null
          activeModel = draft[0]?.models?.[0] ?? null
        }
        paintConns()
      })
    })
    el.querySelectorAll('.model-chip').forEach((chip) => {
      chip.addEventListener('click', (e) => {
        if (e.target.classList.contains('model-del')) return
        activeId = chip.dataset.conn
        activeModel = chip.dataset.model
        paintConns()
      })
    })
    el.querySelectorAll('.model-del').forEach((x) => {
      x.addEventListener('click', (e) => {
        e.stopPropagation()
        const c = draft.find((y) => y.id === x.dataset.conn)
        if (!c) return
        c.models = c.models.filter((m) => m !== x.dataset.model)
        if (activeId === c.id && activeModel === x.dataset.model) activeModel = c.models[0] ?? null
        paintConns()
      })
    })
    el.querySelectorAll('.conn-add').forEach((btn) => {
      btn.addEventListener('click', () => {
        const input = el.querySelector(`.conn-new[data-id="${btn.dataset.id}"]`)
        const name = input.value.trim()
        if (!name) return
        const c = draft.find((x) => x.id === btn.dataset.id)
        if (!c) return
        if (!c.models.includes(name)) c.models.push(name)
        if (!activeId) {
          activeId = c.id
          activeModel = name
        }
        input.value = ''
        paintConns()
      })
    })
  }
  paintConns()

  $('cfg-conn-add').addEventListener('click', () => {
    draft.push(newConnection({ label: `接続${draft.length + 1}` }))
    paintConns()
  })

  $('cfg-cancel').addEventListener('click', () => (root.innerHTML = ''))

  $('cfg-json-export').addEventListener('click', () => {
    $('cfg-json').value = JSON.stringify(
      {
        gasUrl: $('cfg-gas').value.trim(),
        accessToken: $('cfg-token').value.trim(),
        code: $('cfg-code').value.trim(),
        connections: draft.map(({ label, baseUrl, apiKey, models }) => ({ label, baseUrl, apiKey, models })),
        activeConnectionLabel: draft.find((c) => c.id === activeId)?.label,
        activeModel,
      },
      null,
      2
    )
  })

  $('cfg-json-import').addEventListener('click', () => {
    let parsed
    try {
      parsed = JSON.parse($('cfg-json').value)
    } catch (err) {
      alert('JSONの形式が不正です: ' + (err.message || err))
      return
    }
    if (parsed.gasUrl !== undefined) $('cfg-gas').value = parsed.gasUrl
    if (parsed.accessToken !== undefined) $('cfg-token').value = parsed.accessToken
    if (parsed.code !== undefined) $('cfg-code').value = parsed.code

    if (Array.isArray(parsed.connections) && parsed.connections.length) {
      draft.length = 0
      parsed.connections.forEach((c) => draft.push(newConnection(c)))
      const match = draft.find((c) => c.label === parsed.activeConnectionLabel) || draft[0]
      activeId = match.id
      activeModel = match.models?.includes(parsed.activeModel) ? parsed.activeModel : match.models?.[0] ?? null
      paintConns()
    }
    if (parsed.code) $('cfg-verify').click()
  })

  $('cfg-verify').addEventListener('click', async () => {
    const gasUrl = $('cfg-gas').value.trim()
    const roleEl = $('cfg-role')
    if (!gasUrl) {
      roleEl.innerHTML = '<span class="error-text">GAS URLを先に入力してください</span>'
      return
    }
    roleEl.textContent = '確認中...'
    try {
      const res = await verifyCode(gasUrl, $('cfg-code').value.trim())
      verifiedRole = res.role
      roleEl.innerHTML = roleLabel(res.role)
    } catch (err) {
      roleEl.innerHTML = `<span class="error-text">${escapeHtml(String(err.message || err))}</span>`
    }
  })

  $('cfg-save').addEventListener('click', () => {
    saveConfig({
      ...config,
      gasUrl: $('cfg-gas').value.trim(),
      accessToken: $('cfg-token').value.trim(),
      code: $('cfg-code').value.trim(),
      role: verifiedRole,
    })
    saveSettings({ ...settings, connections: draft, activeConnectionId: activeId, activeModel })
    root.innerHTML = ''
    onSaved?.()
  })
}

/** 汎用の入力ダイアログ。タグ追加や手編集のテキストエリアに使う */
export function openEditor({ title, value = '', multiline = true, hint = '', onSave }) {
  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div class="overlay">
      <div class="modal">
        <h2 class="modal-title">${escapeHtml(title)}</h2>
        ${hint ? `<p class="foot-note">${escapeHtml(hint)}</p>` : ''}
        ${multiline
          ? `<textarea id="ed-value" class="input" rows="14">${escapeHtml(value)}</textarea>`
          : `<input id="ed-value" class="input" value="${escapeHtml(value)}" />`}
        <div class="modal-foot">
          <button id="ed-cancel" class="btn">キャンセル</button>
          <button id="ed-save" class="btn btn-primary">保存</button>
        </div>
      </div>
    </div>
  `
  const input = document.getElementById('ed-value')
  input.focus()
  document.getElementById('ed-cancel').addEventListener('click', () => (root.innerHTML = ''))
  document.getElementById('ed-save').addEventListener('click', async () => {
    const next = input.value
    root.innerHTML = ''
    await onSave?.(next)
  })
}
