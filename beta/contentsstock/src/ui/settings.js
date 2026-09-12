/** 設定モーダル(GAS接続・権限コード・AI接続とモデル) */
import { loadConfig, saveConfig } from '../lib/contents-config.js'
import { loadSettings, saveSettings, newConnection } from '../lib/llm-settings.js'
import { verifyCode } from '../lib/gas.js'
import { escapeHtml } from './render.js'

export function openSettings(onClose) {
  const root = document.getElementById('modal-root')
  const config = loadConfig()
  const settings = loadSettings()
  // 保存を押すまで localStorage を書き換えないよう、下書きに対して編集する
  const draft = { ...settings, connections: settings.connections.map((c) => ({ ...c, models: [...(c.models || [])] })) }

  const paint = () => {
    root.innerHTML = `
      <div class="modal-overlay">
        <div class="modal">
          <div class="modal-head"><span>設定</span><button class="btn-ghost btn-close" aria-label="閉じる"><i class="ti ti-x"></i></button></div>
          <div class="modal-body">
            <h4>データの接続</h4>
            <label>GAS ウェブアプリURL</label>
            <input id="cfg-gas" class="input" value="${escapeHtml(config.gasUrl)}" placeholder="https://script.google.com/macros/s/.../exec" />
            <label>共有トークン(GASの ACCESS_TOKEN)</label>
            <input id="cfg-token" class="input" value="${escapeHtml(config.accessToken)}" />
            <label>権限コード</label>
            <div class="row">
              <input id="cfg-code" class="input" value="${escapeHtml(config.code)}" />
              <button class="btn btn-verify">確認</button>
              <span id="cfg-role" class="foot-note">${config.role ? `権限: ${escapeHtml(config.role)}` : ''}</span>
            </div>

            <h4>AI接続</h4>
            ${draft.connections.map((c) => `
              <div class="conn" data-id="${c.id}">
                <div class="row">
                  <input class="input conn-label" data-id="${c.id}" value="${escapeHtml(c.label)}" placeholder="表示名(例: Gemini)" />
                  <button class="btn-ghost conn-del" data-id="${c.id}" aria-label="削除"><i class="ti ti-trash"></i></button>
                </div>
                <input class="input conn-url" data-id="${c.id}" value="${escapeHtml(c.baseUrl)}" placeholder="baseURL (例: https://generativelanguage.googleapis.com/v1beta/openai)" />
                <input class="input conn-key" data-id="${c.id}" value="${escapeHtml(c.apiKey)}" placeholder="APIキー" type="password" />
                <div class="chips">
                  ${(c.models || []).map((m) => `
                    <span class="chip ${c.id === draft.activeConnectionId && m === draft.activeModel ? 'on' : ''}" data-conn="${c.id}" data-model="${escapeHtml(m)}">
                      ${escapeHtml(m)}<i class="ti ti-x chip-del" data-conn="${c.id}" data-model="${escapeHtml(m)}"></i>
                    </span>`).join('') || '<span class="foot-note">モデル未登録</span>'}
                </div>
                <div class="row">
                  <input class="input conn-model-new" data-id="${c.id}" placeholder="モデル名を追加(例: gemini-3.6-flash)" />
                  <button class="btn conn-model-add" data-id="${c.id}">追加</button>
                </div>
              </div>
            `).join('')}
            <button class="btn btn-conn-add"><i class="ti ti-plus"></i>接続を追加</button>
          </div>
          <div class="modal-foot">
            <button class="btn btn-cancel">キャンセル</button>
            <button class="btn btn-primary btn-save">保存</button>
          </div>
        </div>
      </div>
    `
    wire()
  }

  const close = () => {
    root.innerHTML = ''
    onClose?.()
  }

  function wire() {
    root.querySelector('.btn-close').addEventListener('click', close)
    root.querySelector('.btn-cancel').addEventListener('click', close)

    root.querySelector('.btn-verify').addEventListener('click', async () => {
      const gasUrl = root.querySelector('#cfg-gas').value.trim()
      const code = root.querySelector('#cfg-code').value.trim()
      const el = root.querySelector('#cfg-role')
      el.textContent = '確認中...'
      try {
        const { role } = await verifyCode(gasUrl, code)
        config.role = role
        el.textContent = role === 'err' ? 'コードが違います' : `権限: ${role}`
      } catch (err) {
        el.textContent = '確認できませんでした: ' + (err.message || err)
      }
    })

    root.querySelectorAll('.conn-label').forEach((el) => el.addEventListener('input', () => set(el.dataset.id, 'label', el.value)))
    root.querySelectorAll('.conn-url').forEach((el) => el.addEventListener('input', () => set(el.dataset.id, 'baseUrl', el.value.trim())))
    root.querySelectorAll('.conn-key').forEach((el) => el.addEventListener('input', () => set(el.dataset.id, 'apiKey', el.value.trim())))

    root.querySelectorAll('.conn-model-add').forEach((btn) => btn.addEventListener('click', () => {
      const input = root.querySelector(`.conn-model-new[data-id="${btn.dataset.id}"]`)
      const name = input.value.trim()
      if (!name) return
      const conn = draft.connections.find((c) => c.id === btn.dataset.id)
      if (!conn.models.includes(name)) conn.models.push(name)
      draft.activeConnectionId = conn.id
      draft.activeModel = name
      paint()
    }))

    root.querySelectorAll('.chip').forEach((el) => el.addEventListener('click', (e) => {
      if (e.target.classList.contains('chip-del')) return
      draft.activeConnectionId = el.dataset.conn
      draft.activeModel = el.dataset.model
      paint()
    }))

    root.querySelectorAll('.chip-del').forEach((el) => el.addEventListener('click', () => {
      const conn = draft.connections.find((c) => c.id === el.dataset.conn)
      conn.models = conn.models.filter((m) => m !== el.dataset.model)
      paint()
    }))

    root.querySelectorAll('.conn-del').forEach((el) => el.addEventListener('click', () => {
      draft.connections = draft.connections.filter((c) => c.id !== el.dataset.id)
      if (!draft.connections.length) draft.connections = [newConnection()]
      paint()
    }))

    root.querySelector('.btn-conn-add').addEventListener('click', () => {
      draft.connections.push(newConnection())
      paint()
    })

    root.querySelector('.btn-save').addEventListener('click', () => {
      saveConfig({
        ...config,
        gasUrl: root.querySelector('#cfg-gas').value.trim(),
        accessToken: root.querySelector('#cfg-token').value.trim(),
        code: root.querySelector('#cfg-code').value.trim(),
      })
      saveSettings(draft)
      close()
    })
  }

  function set(id, key, value) {
    const conn = draft.connections.find((c) => c.id === id)
    if (conn) conn[key] = value
  }

  paint()
}
