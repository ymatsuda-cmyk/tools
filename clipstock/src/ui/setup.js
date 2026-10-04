import { escapeHtml } from './render.js'
import { loadConfig, saveConfig, connectJsonbin, saveJsonbinPassphrase } from '../lib/videos-config.js'
import { loadSettings, saveSettings, newConnection } from '../lib/llm-settings.js'
import { verifyCode } from '../lib/gas.js'
import { openSettings } from './settings.js'

/**
 * 初回設定画面。GAS URL と共有トークンが未設定のときに、起動直後に出す。
 *
 * 保存先を最初に選ぶ:
 *   JSONBin  … Bin ID・APIキー・パスフレーズで保存済みの設定を読み込む。
 *              以降の変更(設定画面での保存)もJSONBinに書く。
 *   この端末のみ … GAS URL・共有トークン・一覧JSONの場所・コード・AI接続を入力し、
 *              これまでどおりlocalStorageだけに保存する。JSONBinには何も送らない。
 *
 * 設定画面(settings.js)と違い、保存を押すまで何も書かない。読み込みに失敗したときも
 * この端末の設定は変えない。
 *
 * @param {() => void} onDone 設定が済んだあとに呼ぶ(一覧の読み込みを始める)
 */
export function openSetup(onDone) {
  const config = loadConfig()
  let mode = config.useJsonbin ? 'jsonbin' : 'local'
  let verifiedRole = ''

  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div class="overlay">
      <div class="modal modal-sticky">
        <h2 class="modal-title">はじめの設定</h2>
        <div class="modal-body">
          <div class="foot-note">設定の保存先を選んでください。あとから設定画面(右上の歯車)でも変えられます</div>

          <label class="row">
            <input type="radio" name="su-mode" value="jsonbin" ${mode === 'jsonbin' ? 'checked' : ''} />
            <span>JSONBinを使う(保存してある設定を読み込み、変更もJSONBinに書く)</span>
          </label>
          <label class="row">
            <input type="radio" name="su-mode" value="local" ${mode === 'local' ? 'checked' : ''} />
            <span>JSONBinを使わない(この端末のブラウザにだけ保存する)</span>
          </label>

          <div id="su-jsonbin">
            <label class="field-label">Bin ID</label>
            <input id="su-bin-id" class="input" value="${escapeHtml(config.jsonbinBinId)}" placeholder="jsonbin.ioのBinのID" />
            <label class="field-label">APIキー(X-Master-Key)</label>
            <input id="su-bin-key" class="input" type="password" value="${escapeHtml(config.jsonbinApiKey)}" placeholder="jsonbin.ioで発行したキー" />
            <label class="field-label">パスフレーズ(暗号化)</label>
            <input id="su-bin-pass" class="input" type="password" autocomplete="off" placeholder="Binを暗号化したときのパスフレーズ" />
            <div class="foot-note">パスフレーズはこの端末のブラウザにだけ保存します。Binの中身は設定画面と同じ形式(AES-256-GCM)で読み書きします</div>
          </div>

          <div id="su-local">
            <label class="field-label">GAS URL</label>
            <input id="su-gas" class="input" value="${escapeHtml(config.gasUrl)}" placeholder="https://script.google.com/macros/s/.../exec" />

            <label class="field-label">共有トークン</label>
            <input id="su-token" class="input" value="${escapeHtml(config.accessToken)}" placeholder="GASの ACCESS_TOKEN と同じ値" />

            <label class="field-label">一覧JSONの場所</label>
            <input id="su-data" class="input" value="${escapeHtml(config.dataUrl)}" placeholder="空欄で data/clipstock/ を使う" />

            <label class="field-label">コード</label>
            <div class="row">
              <input id="su-code" class="input grow" value="${escapeHtml(config.code)}" />
              <button id="su-verify" class="btn">確認</button>
            </div>
            <div id="su-role" class="foot-note">使わなければ空欄のままで構いません</div>

            <label class="field-label">AI接続(あとから設定画面でも追加できます)</label>
            <input id="su-ai-label" class="input sm" placeholder="表示名(例: Gemini)" />
            <input id="su-ai-base" class="input sm" placeholder="baseUrl (例: https://generativelanguage.googleapis.com/v1beta/openai)" />
            <input id="su-ai-key" class="input sm" type="password" placeholder="APIキー" />
            <input id="su-ai-model" class="input sm" placeholder="モデル名(例: gemini-2.5-flash)" />
          </div>

          <div id="su-msg" class="foot-note"></div>
        </div>

        <div class="modal-foot">
          <button id="su-later" class="btn">あとで</button>
          <button id="su-go" class="btn btn-primary"></button>
        </div>
      </div>
    </div>
  `

  const $ = (id) => document.getElementById(id)
  const msg = (html) => ($('su-msg').innerHTML = html)
  const err = (e) => msg(`<span class="error-text">${escapeHtml(String(e?.message || e))}</span>`)

  function paintMode() {
    $('su-jsonbin').hidden = mode !== 'jsonbin'
    $('su-local').hidden = mode !== 'local'
    $('su-go').textContent = mode === 'jsonbin' ? '読み込む' : '保存して始める'
    msg('')
  }
  paintMode()

  root.querySelectorAll('input[name="su-mode"]').forEach((r) =>
    r.addEventListener('change', () => {
      mode = r.value
      paintMode()
    })
  )

  // 閉じても次回の起動でまた出る(未設定のままなので)。一覧側にも案内が出る
  $('su-later').addEventListener('click', () => {
    root.innerHTML = ''
    onDone?.()
  })

  $('su-verify').addEventListener('click', async () => {
    const gasUrl = $('su-gas').value.trim()
    if (!gasUrl) return ($('su-role').innerHTML = '<span class="error-text">GAS URLを先に入力してください</span>')
    $('su-role').textContent = '確認中...'
    try {
      const res = await verifyCode(gasUrl, $('su-code').value.trim())
      verifiedRole = res.role
      $('su-role').innerHTML =
        res.role === 'err'
          ? '<span class="error-text">権限がありません</span>'
          : res.role === 'xYz'
            ? '<span class="ok-text">管理者 — 全機能が使えます</span>'
            : `<span class="ok-text">権限: ${escapeHtml(res.role)}</span>`
    } catch (e) {
      $('su-role').innerHTML = `<span class="error-text">${escapeHtml(String(e?.message || e))}</span>`
    }
  })

  $('su-go').addEventListener('click', async () => {
    const btn = $('su-go')
    btn.disabled = true
    try {
      if (mode === 'jsonbin') await goJsonbin()
      else goLocal()
    } catch (e) {
      err(e)
    } finally {
      btn.disabled = false
    }
  })

  /** JSONBinから読み込む。読めなければ(キー違い・パスフレーズ違い等)何も書き換えずにここへ戻る */
  async function goJsonbin() {
    const binId = $('su-bin-id').value.trim()
    const apiKey = $('su-bin-key').value.trim()
    const passphrase = $('su-bin-pass').value.trim()
    if (!binId || !apiKey) throw new Error('Bin IDとAPIキーを入力してください')
    if (!passphrase) throw new Error('パスフレーズを入力してください(JSONBinの中身は暗号化されています)')
    msg('読み込み中...')
    const merged = await connectJsonbin({ binId, apiKey, passphrase })
    root.innerHTML = ''
    // Binは読めたが GAS URL・共有トークンが入っていなかったときは、設定画面で足してもらう。
    // 設定画面の保存はJSONBinにも書くので、足した内容がそのままBinに反映される
    if (!merged.gasUrl || !merged.accessToken) {
      openSettings(() => onDone?.())
      return
    }
    onDone?.()
  }

  /** この端末のlocalStorageだけに保存する。JSONBinへは何も送らない */
  function goLocal() {
    const gasUrl = $('su-gas').value.trim()
    const accessToken = $('su-token').value.trim()
    if (!gasUrl || !accessToken) throw new Error('GAS URLと共有トークンは必須です')

    saveConfig({
      ...config,
      gasUrl,
      accessToken,
      dataUrl: $('su-data').value.trim(),
      code: $('su-code').value.trim(),
      role: verifiedRole || config.role,
      useJsonbin: false,
    })
    saveJsonbinPassphrase('') // 使わない選択なので、以前のパスフレーズも残さない

    // AI接続は表示名以外が揃っているときだけ登録する(空のままなら今ある接続に触らない)
    const base = $('su-ai-base').value.trim()
    const model = $('su-ai-model').value.trim()
    if (base && model) {
      const settings = loadSettings()
      const conn = newConnection({
        label: $('su-ai-label').value.trim() || '接続1',
        baseUrl: base,
        apiKey: $('su-ai-key').value.trim(),
        models: [model],
      })
      // 初期状態の空の接続(baseUrlなし・モデルなし)は置き換え、既存の接続は残す
      const kept = settings.connections.filter((c) => c.baseUrl || (c.models || []).length)
      saveSettings({ ...settings, connections: [...kept, conn], activeConnectionId: conn.id, activeModel: model })
    }

    root.innerHTML = ''
    onDone?.()
  }
}
