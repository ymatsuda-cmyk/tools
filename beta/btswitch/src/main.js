/**
 * 画面。GASの「いまどの端末が使うか」を読み書きするだけで、
 * 実際のBluetooth操作は各PCの常駐エージェント(agent/)が行う。
 */
const KEY = 'btswitch:config'
const POLL_MS = 5000
// エージェントの生存とみなす猶予。ハートビート間隔(既定10秒)の3倍
const ALIVE_MS = 35000

const $ = (id) => document.getElementById(id)
let timer = null

function loadConfig() {
  try {
    const raw = localStorage.getItem(KEY)
    return { gasUrl: '', token: '', me: '', ...(raw ? JSON.parse(raw) : {}) }
  } catch {
    return { gasUrl: '', token: '', me: '' }
  }
}

function saveConfig(c) {
  localStorage.setItem(KEY, JSON.stringify(c))
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
}

async function callGas(action, params = {}) {
  const config = loadConfig()
  if (!config.gasUrl || !config.token) throw new Error('設定からGASのURLと合言葉を入力してください')
  // text/plain にしないと CORS の preflight が飛び、GAS は応答できない
  const res = await fetch(config.gasUrl + (config.gasUrl.includes('?') ? '&' : '?') + 'r=' + Date.now().toString(36), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token: config.token, ...params }),
    redirect: 'follow',
  })
  if (!res.ok) throw new Error(`GAS HTTP ${res.status}`)
  const json = await res.json()
  if (!json.ok) throw new Error(json.error || 'GASがエラーを返しました')
  return json.data
}

function fmtTime(iso) {
  if (!iso) return ''
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  const p = (n) => String(n).padStart(2, '0')
  return `${p(d.getHours())}:${p(d.getMinutes())}`
}

function paint(state) {
  const config = loadConfig()
  const devices = Object.values(state.devices || {}).sort((a, b) => a.label.localeCompare(b.label))
  const now = Date.now()

  $('devices').innerHTML = [
    ...devices.map((d) => {
      const alive = now - new Date(d.lastSeenAt || 0).getTime() < ALIVE_MS
      const owner = state.owner === d.id
      return `
        <button class="device ${owner ? 'owner' : ''}" data-id="${escapeHtml(d.id)}">
          <i class="ti ti-device-laptop" aria-hidden="true"></i>
          <span class="device-label">${escapeHtml(d.label)}${d.id === config.me ? '（この端末）' : ''}</span>
          <span class="device-state">
            ${owner ? '<span class="pill on">使用中</span>' : ''}
            <span class="pill ${d.connected ? 'on' : ''}">${d.connected ? '接続あり' : '未接続'}</span>
            <span class="pill ${alive ? '' : 'warn'}">${alive ? `常駐OK ${fmtTime(d.lastSeenAt)}` : '常駐が止まっています'}</span>
          </span>
          ${d.note ? `<span class="device-note">${escapeHtml(d.note)}</span>` : ''}
        </button>`
    }),
    `<button class="device ${state.owner === 'phone' ? 'owner' : ''}" data-id="phone">
       <i class="ti ti-device-mobile" aria-hidden="true"></i>
       <span class="device-label">スマホ</span>
       <span class="device-state">${state.owner === 'phone' ? '<span class="pill on">使用中</span>' : ''}<span class="pill">PCが切断された状態</span></span>
     </button>`,
  ].join('')

  $('devices').querySelectorAll('.device').forEach((el) => {
    el.addEventListener('click', () => claim(el.dataset.id))
  })

  $('history').innerHTML = (state.history || []).map((h) => {
    const label = h.owner === 'phone' ? 'スマホ' : (state.devices?.[h.owner]?.label || h.owner || 'すべて切断')
    return `<li><span class="hist-time">${fmtTime(h.at)}</span>${escapeHtml(label)}<span class="hist-by">${escapeHtml(h.by || '')}</span></li>`
  }).join('') || '<li class="muted">まだありません</li>'

  $('status').textContent = state.owner
    ? `いまは「${state.owner === 'phone' ? 'スマホ' : (state.devices?.[state.owner]?.label || state.owner)}」`
    : 'どの端末も掴んでいません'
}

async function refresh() {
  try {
    paint(await callGas('getState'))
  } catch (err) {
    $('status').textContent = String(err.message || err)
  }
}

async function claim(deviceId) {
  $('status').textContent = '切り替えています...'
  try {
    paint(await callGas('claim', { deviceId, by: loadConfig().me || 'browser' }))
  } catch (err) {
    $('status').textContent = String(err.message || err)
  }
}

function openSettings() {
  const config = loadConfig()
  const root = $('modal-root')
  root.innerHTML = `
    <div class="modal-overlay">
      <div class="modal">
        <div class="modal-head"><span>設定</span><button class="btn-ghost btn-close" aria-label="閉じる"><i class="ti ti-x"></i></button></div>
        <div class="modal-body">
          <label>GAS ウェブアプリURL</label>
          <input id="cfg-url" class="input" value="${escapeHtml(config.gasUrl)}" placeholder="https://script.google.com/macros/s/.../exec" />
          <label>合言葉(GASの ACCESS_TOKEN)</label>
          <input id="cfg-token" class="input" value="${escapeHtml(config.token)}" />
          <label>この端末のID(エージェントの DEVICE_ID と同じ。スマホから見るだけなら空でよい)</label>
          <input id="cfg-me" class="input" value="${escapeHtml(config.me)}" placeholder="例: pc-home" />
        </div>
        <div class="modal-foot">
          <button class="btn btn-cancel">キャンセル</button>
          <button class="btn btn-primary btn-save">保存</button>
        </div>
      </div>
    </div>`
  const close = () => (root.innerHTML = '')
  root.querySelector('.btn-close').addEventListener('click', close)
  root.querySelector('.btn-cancel').addEventListener('click', close)
  root.querySelector('.btn-save').addEventListener('click', () => {
    saveConfig({
      gasUrl: root.querySelector('#cfg-url').value.trim(),
      token: root.querySelector('#cfg-token').value.trim(),
      me: root.querySelector('#cfg-me').value.trim(),
    })
    close()
    refresh()
  })
}

$('reload').addEventListener('click', refresh)
$('open-settings').addEventListener('click', openSettings)
$('release').addEventListener('click', () => claim('phone'))

// 他の端末から切り替えられたときも追いつけるよう、開いている間は見に行く
timer = setInterval(refresh, POLL_MS)
document.addEventListener('visibilitychange', () => {
  clearInterval(timer)
  if (!document.hidden) {
    refresh()
    timer = setInterval(refresh, POLL_MS)
  }
})

if (!loadConfig().gasUrl) openSettings()
refresh()
