// src/text-register.js — テキストから議事録を新規登録する
//
// タイトル・タグ・開始日時・会議時間(時分秒)・原文を入力し、GAS の
// createFromText で Notion 議事録DBにページを作る。登録先は既存カラム:
//   タイトル → ミーティング名 / タグ → カテゴリー / 開始日時 → 日時
//   会議時間 → 会議時間(「1時間 12分 36秒」形式) / 原文 → ページ本文
// 状態は「文字起こし」にし、要約は作らない。
//
// upload.js と同じく main.js の内部には依存せず、localStorage の
// "minutes:config"(GAS URL・トークン・検証済み role)をそのまま使う。
// 一覧は index.json から描画されるため、登録直後は upload.js の仮カードを
// 「一覧への反映待ち」として出し、index.json 反映後に自動で消す。

import { addPending } from './upload.js'
import { allKnownTags, filterByPermission, excludeDeleted } from './lib/filters.js'

const MAIN_CFG_KEY = 'minutes:config'
const INDEX_URL = 'https://ymatsuda-cmyk.github.io/tools/data/minutes/index.json'
const MAX_SUGGESTIONS = 12

function loadMainConfig() {
  try { return JSON.parse(localStorage.getItem(MAIN_CFG_KEY) || '{}') } catch { return {} }
}

async function gasCall(gasUrl, payload) {
  const res = await fetch(gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // preflight 回避(gas.js と同じ理由)
    body: JSON.stringify(payload),
  })
  const data = await res.json().catch(() => null)
  if (!res.ok || !data || data.ok === false) {
    throw new Error((data && data.error) || `GAS呼び出しに失敗 (${res.status})`)
  }
  return data.data
}

function pad(n) { return String(n).padStart(2, '0') }
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ))
}

/** GAS の formatDuration_ と同じ書式。0時間のときは「12分 36秒」 */
export function formatDuration(totalSec) {
  const sec = Math.max(0, Math.floor(Number(totalSec) || 0))
  const h = Math.floor(sec / 3600)
  const m = Math.floor((sec % 3600) / 60)
  const s = sec % 60
  return `${h > 0 ? `${h}時間 ` : ''}${m}分 ${s}秒`
}

function nowLocalValue() {
  const d = new Date()
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`
}

/** この利用者が見てよい議事録に付いているタグを、使用回数の多い順で返す */
async function loadTagSuggestions(role) {
  try {
    const res = await fetch(INDEX_URL, { cache: 'no-store' })
    if (!res.ok) return []
    const items = filterByPermission(excludeDeleted(await res.json()), role)
    const count = new Map()
    items.forEach((i) => (i.tags || []).forEach((t) => count.set(t, (count.get(t) || 0) + 1)))
    return allKnownTags(items).sort((a, b) => (count.get(b) || 0) - (count.get(a) || 0))
  } catch {
    return []
  }
}

// ---------------------------------------------------------------- modal

function openModal() {
  const root = document.getElementById('modal-root')
  if (!root) return

  const overlay = document.createElement('div')
  overlay.className = 'raw-modal-overlay'
  overlay.innerHTML = `
    <div class="raw-modal" style="width:min(560px,100%);max-height:90vh">
      <div class="raw-modal-header">
        <span>テキストから議事録を登録</span>
        <button class="btn-ghost" id="tr-close" aria-label="閉じる"><i class="ti ti-x" aria-hidden="true"></i></button>
      </div>
      <div class="raw-modal-body">
        <div class="upload-field">
          <label for="tr-title"><span>タイトル</span><span class="tr-col">ミーティング名</span></label>
          <input type="text" id="tr-title" placeholder="GMO自動化テスト 定例">
        </div>
        <div class="upload-field">
          <label for="tr-tag"><span>タグ</span><span class="tr-col">カテゴリー</span></label>
          <div class="tr-chips" id="tr-chips"></div>
          <input type="text" id="tr-tag" placeholder="入力して Enter で追加" list="tr-tag-list">
          <datalist id="tr-tag-list"></datalist>
          <div class="tr-sugs" id="tr-sugs"></div>
        </div>
        <div class="tr-row">
          <div class="upload-field">
            <label for="tr-at"><span>開始日時</span><span class="tr-col">日時</span></label>
            <input type="datetime-local" id="tr-at">
          </div>
          <div class="upload-field">
            <label><span>会議時間</span><span class="tr-col">会議時間</span></label>
            <div class="tr-dur">
              <input type="number" id="tr-h" min="0" max="99" value="0" aria-label="時間"><span>時間</span>
              <input type="number" id="tr-m" min="0" max="59" value="0" aria-label="分"><span>分</span>
              <input type="number" id="tr-s" min="0" max="59" value="0" aria-label="秒"><span>秒</span>
            </div>
            <div class="tr-hint" id="tr-dur-preview"></div>
          </div>
        </div>
        <div class="upload-field">
          <label for="tr-raw"><span>原文</span><span class="tr-col">ページ本文</span></label>
          <textarea id="tr-raw" class="tr-raw" rows="9" placeholder="[00:00:05] 松田：では定例を始めます"></textarea>
          <div class="tr-hint tr-raw-foot"><span id="tr-count">0 文字</span><span>.txt をドロップしても読み込めます</span></div>
        </div>
        <div class="tr-note">状態は「文字起こし」で登録します。要約は作りません。</div>
        <div class="upload-msg" id="tr-msg" role="status" aria-live="polite"></div>
      </div>
      <div class="raw-modal-footer">
        <button class="btn" id="tr-cancel">キャンセル</button>
        <button class="btn tr-primary" id="tr-go">Notionに登録</button>
      </div>
    </div>
  `
  root.appendChild(overlay)

  const $ = (s) => overlay.querySelector(s)
  const titleInput = $('#tr-title')
  const tagInput = $('#tr-tag')
  const atInput = $('#tr-at')
  const hIn = $('#tr-h'), mIn = $('#tr-m'), sIn = $('#tr-s')
  const raw = $('#tr-raw')
  const msg = $('#tr-msg')
  const go = $('#tr-go')

  let tags = []
  let suggestions = []
  let busy = false

  atInput.value = nowLocalValue()

  function say(text, kind = '') {
    msg.textContent = text
    msg.className = `upload-msg ${kind}`
  }
  // 入力途中の誤クローズを防ぐため、閉じる手段はキャンセルと×だけ(upload.js と同じ)
  function close() { if (!busy) overlay.remove() }
  $('#tr-close').onclick = close
  $('#tr-cancel').onclick = close

  // ---- タグ
  function drawTags() {
    $('#tr-chips').innerHTML = tags.map((t, i) => `
      <span class="tag-chip">${escapeHtml(t)}<span class="tag-remove" data-i="${i}" role="button" tabindex="0" aria-label="${escapeHtml(t)}を外す">×</span></span>
    `).join('')
    $('#tr-chips').querySelectorAll('.tag-remove').forEach((el) => {
      const remove = () => { tags.splice(Number(el.dataset.i), 1); drawTags() }
      el.onclick = remove
      el.onkeydown = (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); remove() } }
    })
    const rest = suggestions.filter((t) => !tags.includes(t))
    $('#tr-sugs').innerHTML = rest.length
      ? `<span class="tr-hint">既存:</span>${rest.slice(0, MAX_SUGGESTIONS).map((t) => `<button type="button" class="tr-sug">${escapeHtml(t)}</button>`).join('')}`
      : ''
    $('#tr-sugs').querySelectorAll('.tr-sug').forEach((b) => { b.onclick = () => addTag(b.textContent) })
    $('#tr-tag-list').innerHTML = rest.map((t) => `<option value="${escapeHtml(t)}"></option>`).join('')
  }
  function addTag(v) {
    const t = String(v || '').trim()
    if (t && !tags.includes(t)) { tags.push(t); drawTags() }
  }
  tagInput.onkeydown = (e) => {
    if (e.key === 'Enter' && !e.isComposing) {
      e.preventDefault()
      addTag(tagInput.value)
      tagInput.value = ''
    }
  }
  // datalist から選んだ場合は Enter を待たずに追加する
  tagInput.oninput = () => {
    if (suggestions.includes(tagInput.value)) { addTag(tagInput.value); tagInput.value = '' }
  }
  drawTags()
  loadTagSuggestions(loadMainConfig().role).then((list) => { suggestions = list; drawTags() })

  // ---- 会議時間
  function durationSec() {
    const n = (el) => Math.max(0, Math.floor(Number(el.value) || 0))
    return n(hIn) * 3600 + n(mIn) * 60 + n(sIn)
  }
  function paintDuration() {
    const sec = durationSec()
    let text = `保存値「${formatDuration(sec)}」`
    if (atInput.value && sec > 0) {
      const end = new Date(new Date(atInput.value).getTime() + sec * 1000)
      text += ` ／ 終了 ${pad(end.getHours())}:${pad(end.getMinutes())}`
    }
    $('#tr-dur-preview').textContent = text
  }
  ;[atInput, hIn, mIn, sIn].forEach((el) => el.addEventListener('input', paintDuration))
  paintDuration()

  // ---- 原文
  function paintCount() { $('#tr-count').textContent = `${raw.value.length.toLocaleString()} 文字` }
  raw.addEventListener('input', paintCount)
  raw.ondragover = (e) => { e.preventDefault(); raw.classList.add('drag-over') }
  raw.ondragleave = () => raw.classList.remove('drag-over')
  raw.ondrop = async (e) => {
    const file = e.dataTransfer.files[0]
    if (!file) return
    e.preventDefault()
    raw.classList.remove('drag-over')
    if (!/\.(txt|md|vtt|srt)$/i.test(file.name) && !file.type.startsWith('text/')) {
      return say('テキストファイル(.txt など)をドロップしてください。', 'error')
    }
    raw.value = await file.text()
    if (!titleInput.value.trim()) titleInput.value = file.name.replace(/\.[^.]+$/, '')
    paintCount()
    say('')
  }

  // ---- 登録
  go.onclick = async () => {
    const cfg = loadMainConfig()
    if (!cfg.gasUrl || !cfg.notionToken) {
      return say('設定画面(歯車アイコン)で GAS の URL とコードを先に登録してください。', 'error')
    }
    const title = titleInput.value.trim()
    if (!title) { titleInput.focus(); return say('タイトルを入力してください。', 'error') }
    if (!atInput.value) { atInput.focus(); return say('開始日時を入れてください。', 'error') }
    if (!raw.value.trim()) { raw.focus(); return say('原文を貼り付けてください。', 'error') }

    // Enter を押さずに残っているタグ入力も拾う
    if (tagInput.value.trim()) { addTag(tagInput.value); tagInput.value = '' }

    const meetingAt = `${atInput.value}:00+09:00`
    busy = true
    go.disabled = true
    say('Notionに登録しています…')
    try {
      await gasCall(cfg.gasUrl, {
        token: cfg.notionToken,
        action: 'createFromText',
        title,
        tags,
        meetingAt,
        durationSec: durationSec(),
        text: raw.value,
        role: cfg.role || '',
      })
      addPending({ id: `t${Date.now()}`, title, meetingAt, status: 'registered', createdAt: Date.now() })
      say('登録しました。一覧には index.json の更新後に反映されます。', 'ok')
      busy = false
      setTimeout(close, 900)
    } catch (err) {
      console.error(err)
      busy = false
      go.disabled = false
      say(err.message, 'error')
    }
  }

  titleInput.focus()
}

// ---------------------------------------------------------------- init

export function initTextRegister() {
  const btn = document.getElementById('open-text-register')
  if (btn) btn.onclick = openModal
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initTextRegister)
} else {
  initTextRegister()
}
