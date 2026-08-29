import { renderList, renderDetailHtml } from './ui/render.js'
import { fetchSummary, fetchTranscript, saveSummary } from './lib/gas.js'
import { getDetailCache, setDetailCache, isCacheFresh } from './lib/cache.js'
import { generateSummary } from './lib/summarize.js'
import { loadConfig, saveConfig, isConfigured } from './lib/minutes-config.js'
import { loadSettings, saveSettings, newProfile } from './lib/llm-settings.js'

const listEl = document.getElementById('list')
const detailEl = document.getElementById('detail')
const syncStatusEl = document.getElementById('sync-status')

let items = []
let selectedKey = null

const MOBILE_BREAKPOINT = 720

function isMobile() {
  return window.innerWidth <= MOBILE_BREAKPOINT
}

async function loadIndex() {
  syncStatusEl.textContent = '読み込み中...'
  try {
    const res = await fetch('./data/index.json', { cache: 'no-store' })
    items = await res.json()
    syncStatusEl.textContent = `${items.length}件`
  } catch (err) {
    syncStatusEl.textContent = 'index.json の読み込みに失敗しました'
    items = []
  }
  renderList(listEl, items, selectedKey, onSelect)
}

/**
 * 詳細を表示するターゲット要素を決める。
 * モバイルは選択した行の直後にインライン挿入、PCは右ペインに固定表示。
 */
function detailTarget(rowEl) {
  if (!isMobile()) {
    detailEl.classList.add('side-panel')
    return detailEl
  }
  // 既存のインライン展開を除去してから作り直す
  document.querySelectorAll('.inline-detail').forEach((el) => el.remove())
  const wrap = document.createElement('div')
  wrap.className = 'inline-detail'
  wrap.innerHTML = '<div class="detail-pane-inner"></div>'
  rowEl.after(wrap)
  return wrap.querySelector('.detail-pane-inner')
}

function paintDetail(target, item, state) {
  target.innerHTML = renderDetailHtml(item, state)
  const generateBtn = target.querySelector('.btn-generate, .btn-regenerate')
  generateBtn?.addEventListener('click', () => runGenerate(target, item))
  target.querySelector('.btn-retry')?.addEventListener('click', () => onSelect(item, findRow(item.key)))
}

function findRow(key) {
  return listEl.querySelector(`.list-item[data-key="${key}"]`)
}

async function onSelect(item, rowEl) {
  selectedKey = item.key
  renderList(listEl, items, selectedKey, onSelect)
  const target = detailTarget(isMobile() ? findRow(item.key) : rowEl)

  paintDetail(target, item, { phase: 'loading' })

  const cache = getDetailCache(item.key)
  try {
    const remote = await fetchSummary(item.notionPageId)

    if (!remote.generatedAt) {
      paintDetail(target, item, { phase: 'no-summary' })
      return
    }

    const fresh = isCacheFresh(cache, remote.generatedAt) ? cache : setDetailCache(item.key, remote)
    paintDetail(target, item, { phase: 'ready', summary: fresh })
  } catch (err) {
    // 通信に失敗してもキャッシュがあればそれを出す
    if (cache) {
      paintDetail(target, item, { phase: 'ready', summary: cache })
    } else {
      paintDetail(target, item, { phase: 'error', message: String(err.message || err) })
    }
  }
}

async function runGenerate(target, item) {
  paintDetail(target, item, { phase: 'generating' })
  try {
    const { text } = await fetchTranscript(item.notionPageId)
    const result = await generateSummary(text, (partial) => {
      paintDetail(target, item, { phase: 'generating', progress: partial.slice(0, 200) })
    })
    await saveSummary(item.notionPageId, result.cardSummary, {
      decisions: result.decisions,
      todos: result.todos,
      topics: result.topics,
    }, result.model)

    const saved = setDetailCache(item.key, {
      cardSummary: result.cardSummary,
      detail: { decisions: result.decisions, todos: result.todos, topics: result.topics },
      model: result.model,
      generatedAt: new Date().toISOString(),
    })
    paintDetail(target, item, { phase: 'ready', summary: saved })
  } catch (err) {
    paintDetail(target, item, { phase: 'error', message: String(err.message || err) })
  }
}

// --- 設定モーダル(簡易版) ---
document.getElementById('open-settings').addEventListener('click', openSettings)

function openSettings() {
  const config = loadConfig()
  const settings = loadSettings()
  const profile = settings.profiles[0] || newProfile()

  const root = document.getElementById('modal-root')
  root.innerHTML = `
    <div style="position:fixed;inset:0;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center;z-index:10">
      <div style="background:var(--surface-2);border-radius:12px;padding:20px;width:320px;max-width:90vw">
        <h2 style="font-size:15px;margin:0 0 12px">設定</h2>
        <label style="font-size:12px;color:var(--text-secondary)">GAS URL</label>
        <input id="cfg-gas" value="${config.gasUrl}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">共有トークン</label>
        <input id="cfg-token" value="${config.notionToken}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">LLM baseUrl</label>
        <input id="cfg-llm-url" value="${profile.baseUrl}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">LLM APIキー</label>
        <input id="cfg-llm-key" value="${profile.apiKey}" style="width:100%;margin-bottom:8px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <label style="font-size:12px;color:var(--text-secondary)">モデル名</label>
        <input id="cfg-llm-model" value="${profile.model}" style="width:100%;margin-bottom:16px;padding:6px;border:0.5px solid var(--border);border-radius:6px" />
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="cfg-cancel" class="btn">キャンセル</button>
          <button id="cfg-save" class="btn">保存</button>
        </div>
      </div>
    </div>
  `
  document.getElementById('cfg-cancel').addEventListener('click', () => (root.innerHTML = ''))
  document.getElementById('cfg-save').addEventListener('click', () => {
    saveConfig({
      ...config,
      gasUrl: document.getElementById('cfg-gas').value.trim(),
      notionToken: document.getElementById('cfg-token').value.trim(),
    })

    const p = { ...profile,
      baseUrl: document.getElementById('cfg-llm-url').value.trim(),
      apiKey: document.getElementById('cfg-llm-key').value.trim(),
      model: document.getElementById('cfg-llm-model').value.trim(),
    }
    const nextProfiles = settings.profiles.length ? [p, ...settings.profiles.slice(1)] : [p]
    saveSettings({ ...settings, profiles: nextProfiles, activeId: p.id })

    root.innerHTML = ''
  })
}

loadIndex()
