import { loadConfig } from './minutes-config.js'

/**
 * GAS の doPost を呼ぶ。
 * Content-Type は必ず text/plain にすること — application/json にすると
 * ブラウザが CORS preflight (OPTIONS) を送るが、GAS は OPTIONS に応答できず
 * 常に失敗する。
 */
async function callGas(action, params = {}) {
  const config = loadConfig()
  if (!config.gasUrl || !config.notionToken) {
    throw new Error('GAS の接続設定が未入力です')
  }

  const res = await fetch(config.gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token: config.notionToken, ...params }),
  })

  if (!res.ok) {
    throw new Error(`GAS HTTP ${res.status}`)
  }
  const json = await res.json()
  if (!json.ok) {
    throw new Error(json.error || 'GAS がエラーを返しました')
  }
  return json.data
}

/** @returns {Promise<{text: string, updatedAt: string}>} */
export function fetchTranscript(pageId) {
  return callGas('fetchTranscript', { pageId })
}

/** @returns {Promise<{cardSummary: string|null, detail: object|null, model: string|null, generatedAt: string|null, updatedAt: string}>} */
export function fetchSummary(pageId) {
  return callGas('fetchSummary', { pageId })
}

/** @returns {Promise<{saved: true}>} */
export function saveSummary(pageId, cardSummary, detail, model) {
  return callGas('saveSummary', { pageId, cardSummary, detail, model })
}

/** @returns {Promise<{saved: true, tags: string[]}>} */
export function saveTags(pageId, tags) {
  return callGas('saveTags', { pageId, tags })
}

/** @returns {Promise<{saved: true, title: string}>} */
export function saveTitle(pageId, title) {
  return callGas('saveTitle', { pageId, title })
}

/** 状態を「再取得」にし、次回バッチでの文字起こしやり直しをリクエストする */
export function requestRetranscribe(pageId) {
  return callGas('requestRetranscribe', { pageId })
}

/** 人手編集の保存。要約日時・モデル・状態は変更されない */
export function saveDetail(pageId, cardSummary, detail) {
  return callGas('saveDetail', { pageId, cardSummary, detail })
}