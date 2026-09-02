import { loadConfig } from './videos-config.js'

/**
 * GAS の doPost を呼ぶ。
 * Content-Type は必ず text/plain にすること — application/json にすると
 * ブラウザが CORS preflight (OPTIONS) を送るが、GAS は OPTIONS に応答できず
 * 常に失敗する。
 */
async function callGas(action, params = {}) {
  const config = loadConfig()
  if (!config.gasUrl || !config.accessToken) {
    throw new Error('GAS の接続設定が未入力です')
  }

  const res = await fetch(config.gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, token: config.accessToken, ...params }),
  })

  if (!res.ok) throw new Error(`GAS HTTP ${res.status}`)
  const json = await res.json()
  if (!json.ok) throw new Error(json.error || 'GAS がエラーを返しました')
  return json.data
}

/** 一覧を Notion から直接取得する(index.json のような中間ファイルは使わない) */
export function listVideos() {
  return callGas('listVideos')
}

/** 応用・活用アイデアだけを全件取得する(アイデア一覧画面用) */
export function listIdeas() {
  return callGas('listIdeas')
}

/** @returns {Promise<{text: string, updatedAt: string}>} */
export function fetchTranscript(pageId) {
  return callGas('fetchTranscript', { pageId })
}

/** AI生成物・メモ・メタをまとめて取得 */
export function fetchDetail(pageId) {
  return callGas('fetchDetail', { pageId })
}

/**
 * AI生成物を保存する。detail に入れたキーだけが更新される。
 * キー: summary / mindmap / fields / apply / ideas / tags
 */
export function saveGenerated(pageId, detail, model, rawCount) {
  return callGas('saveGenerated', { pageId, detail, model, rawCount })
}

/** 人手編集の保存。要約日時・モデル・状態は変更されない */
export function saveField(pageId, field, value) {
  return callGas('saveField', { pageId, field, value })
}

export function saveMemo(pageId, memo) {
  return callGas('saveMemo', { pageId, memo })
}

export function saveTags(pageId, tags) {
  return callGas('saveTags', { pageId, tags })
}

/** タグを統合する。from が付いている全ページを to に置き換える */
export function mergeTag(from, to) {
  return callGas('mergeTag', { from, to })
}

export function saveTitle(pageId, title) {
  return callGas('saveTitle', { pageId, title })
}

/** 状態変更。'新規' に戻すと次回バッチで文字起こしをやり直す。'除外' は論理削除 */
export function setStatus(pageId, status) {
  return callGas('setStatus', { pageId, status })
}

export function updateRawCount(pageId, count) {
  return callGas('updateRawCount', { pageId, count })
}

/** 権限コードを検証する。共有トークンは不要(初回はまだ手元に無いため) */
export async function verifyCode(gasUrl, code) {
  const res = await fetch(gasUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action: 'verifyCode', code }),
  })
  if (!res.ok) throw new Error(`GAS HTTP ${res.status}`)
  const json = await res.json()
  if (!json.ok) throw new Error(json.error || 'GAS がエラーを返しました')
  return json.data
}
