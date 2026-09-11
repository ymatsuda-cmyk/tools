import { loadConfig } from './videos-config.js'

/**
 * ページIDごとの取り込み元。
 * web記事DBは動画DBと別のNotion(統合も別のことがある)なので、
 * 書き戻すときにどちらのトークンを使うかをGAS側へ伝える必要がある。
 */
const pageSources = new Map()

/** 一覧を読み込んだら呼ぶ。以降そのページへの操作に source が付く */
export function registerPageSources(items) {
  ;(items || []).forEach((item) => {
    if (item && item.key) pageSources.set(item.key, item.source === 'web' ? 'web' : 'video')
  })
}

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

  const source = params.pageId ? pageSources.get(params.pageId) : null
  const body = JSON.stringify({ action, token: config.accessToken, ...params, ...(source ? { source } : {}) })

  const json = await postWithRetry(config.gasUrl, body)
  if (!json.ok) throw new Error(json.error || 'GAS がエラーを返しました')
  return json.data
}

async function post(url, body) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body,
  })
  if (!res.ok) throw new Error(`GAS HTTP ${res.status}`)
  return res.json()
}

/**
 * 通信そのものが失敗したときだけ一度やり直す。
 * GAS は再デプロイ直後などに 404 や "Failed to fetch"(リダイレクト先がCORSを返さない)を
 * 返すことがあり、Notion への保存は成功しているのに失敗表示になっていた。
 * どの action も同じ値を書き直すだけなので、投げ直しても副作用は増えない。
 * Notion 由来のエラーは ok:false で返ってくるため、ここでは再送しない。
 */
async function postWithRetry(url, body) {
  try {
    return await post(url, body)
  } catch {
    await new Promise((resolve) => setTimeout(resolve, 800))
    return post(url, body)
  }
}

/** 一覧を Notion から直接取得する(index-video.json のような中間ファイルは使わない) */
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

/** マインドマップ一覧に出すかどうか(Notionの「公開」チェックボックス) */
export function setPublic(pageId, isPublic) {
  return callGas('setPublic', { pageId, isPublic })
}

/** アイデア一覧に出すかどうか。kind は 'apply'(応用) か 'ideas'(活用) */
export function setIdeaPublic(pageId, kind, isPublic) {
  return callGas('setIdeaPublic', { pageId, kind, isPublic })
}

/** Notionページをゴミ箱へ移す。Notion側からなら30日間は復元できる */
export function deleteVideo(pageId) {
  return callGas('deleteVideo', { pageId })
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
