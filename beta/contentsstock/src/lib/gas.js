/**
 * GAS(doPost)の呼び出し口。
 *
 * Content-Type は必ず text/plain にすること — application/json にすると
 * ブラウザが CORS preflight (OPTIONS) を送るが、GAS は OPTIONS に応答できず必ず失敗する。
 */
import { loadConfig } from './contents-config.js'

async function callGas(action, params = {}) {
  const config = loadConfig()
  if (!config.gasUrl || !config.accessToken) {
    throw new Error('GAS の接続設定が未入力です')
  }
  const body = JSON.stringify({ action, token: config.accessToken, ...params })
  const json = await enqueue(() => postWithRetry(config.gasUrl, body))
  if (!json.ok) throw new Error(json.error || 'GAS がエラーを返しました')
  return json.data
}

/**
 * GAS へのリクエストは1本ずつ、しかも一定の間隔を空けて流す。
 * 立て続けに投げると Google 側がリダイレクト先を 404 で返すことがあり、
 * ブラウザからは CORS エラーとして見えてしまう。
 */
const MIN_INTERVAL_MS = 2500
let queueTail = Promise.resolve()
let lastSentAt = 0

function enqueue(task) {
  const run = queueTail.then(async () => {
    const wait = lastSentAt + MIN_INTERVAL_MS - Date.now()
    if (wait > 0) await sleep(wait)
    try {
      return await task()
    } finally {
      lastSentAt = Date.now()
    }
  })
  queueTail = run.catch(() => {})
  return run
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

async function post(url, body) {
  // 毎回URLを変える。同じURLだと壊れたリダイレクトを掴んだまま繰り返すことがある
  const target = url + (url.includes('?') ? '&' : '?') + 'r=' + Date.now().toString(36)
  const res = await fetch(target, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body,
    redirect: 'follow',
  })
  if (!res.ok) throw new Error(`GAS HTTP ${res.status}`)
  return res.json()
}

/**
 * 通信そのものが失敗したときだけやり直す。
 * どの action も同じ値を書き直すだけなので、投げ直しても副作用は増えない。
 * Notion 由来のエラーは ok:false で返るため、ここでは再送しない。
 */
async function postWithRetry(url, body) {
  const waits = [400, 1500, 4000, 10000, 20000]
  for (let i = 0; ; i++) {
    try {
      return await post(url, body)
    } catch (err) {
      if (i >= waits.length) {
        console.warn('[contentsstock] GAS通信に失敗しました(再送しきりました)', err.message || err)
        throw err
      }
      await sleep(waits[i])
    }
  }
}

export function listContents() {
  return callGas('listContents')
}

export function listIdeas() {
  return callGas('listIdeas')
}

/** 文字起こし全文(ページ本文) */
export function fetchTranscript(pageId) {
  return callGas('fetchTranscript', { pageId })
}

/** AI生成物・メモ・メタをまとめて取得 */
export function fetchDetail(pageId) {
  return callGas('fetchDetail', { pageId })
}

/** AI生成物を保存する。detail に入れたキーだけが更新される */
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

export function saveTitle(pageId, title) {
  return callGas('saveTitle', { pageId, title })
}

export function setStatus(pageId, status) {
  return callGas('setStatus', { pageId, status })
}

export function setPublic(pageId, isPublic) {
  return callGas('setPublic', { pageId, isPublic })
}

export function deleteContent(pageId) {
  return callGas('deleteContent', { pageId })
}

export function updateRawCount(pageId, count) {
  return callGas('updateRawCount', { pageId, count })
}

export function verifyCode(gasUrl, code) {
  return postWithRetry(gasUrl, JSON.stringify({ action: 'verifyCode', code })).then((json) => {
    if (!json.ok) throw new Error(json.error || 'GAS がエラーを返しました')
    return json.data
  })
}

// ---- 動画のアップロード(Drive の inbox へ) ----

/** アップロード先のセッションURLを発行してもらう */
export function initUpload(filename, mimeType, size) {
  return callGas('initUpload', { filename, mimeType, size })
}

/** base64にしたチャンクを中継してもらう。done:true で完了 */
export function putChunk(sessionUrl, chunk, offset, total) {
  return callGas('putChunk', { sessionUrl, chunk, offset, total })
}

/** 打合せ名などのメタデータを同名の .json として置く */
export function writeSidecar(name, meta) {
  return callGas('writeSidecar', { name, meta })
}
