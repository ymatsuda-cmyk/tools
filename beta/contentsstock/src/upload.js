/**
 * 動画のアップロード。
 *
 * ブラウザ -> GAS -> Google Drive の inbox フォルダ、という中継にしている。
 * Drive の resumable セッションURLへブラウザから直接PUTするとCORSで弾かれるため、
 * バイト列もGASを通す。1回の doPost は数MBのチャンクなので50MB制限には当たらない。
 *
 * 置いたあとの文字起こしは Mac 側の video_inbox.py が担当し、
 * Notion にページができたところで一覧に並ぶ。
 */
import { initUpload, putChunk, writeSidecar } from './lib/gas.js'
import { loadConfig, isConfigured } from './lib/contents-config.js'

const PUT_CHUNK = 3 * 1024 * 1024 // base64化すると4/3倍になるので小さめに保つ
const VIDEO_EXT = /\.(mp4|mov|m4v|webm|mkv|avi|mp3|m4a|wav|aac)$/i

function arrayBufferToBase64(buffer) {
  let binary = ''
  const bytes = new Uint8Array(buffer)
  const chunkSize = 0x8000
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize))
  }
  return btoa(binary)
}

function slugify(s) {
  return (s || 'untitled')
    .replace(/[\\/:*?"<>|\s]+/g, '_')
    .replace(/_{2,}/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 60)
}

function stamp(date = new Date()) {
  const p = (n) => String(n).padStart(2, '0')
  return `${date.getFullYear()}${p(date.getMonth() + 1)}${p(date.getDate())}-${p(date.getHours())}${p(date.getMinutes())}`
}

/** 1ファイルを inbox へ送り、タイトルなどを書いたサイドカーJSONを添える */
export async function uploadVideo(file, { title, tags = [], onProgress } = {}) {
  if (!isConfigured(loadConfig())) throw new Error('設定からGASのURLと共有トークンを入力してください')
  if (!VIDEO_EXT.test(file.name)) throw new Error('動画(mp4/mov/webm など)か音声ファイルを選んでください')

  const base = `${stamp()}_${slugify(title || file.name.replace(/\.[^.]+$/, ''))}`
  const ext = (file.name.match(/\.[^.]+$/) || ['.mp4'])[0].toLowerCase()
  const filename = base + ext

  const { sessionUrl } = await initUpload(filename, file.type || 'application/octet-stream', file.size)

  let offset = 0
  let uploaded = null
  while (offset < file.size) {
    const end = Math.min(offset + PUT_CHUNK, file.size)
    const buf = await file.slice(offset, end).arrayBuffer()
    const result = await putChunk(sessionUrl, arrayBufferToBase64(buf), offset, file.size)
    offset = end
    onProgress?.(offset / file.size)
    if (result.done) {
      uploaded = result.file
      break
    }
  }
  if (!uploaded) throw new Error('アップロードが完了しませんでした')

  // Mac側はこのJSONの出現を処理開始の合図にする。動画より後に置くこと
  await writeSidecar(base + '.json', {
    filename,
    title: title || file.name.replace(/\.[^.]+$/, ''),
    tags,
    uploadedAt: new Date().toISOString(),
    driveFileId: uploaded.id,
  })

  return { filename, driveFileId: uploaded.id }
}
