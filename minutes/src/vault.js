// src/vault.js — この端末のパスコードで包む保管庫
//
//   入れるもの：権限コード・AI の接続先（URL・API キー）。localStorage には暗号文だけを置く
//   鍵    ：パスコード → PBKDF2-SHA256（600,000 回）→ AES-GCM 256。パスコードはサーバー（GAS）へ送らない
//   開いた後：パスコードと鍵はこのページのメモリにだけ置く。鍵をかける・ページを閉じると消える
//   誤入力：MAX_FAIL 回続くと保管庫を消す（会議のデータは GitHub にあるので、権限コードを入れ直せば戻れる）
import { b64, unb64, rnd, PBKDF2_ITER } from './crypto.js'

const KEY = 'm2.vault'
const FAIL = 'm2.vaultFail'
export const MAX_FAIL = 5
export const MIN_PASS = 8

const subtle = globalThis.crypto.subtle
const te = new TextEncoder()
const td = new TextDecoder()

let mem = null            // { pass, key, salt, iter, data }
let queue = Promise.resolve()

async function derive(pass, salt, iter) {
  const base = await subtle.importKey('raw', te.encode(pass), 'PBKDF2', false, ['deriveKey'])
  return subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base,
    { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt'])
}

/** メモリの中身を暗号化して置く。続けて呼ばれても順番に書く */
function persist() {
  const snap = mem
  queue = queue.then(async () => {
    if (!snap || mem !== snap) return   // 書いている間に鍵をかけた・作り直した
    const iv = rnd(12)
    const ct = new Uint8Array(await subtle.encrypt({ name: 'AES-GCM', iv }, snap.key, te.encode(JSON.stringify(snap.data))))
    localStorage.setItem(KEY, JSON.stringify({ v: 1, iter: snap.iter, salt: b64(snap.salt), iv: b64(iv), ct: b64(ct) }))
  })
  return queue
}

export const vaultExists = () => !!localStorage.getItem(KEY)
export const vaultOpened = () => !!mem
export const vaultData = () => (mem ? mem.data : null)
export const vaultPass = () => (mem ? mem.pass : null)
/** 誤入力であと何回まで試せるか */
export const failsLeft = () => Math.max(0, MAX_FAIL - Number(localStorage.getItem(FAIL) || 0))

/** 新しく作る（すでにあれば置き換える） */
export async function vaultCreate(pass, data) {
  const salt = rnd(16)
  mem = { pass, key: await derive(pass, salt, PBKDF2_ITER), salt, iter: PBKDF2_ITER, data: { ...data } }
  localStorage.removeItem(FAIL)
  await persist()
}

/** パスコードで開く。違えば BADPASS（e.left＝あと何回）、規定回数に達したら保管庫を消して WIPED */
export async function vaultOpen(pass) {
  let v
  try { v = JSON.parse(localStorage.getItem(KEY) || 'null') } catch { v = null }
  if (!v) throw new Error('NOVAULT')
  const salt = unb64(v.salt)
  const iter = Number(v.iter) || PBKDF2_ITER
  const key = await derive(pass, salt, iter)
  let data
  try {
    data = JSON.parse(td.decode(await subtle.decrypt({ name: 'AES-GCM', iv: unb64(v.iv) }, key, unb64(v.ct))))
  } catch {
    const n = Number(localStorage.getItem(FAIL) || 0) + 1
    if (n >= MAX_FAIL) { vaultWipe(); throw new Error('WIPED') }
    localStorage.setItem(FAIL, String(n))
    const e = new Error('BADPASS'); e.left = MAX_FAIL - n
    throw e
  }
  localStorage.removeItem(FAIL)
  mem = { pass, key, salt, iter, data }
  return data
}

/** 開いている間だけ。メモリをすぐ書き換え、暗号化して置くのは続けて行う（同期の呼び出し元は待たなくてよい） */
export function vaultUpdate(patch) {
  if (!mem) return Promise.reject(new Error('保管庫が開いていません。'))
  mem.data = { ...mem.data, ...patch }
  return persist()
}

/** パスコードを変える（塩も鍵も作り直す） */
export async function vaultRepass(newPass) {
  if (!mem) throw new Error('保管庫が開いていません。')
  const salt = rnd(16)
  mem = { ...mem, pass: newPass, salt, iter: PBKDF2_ITER, key: await derive(newPass, salt, PBKDF2_ITER) }
  await persist()
}

/** 鍵をかける（暗号文は残す） */
export function vaultLock() { mem = null }

/** 消す（暗号文も誤入力の回数も） */
export function vaultWipe() {
  mem = null
  localStorage.removeItem(KEY)
  localStorage.removeItem(FAIL)
}
