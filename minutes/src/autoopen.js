// src/autoopen.js — この端末で「自動で開く」ための権限コードの控え
//
//   ・権限コードは、この端末だけで作った「取り出せない鍵」（AES-GCM, extractable:false）で暗号化して置く
//   ・鍵は IndexedDB、暗号文は localStorage。鍵そのものはブラウザの外へ出せないので、
//     localStorage の中身を写されただけでは権限コードは読めない
//   ・端末そのものを使われれば開けてしまうので、心配なときは設定で「自動で開く」をオフにする
import { b64, unb64, rnd } from './crypto.js'

const DB = 'minutestock'
const STORE = 'kv'
const KEY_ID = 'autoKey'
const LS = 'm2.auto'
const OFF = 'm2.autoOff'
const te = new TextEncoder()
const td = new TextDecoder()

function db() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1)
    r.onupgradeneeded = () => r.result.createObjectStore(STORE)
    r.onsuccess = () => res(r.result)
    r.onerror = () => rej(r.error)
  })
}
async function kv(mode, fn) {
  const d = await db()
  return new Promise((res, rej) => {
    const q = fn(d.transaction(STORE, mode).objectStore(STORE))
    q.onsuccess = () => res(q.result)
    q.onerror = () => rej(q.error)
  })
}

async function deviceKey(create) {
  let k = await kv('readonly', (s) => s.get(KEY_ID))
  if (!k && create) {
    k = await crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt'])
    await kv('readwrite', (s) => s.put(k, KEY_ID))
  }
  return k || null
}

/** 自動で開く設定（既定はオン） */
export const autoEnabled = () => localStorage.getItem(OFF) !== 'true'
export function setAutoEnabled(on) { if (on) localStorage.removeItem(OFF); else localStorage.setItem(OFF, 'true') }

/** 権限コードの控えがあるか */
export const autoSaved = () => !!localStorage.getItem(LS)

export async function autoSave(code) {
  const key = await deviceKey(true)
  const iv = rnd(12)
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, te.encode(code)))
  localStorage.setItem(LS, JSON.stringify({ iv: b64(iv), ct: b64(ct) }))
}

/** → 権限コード（読めなければ控えを消して null） */
export async function autoLoad() {
  let v
  try { v = JSON.parse(localStorage.getItem(LS) || 'null') } catch { v = null }
  if (!v) return null
  try {
    const key = await deviceKey(false)
    if (!key) throw new Error('no key')
    return td.decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(v.iv) }, key, unb64(v.ct)))
  } catch {
    await autoForget()
    return null
  }
}

export async function autoForget() {
  localStorage.removeItem(LS)
  try { await kv('readwrite', (s) => s.delete(KEY_ID)) } catch { /* IndexedDB が使えないときは控えも作れていない */ }
}
