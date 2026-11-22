// src/pass.js — 合言葉（権限コードをこの端末に覚えておき、合言葉だけで開けるようにする）
//   ・権限コードは、合言葉から PBKDF2 で作った鍵で暗号化して、この端末の localStorage にだけ置く
//   ・合言葉も権限コードも、GAS や GitHub には送らない（権限コードは今までどおりログインのときだけ GAS に送る）
//   ・合言葉を忘れたら、権限コードで開き直して決め直す
import { sealText, openText } from './crypto.js'

const KEY = 'm2.pw'
export const PASS_MIN = 6

export const passRegistered = () => !!localStorage.getItem(KEY)
export const passForget = () => localStorage.removeItem(KEY)

export async function passRegister(pass, code) {
  localStorage.setItem(KEY, JSON.stringify(await sealText(pass, code)))
}

/** → 権限コード。合言葉が違うときは「合言葉が違います。」 */
export async function passUnlock(pass) {
  let v = null
  try { v = JSON.parse(localStorage.getItem(KEY) || 'null') } catch { /* 壊れていれば未登録として扱う */ }
  if (!v) throw new Error('合言葉が登録されていません。権限コードで開いてください。')
  try { return await openText(pass, v) } catch { throw new Error('合言葉が違います。') }
}
