/**
 * トークン数の推定。
 * 実測（P100 16GB / gemma4:12b）:
 *   日本語の散文    約 2.3 字/tok
 *   英数字・ID混在  約 1.2 字/tok
 * 係数は安全側（多めに見積もる方向）に振ってある。
 */
const CJK = /[\u3000-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]/

/** @param {string} text @returns {number} */
export function estimateTokens(text) {
  let cjk = 0
  let other = 0
  for (const ch of text) {
    if (CJK.test(ch)) cjk++
    else other++
  }
  return Math.ceil(cjk * 0.7 + other * 0.5)
}

/** num_ctx 32768 のうち、出力用に 2048 を残した実質の入力上限 */
export const MAX_INPUT_TOKENS = 30000

/** 実測の prefill 速度 */
export const TOKENS_PER_SEC = 270

export function estimateWaitSec(tokens) {
  return Math.max(1, Math.round(tokens / TOKENS_PER_SEC))
}

export function formatWait(sec) {
  return sec < 60 ? `約${sec}秒` : `約${Math.round(sec / 60)}分`
}

export function formatTokens(n) {
  return n < 1000 ? String(n) : `${(n / 1000).toFixed(1)}k`
}
