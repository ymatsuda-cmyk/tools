/**
 * トークン数の推定。
 *
 * 実測（P100 16GB / gemma4:12b、5往復の会話から回帰）:
 *   tokens ≒ 0.50 × 文字数 + 17   （日本語主体）
 *   英数字・ID の羅列は約 0.82 トークン/文字
 *
 * 係数はやや安全側（多めに見積もる方向）。
 * さらに実測 usage.prompt_tokens との比を学習して補正する。
 */
const CJK = /[\u3000-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uff00-\uffef]/

const CJK_PER_CHAR = 0.55
const OTHER_PER_CHAR = 0.45

/** num_ctx 32768 のうち、出力用に 2048 を残した実質の入力上限 */
export const MAX_INPUT_TOKENS = 30000

/** 実測の prefill 速度 */
export const TOKENS_PER_SEC = 270

// ---- 実測にもとづく自動補正 ----

const CAL_KEY = 'gemma-chat.calibration'
const CAL_MIN = 0.6
const CAL_MAX = 1.5
const CAL_ALPHA = 0.25 // 指数移動平均の重み

let factor = null

function loadFactor() {
  if (factor !== null) return factor
  const raw = Number(localStorage.getItem(CAL_KEY))
  factor = Number.isFinite(raw) && raw > 0 ? raw : 1
  return factor
}

/**
 * サーバの実測値で補正係数を更新する。
 * 画像を含む場合は推定根拠が異なるため呼び出さないこと。
 */
export function recordCalibration(estimated, actual) {
  if (!estimated || !actual) return
  const observed = actual / estimated
  if (!Number.isFinite(observed) || observed < 0.2 || observed > 3) return
  const next = loadFactor() * (1 - CAL_ALPHA) + observed * CAL_ALPHA
  factor = Math.min(CAL_MAX, Math.max(CAL_MIN, next))
  try {
    localStorage.setItem(CAL_KEY, String(factor))
  } catch {
    /* 保存できなくても推定は続行 */
  }
}

export function getCalibration() {
  return loadFactor()
}

// ---- 推定 ----

/** 補正前の素の推定 */
export function estimateRaw(text) {
  let cjk = 0
  let other = 0
  for (const ch of text ?? '') {
    if (CJK.test(ch)) cjk++
    else other++
  }
  return Math.ceil(cjk * CJK_PER_CHAR + other * OTHER_PER_CHAR)
}

/** 補正込みの推定。UI とガードはこちらを使う */
export function estimateTokens(text) {
  return Math.ceil(estimateRaw(text) * loadFactor())
}

export function estimateWaitSec(tokens) {
  return Math.max(1, Math.round(tokens / TOKENS_PER_SEC))
}

export function formatWait(sec) {
  return sec < 60 ? `約${sec}秒` : `約${Math.round(sec / 60)}分`
}

export function formatTokens(n) {
  return n < 1000 ? String(n) : `${(n / 1000).toFixed(1)}k`
}
