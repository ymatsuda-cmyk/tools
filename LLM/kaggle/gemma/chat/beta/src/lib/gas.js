const TIMEOUT_MS = 25000

/**
 * GAS web app を JSONP で呼ぶ。プロファイル単体（gasUrl + controlToken）を渡す。
 *
 * fetch を使わない理由: GAS の /exec はレスポンス時に
 * script.googleusercontent.com へリダイレクトするため、
 * ブラウザからの CORS が環境によって不安定になる。
 * JSONP なら CORS の対象外なので確実に通る。
 * その代わり controlToken が URL に載るので、
 * このトークンで守るのは「GPU の起動」だけに留めること。
 *
 * このプロジェクトは1エンドポイント=1GASなので id パラメータは不要。
 */
export function callGas(profile, action, params = {}) {
  return new Promise((resolve, reject) => {
    if (!profile?.gasUrl) return reject(new Error('GAS URL が未設定です'))
    if (!profile?.controlToken) return reject(new Error('CONTROL_TOKEN が未設定です'))

    const cb = `__gas_${Math.random().toString(36).slice(2, 10)}`
    const script = document.createElement('script')
    let settled = false

    const cleanup = () => {
      delete window[cb]
      script.remove()
      clearTimeout(timer)
    }
    const timer = setTimeout(() => {
      if (settled) return
      settled = true
      cleanup()
      reject(new Error('GAS が応答しません（タイムアウト）'))
    }, TIMEOUT_MS)

    window[cb] = (data) => {
      if (settled) return
      settled = true
      cleanup()
      if (data && data.error) reject(new Error(data.error))
      else resolve(data)
    }
    script.onerror = () => {
      if (settled) return
      settled = true
      cleanup()
      reject(new Error('GAS に接続できません（URL とデプロイ設定を確認）'))
    }

    const q = new URLSearchParams({ ...params, action, token: profile.controlToken, callback: cb })
    script.src = `${profile.gasUrl}?${q.toString()}`
    document.head.append(script)
  })
}

export const getStatus = (profile) => callGas(profile, 'status')
export const startKernel = (profile) => callGas(profile, 'start')
export const stopKernel = (profile) => callGas(profile, 'stop')
export const forceStopKernel = (profile) => callGas(profile, 'forceStop')

/**
 * Kaggle の status と proxyAlive から表示用の状態を決める。
 * status が running でもモデルのロード中は使えないため、
 * 「利用可能」の判定にはハートビートを使う。
 */
export function deriveState(st) {
  if (!st || !st.success) {
    return { key: 'unknown', label: '状態不明', tone: 'ng', canStart: true, canStop: false }
  }
  if (st.stopRequested) {
    return { key: 'stopping', label: '停止処理中', tone: 'wait', canStart: false, canStop: false }
  }
  const live = st.status === 'running' || st.status === 'queued'
  if (live && st.proxyAlive) {
    return { key: 'ready', label: '稼働中', tone: 'ok', canStart: false, canStop: true }
  }
  // Kaggle 上は動いているのに応答がない。停止を押せば GAS が強制停止に切り替える。
  if (st.zombie) {
    return { key: 'zombie', label: '応答なし', tone: 'ng', canStart: false, canStop: true }
  }
  if (live) {
    return { key: 'booting', label: '起動中', tone: 'wait', canStart: false, canStop: true }
  }
  return { key: 'stopped', label: '停止中', tone: 'ng', canStart: true, canStop: false }
}

export function formatHours(min) {
  if (typeof min !== 'number') return '—'
  return `${(min / 60).toFixed(1)}h`
}

/** 短い経過表示。1時間未満は分で出す */
export function formatElapsed(min) {
  if (typeof min !== 'number') return null
  if (min < 60) return `${min}分`
  return `${(min / 60).toFixed(1)}時間`
}
