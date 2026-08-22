const TIMEOUT_MS = 20000

/**
 * GAS web app を JSONP で呼ぶ。
 *
 * fetch を使わない理由: GAS の /exec はレスポンス時に
 * script.googleusercontent.com へリダイレクトするため、
 * ブラウザからの CORS が環境によって不安定になる。
 * JSONP なら CORS の対象外なので確実に通る。
 * その代わり CONTROL_TOKEN が URL に載るので、
 * このトークンで守っているのは「GPU の起動」だけに留めること。
 */
export function callGas(settings, action, params = {}) {
  return new Promise((resolve, reject) => {
    if (!settings.gasUrl) return reject(new Error('GAS URL が未設定です'))
    if (!settings.controlToken) return reject(new Error('CONTROL_TOKEN が未設定です'))

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

    const q = new URLSearchParams({
      ...params,
      action,
      token: settings.controlToken,
      callback: cb,
    })
    script.src = `${settings.gasUrl}?${q.toString()}`
    document.head.append(script)
  })
}

export const getStatus = (s) => callGas(s, 'status')
export const startKernel = (s) => callGas(s, 'start')
export const stopKernel = (s) => callGas(s, 'stop')

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
  if (live) {
    return { key: 'booting', label: '起動中', tone: 'wait', canStart: false, canStop: true }
  }
  return { key: 'stopped', label: '停止中', tone: 'ng', canStart: true, canStop: false }
}

export function formatQuota(st) {
  if (!st || typeof st.weeklyRemainMin !== 'number') return null
  return `週残 ${(st.weeklyRemainMin / 60).toFixed(1)}h`
}
