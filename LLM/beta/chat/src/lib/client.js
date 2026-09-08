/**
 * Gemini は OpenAI 互換エンドポイントを持つが、Ollama 向けの拡張
 * （num_ctx / think）や独自ヘッダを付けると 400 で弾かれる。
 */
export function isGemini(s) {
  return /generativelanguage\.googleapis\.com/i.test(s?.baseUrl ?? '')
}

function headers(s) {
  const h = {
    'Content-Type': 'application/json',
    Authorization: `Bearer ${s.apiKey}`,
  }
  // ngrok の警告ページを回避する（予約ドメインでは通常不要だが保険）
  if (!isGemini(s)) h['ngrok-skip-browser-warning'] = 'true'
  return h
}

function chatBody(s, messages) {
  const base = { model: s.model, messages, stream: true, temperature: s.temperature }
  if (isGemini(s)) return { ...base, stream_options: { include_usage: true } }
  return {
    ...base,
    num_ctx: s.numCtx,
    // qwen3 系はデフォルトで思考モードが有効。オフにすると応答開始が大きく速くなる。
    // プロキシ側は think があればそのまま Ollama に転送する実装。
    ...(typeof s.think === 'boolean' ? { think: s.think } : {}),
  }
}

/** 接続確認。利用可能なモデル ID の配列を返す */
export async function fetchModels(s, signal) {
  const res = await fetch(`${s.baseUrl}/models`, { headers: headers(s), signal })
  if (!res.ok) {
    const body = await res.text().catch(() => '')
    throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`)
  }
  const json = await res.json()
  // Gemini は "models/gemini-2.5-flash" の形で返すが、指定は接頭辞なしでよい
  return (json?.data ?? []).map((m) => String(m.id).replace(/^models\//, ''))
}

/**
 * /v1/chat/completions を SSE で読む非同期ジェネレータ。
 * yield されるのは { delta } または { usage }。
 */
export async function* streamChat(s, messages, signal) {
  const res = await fetch(`${s.baseUrl}/chat/completions`, {
    method: 'POST',
    headers: headers(s),
    signal,
    body: JSON.stringify(chatBody(s, messages)),
  })

  if (!res.ok || !res.body) {
    const body = await res.text().catch(() => '')
    throw new Error(`HTTP ${res.status} ${body.slice(0, 300)}`)
  }

  const reader = res.body.getReader()
  const decoder = new TextDecoder()
  let buf = ''

  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    buf += decoder.decode(value, { stream: true })

    let nl
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).trim()
      buf = buf.slice(nl + 1)
      if (!line.startsWith('data:')) continue

      const payload = line.slice(5).trim()
      if (payload === '[DONE]') return

      try {
        const json = JSON.parse(payload)
        const delta = json?.choices?.[0]?.delta?.content ?? ''
        if (delta) yield { delta }
        if (json?.usage) yield { usage: json.usage }
      } catch {
        // 不完全な行は次のチャンクで補完されるため無視
      }
    }
  }
}
