function headers(s) {
  return {
    'Content-Type': 'application/json',
    Authorization: `Bearer ${s.apiKey}`,
    'ngrok-skip-browser-warning': 'true',
  }
}

export async function fetchModels(s, signal) {
  const res = await fetch(`${s.baseUrl}/models`, { headers: headers(s), signal })
  if (!res.ok) {
    const body = await res.text().catch(() => '')
    throw new Error(`HTTP ${res.status} ${body.slice(0, 200)}`)
  }
  const json = await res.json()
  return (json?.data ?? []).map((m) => m.id)
}

export async function* streamChat(s, messages, signal) {
  const res = await fetch(`${s.baseUrl}/chat/completions`, {
    method: 'POST',
    headers: headers(s),
    signal,
    body: JSON.stringify({
      model: s.model,
      messages,
      stream: true,
      num_ctx: s.numCtx,
      temperature: s.temperature,
      ...(typeof s.think === 'boolean' ? { think: s.think } : {}),
    }),
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
