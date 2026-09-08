import { newProfile } from './settings.js'

/**
 * JSON 一括編集で使うプロファイルの見た目上のキー。
 * newProfile() の内部キーと一部名前を変え、人が書きやすい形にしている。
 */
const FIELDS = [
  'id', 'label', 'baseUrl', 'apiKey', 'model', 'numCtx', 'gasUrl', 'controlToken', 'think',
]

/** プロファイル配列 → 貼り付け・保存しやすい JSON 文字列 */
export function profilesToJson(profiles) {
  const out = profiles.map((p) => {
    const o = {}
    for (const k of FIELDS) o[k] = p[k]
    return o
  })
  return JSON.stringify(out, null, 2)
}

/**
 * JSON 文字列 → プロファイル配列。
 * 検証だけ行い、成功時のみ配列を返す。失敗時は理由を投げる。
 */
export function jsonToProfiles(text) {
  let raw
  try {
    raw = JSON.parse(text)
  } catch (e) {
    throw new Error(`JSON として読めません: ${e.message}`)
  }
  if (!Array.isArray(raw)) throw new Error('配列 [ ... ] の形式で入力してください')
  if (!raw.length) throw new Error('少なくとも1件は必要です')

  const seen = new Set()
  const out = raw.map((item, i) => {
    if (!item || typeof item !== 'object') {
      throw new Error(`${i + 1}件目: オブジェクトではありません`)
    }
    const id = String(item.id ?? '').trim()
    if (!id) throw new Error(`${i + 1}件目: id が必須です`)
    if (seen.has(id)) throw new Error(`id "${id}" が重複しています`)
    seen.add(id)

    return newProfile({
      id,
      label: String(item.label ?? id),
      baseUrl: String(item.baseUrl ?? '').trim(),
      apiKey: String(item.apiKey ?? '').trim(),
      model: String(item.model ?? '').trim(),
      numCtx: Number(item.numCtx) || 32768,
      gasUrl: String(item.gasUrl ?? '').trim(),
      controlToken: String(item.controlToken ?? '').trim(),
      // 未指定 / null / 不正値はサーバー既定（think を送らない）扱いにする
      think: typeof item.think === 'boolean' ? item.think : null,
    })
  })

  return out
}

/** 一括編集の入力欄に出す雛形（未入力時） */
export function sampleProfilesJson() {
  return JSON.stringify(
    [
      {
        id: 'a',
        label: 'Gemma 4 12B',
        baseUrl: 'https://xxx.ngrok-free.dev/v1',
        apiKey: 'PROXY_API_KEY',
        model: 'gemma4:12b',
        numCtx: 32768,
        gasUrl: 'https://script.google.com/macros/s/xxx/exec',
        controlToken: 'CONTROL_TOKEN',
        think: null, // null=サーバー既定 / true=オン / false=オフ
      },
    ],
    null,
    2,
  )
}
