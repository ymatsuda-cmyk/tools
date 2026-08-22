import { h } from '../lib/dom.js'
import { normalizeBaseUrl, validateApiKey } from '../lib/settings.js'
import { fetchModels } from '../lib/client.js'
import { openModal } from './modal.js'

export function openSettings(settings, onSave) {
  openModal(({ close }) => {
    const draft = { ...settings }

    const baseUrl = h('input', { value: draft.baseUrl, placeholder: 'https://example.ngrok-free.dev/v1' })
    const apiKey = h('input', { type: 'password', value: draft.apiKey, placeholder: 'PROXY_API_KEY' })
    const keyErr = h('div', { class: 'err' })
    const modelInput = h('input', { value: draft.model })
    const modelSelect = h('select', { style: { display: 'none' } })
    const numCtx = h('input', { type: 'number', value: draft.numCtx })
    const temperature = h('input', { type: 'number', step: '0.1', value: draft.temperature })
    const systemPrompt = h('textarea', { rows: '2' })
    systemPrompt.value = draft.systemPrompt
    const gasUrl = h('input', { value: draft.gasUrl, placeholder: 'https://script.google.com/macros/s/.../exec' })
    const controlToken = h('input', { type: 'password', value: draft.controlToken, placeholder: 'CONTROL_TOKEN' })
    const result = h('div', { class: 'hint', style: { marginBottom: '12px' } })

    const currentModel = () =>
      modelSelect.style.display === 'none' ? modelInput.value.trim() : modelSelect.value

    const collect = () => ({
      ...draft,
      baseUrl: normalizeBaseUrl(baseUrl.value),
      apiKey: apiKey.value,
      model: currentModel(),
      numCtx: Number(numCtx.value),
      temperature: Number(temperature.value),
      systemPrompt: systemPrompt.value,
      gasUrl: gasUrl.value.trim(),
      controlToken: controlToken.value.trim(),
    })

    const testBtn = h('button', {
      text: '接続テスト',
      onClick: async () => {
        const s = collect()
        const key = validateApiKey(s.apiKey)
        keyErr.textContent = key.ok ? '' : key.error
        if (!key.ok) return

        testBtn.disabled = true
        testBtn.textContent = '確認中…'
        result.textContent = ''
        try {
          const list = await fetchModels({ ...s, apiKey: key.value })
          result.textContent = `接続できました（${list.length}件）`
          modelSelect.innerHTML = ''
          for (const id of list) modelSelect.append(h('option', { value: id, text: id }))
          modelSelect.value = list.includes(s.model) ? s.model : (list[0] ?? s.model)
          modelSelect.style.display = ''
          modelInput.style.display = 'none'
        } catch (e) {
          result.textContent = `失敗: ${e.message}`
        } finally {
          testBtn.disabled = false
          testBtn.textContent = '接続テスト'
        }
      },
    })

    const field = (label, ...ctrl) =>
      h('div', { class: 'field' }, h('label', { text: label }), ...ctrl)

    return h(
      'div',
      { class: 'modal' },
      h('h2', { text: '接続設定' }),
      field('ベースURL', baseUrl),
      field('APIキー', apiKey, keyErr),
      field('モデル', modelInput, modelSelect),
      h('div', { class: 'row' }, field('num_ctx', numCtx), field('temperature', temperature)),
      field('システムプロンプト', systemPrompt),
      h('div', { class: 'section', text: 'Kaggle 起動制御（任意）' }),
      field('GAS ウェブアプリ URL', gasUrl),
      field('CONTROL_TOKEN', controlToken),
      result,
      h(
        'div',
        { class: 'row' },
        testBtn,
        h('button', { text: '閉じる', onClick: close }),
        h('button', {
          class: 'primary',
          text: '保存',
          onClick: () => {
            const s = collect()
            const key = validateApiKey(s.apiKey)
            keyErr.textContent = key.ok ? '' : key.error
            if (!key.ok) return
            onSave({ ...s, apiKey: key.value })
            close()
          },
        }),
      ),
    )
  })
}
