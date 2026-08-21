import { h, clear } from '../lib/dom.js'
import { newProfile, normalizeBaseUrl, validateApiKey } from '../lib/settings.js'
import { fetchModels } from '../lib/client.js'
import { openModal } from './modal.js'

const field = (label, ...ctrl) => h('div', { class: 'field' }, h('label', { text: label }), ...ctrl)

export function openSettings(settings, onSave) {
  openModal(({ close }) => {
    // 編集はコピー上で行い、保存時にだけ反映する
    const draft = JSON.parse(JSON.stringify(settings))
    if (!draft.profiles.length) {
      draft.profiles.push(newProfile({ id: 'a', label: 'エンドポイント A' }))
      draft.activeId = draft.profiles[0].id
    }
    let editingId = draft.activeId ?? draft.profiles[0].id

    const tabs = h('div', { class: 'prof-tabs' })
    const form = h('div')
    const result = h('div', { class: 'hint', style: { marginBottom: '12px' } })

    const current = () => draft.profiles.find((p) => p.id === editingId) ?? draft.profiles[0]

    function renderTabs() {
      clear(tabs)
      for (const p of draft.profiles) {
        tabs.append(
          h('button', {
            class: p.id === editingId ? 'prof-tab on' : 'prof-tab',
            text: p.label || p.id,
            onClick: () => {
              commit()
              editingId = p.id
              result.textContent = ''
              renderTabs()
              renderForm()
            },
          }),
        )
      }
      tabs.append(
        h('button', {
          class: 'prof-tab add',
          text: '＋',
          title: '接続を追加',
          onClick: () => {
            commit()
            const p = newProfile({ label: `エンドポイント ${draft.profiles.length + 1}` })
            draft.profiles.push(p)
            editingId = p.id
            renderTabs()
            renderForm()
          },
        }),
      )
    }

    let inputs = {}

    /** 入力中の値を draft に書き戻す */
    function commit() {
      const p = current()
      if (!p || !inputs.label) return
      p.label = inputs.label.value.trim() || p.id
      p.id = inputs.id.value.trim() || p.id
      p.baseUrl = normalizeBaseUrl(inputs.baseUrl.value)
      p.apiKey = inputs.apiKey.value.trim()
      p.model = inputs.model.value.trim()
      p.numCtx = Number(inputs.numCtx.value) || 32768
    }

    function renderForm() {
      const p = current()
      clear(form)

      inputs = {
        label: h('input', { value: p.label }),
        id: h('input', { value: p.id }),
        baseUrl: h('input', { value: p.baseUrl, placeholder: 'https://xxx.ngrok-free.dev/v1' }),
        apiKey: h('input', { type: 'password', value: p.apiKey, placeholder: 'PROXY_API_KEY' }),
        model: h('input', { value: p.model, placeholder: 'gemma4:12b' }),
        numCtx: h('input', { type: 'number', value: p.numCtx }),
      }
      const keyErr = h('div', { class: 'err' })

      const testBtn = h('button', {
        text: '接続テスト',
        onClick: async () => {
          commit()
          const key = validateApiKey(current().apiKey)
          keyErr.textContent = key.ok ? '' : key.error
          if (!key.ok) return
          testBtn.disabled = true
          testBtn.textContent = '確認中…'
          result.textContent = ''
          try {
            const list = await fetchModels({ ...current(), apiKey: key.value })
            result.textContent = `接続できました: ${list.join(', ')}`
            if (!current().model && list.length) inputs.model.value = list[0]
          } catch (e) {
            result.textContent = `失敗: ${e.message}`
          } finally {
            testBtn.disabled = false
            testBtn.textContent = '接続テスト'
          }
        },
      })

      form.append(
        h('div', { class: 'row' }, field('表示名', inputs.label), field('ID', inputs.id)),
        h('div', {
          class: 'hint',
          style: { margin: '-8px 0 14px' },
          text: 'ID は GAS の ENDPOINTS に書いた id と一致させる',
        }),
        field('ベースURL', inputs.baseUrl),
        field('APIキー', inputs.apiKey, keyErr),
        h('div', { class: 'row' }, field('モデル', inputs.model), field('num_ctx', inputs.numCtx)),
        h(
          'div',
          { class: 'row', style: { marginBottom: '18px' } },
          testBtn,
          draft.profiles.length > 1 &&
            h('button', {
              text: 'この接続を削除',
              onClick: () => {
                draft.profiles = draft.profiles.filter((x) => x.id !== current().id)
                editingId = draft.profiles[0].id
                renderTabs()
                renderForm()
              },
            }),
        ),
      )
    }

    const gasUrl = h('input', {
      value: draft.gasUrl,
      placeholder: 'https://script.google.com/macros/s/.../exec',
    })
    const controlToken = h('input', {
      type: 'password',
      value: draft.controlToken,
      placeholder: 'CONTROL_TOKEN',
    })
    const temperature = h('input', { type: 'number', step: '0.1', value: draft.temperature })
    const systemPrompt = h('textarea', { rows: '2' })
    systemPrompt.value = draft.systemPrompt

    renderTabs()
    renderForm()

    return h(
      'div',
      { class: 'modal' },
      h('h2', { text: '設定' }),
      tabs,
      form,
      h('div', { class: 'section', text: '共通' }),
      field('GAS ウェブアプリ URL', gasUrl),
      field('CONTROL_TOKEN', controlToken),
      field('temperature', temperature),
      field('システムプロンプト', systemPrompt),
      result,
      h(
        'div',
        { class: 'row' },
        h('button', { text: '閉じる', onClick: close }),
        h('button', {
          class: 'primary',
          text: '保存',
          onClick: () => {
            commit()
            for (const p of draft.profiles) {
              if (!p.apiKey) continue
              const key = validateApiKey(p.apiKey)
              if (!key.ok) {
                result.textContent = `${p.label}: ${key.error}`
                return
              }
              p.apiKey = key.value
            }
            draft.gasUrl = gasUrl.value.trim()
            draft.controlToken = controlToken.value.trim()
            draft.temperature = Number(temperature.value)
            draft.systemPrompt = systemPrompt.value
            if (!draft.profiles.some((p) => p.id === draft.activeId)) {
              draft.activeId = draft.profiles[0].id
            }
            onSave(draft)
            close()
          },
        }),
      ),
    )
  })
}
