import { h, clear } from '../lib/dom.js'
import { newProfile, normalizeBaseUrl, validateApiKey } from '../lib/settings.js'
import { jsonToProfiles, profilesToJson, sampleProfilesJson } from '../lib/profiles-json.js'
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
    let mode = 'form' // 'form' | 'json'

    const tabs = h('div', { class: 'prof-tabs' })
    const body = h('div')
    const result = h('div', { class: 'hint', style: { marginBottom: '12px' } })

    const current = () => draft.profiles.find((p) => p.id === editingId) ?? draft.profiles[0]

    // ---- タブ行（フォームモード用のプロファイル切替 + JSON切替ボタン） ----
    function renderTabs() {
      clear(tabs)

      if (mode === 'json') {
        tabs.append(
          h('button', {
            class: 'prof-tab',
            text: '← フォームに戻る',
            onClick: () => {
              mode = 'form'
              renderTabs()
              renderBody()
            },
          }),
        )
        return
      }

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
              renderBody()
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
            renderBody()
          },
        }),
      )
      tabs.append(
        h('button', {
          class: 'prof-tab json',
          text: 'JSONで一括編集',
          onClick: () => {
            commit()
            mode = 'json'
            renderTabs()
            renderBody()
          },
        }),
      )
    }

    // ---- フォームモード ----
    let inputs = {}

    /** 入力中の値を draft に書き戻す */
    function commit() {
      if (mode !== 'form') return
      const p = current()
      if (!p || !inputs.label) return
      p.label = inputs.label.value.trim() || p.id
      p.id = inputs.id.value.trim() || p.id
      p.baseUrl = normalizeBaseUrl(inputs.baseUrl.value)
      p.apiKey = inputs.apiKey.value.trim()
      p.model = inputs.model.value.trim()
      p.numCtx = Number(inputs.numCtx.value) || 32768
      p.gasUrl = normalizeBaseUrl(inputs.gasUrl.value)
      p.controlToken = inputs.controlToken.value.trim()
    }

    function renderForm() {
      const p = current()
      clear(body)

      inputs = {
        label: h('input', { value: p.label }),
        id: h('input', { value: p.id }),
        baseUrl: h('input', { value: p.baseUrl, placeholder: 'https://xxx.ngrok-free.dev/v1' }),
        apiKey: h('input', { type: 'password', value: p.apiKey, placeholder: 'PROXY_API_KEY' }),
        model: h('input', { value: p.model, placeholder: 'gemma4:12b' }),
        numCtx: h('input', { type: 'number', value: p.numCtx }),
        gasUrl: h('input', {
          value: p.gasUrl,
          placeholder: 'https://script.google.com/macros/s/.../exec',
        }),
        controlToken: h('input', {
          type: 'password',
          value: p.controlToken,
          placeholder: 'CONTROL_TOKEN',
        }),
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

      body.append(
        h('div', { class: 'row' }, field('表示名', inputs.label), field('ID', inputs.id)),
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
                renderBody()
              },
            }),
        ),
        h('div', {
          class: 'section',
          text: 'このエンドポイントの起動制御（個別の GAS プロジェクト）',
        }),
        field('GAS ウェブアプリ URL', inputs.gasUrl),
        field('CONTROL_TOKEN', inputs.controlToken),
      )
    }

    // ---- JSON 一括編集モード ----
    function renderJson() {
      clear(body)
      result.textContent = ''

      const textarea = h('textarea', {
        rows: '16',
        style: { fontFamily: 'ui-monospace, Menlo, monospace', fontSize: '12.5px' },
      })
      textarea.value = draft.profiles.length ? profilesToJson(draft.profiles) : sampleProfilesJson()

      const jsonErr = h('div', { class: 'err' })

      const applyBtn = h('button', {
        class: 'primary',
        text: 'この内容で置き換える',
        onClick: () => {
          try {
            const parsed = jsonToProfiles(textarea.value)
            // apiKey の混入文字だけ先に検証しておく（空欄は許容し、フォーム側で後埋めできる）
            for (const p of parsed) {
              if (!p.apiKey) continue
              const key = validateApiKey(p.apiKey)
              if (!key.ok) throw new Error(`id "${p.id}": ${key.error}`)
              p.apiKey = key.value
            }
            draft.profiles = parsed
            if (!parsed.some((p) => p.id === draft.activeId)) {
              draft.activeId = parsed[0].id
            }
            editingId = draft.activeId
            jsonErr.textContent = ''
            result.textContent = `${parsed.length}件のエンドポイントを反映しました（保存を押すまで確定しません）`
          } catch (e) {
            jsonErr.textContent = e.message
          }
        },
      })

      body.append(
        h('div', {
          class: 'hint',
          style: { marginBottom: '10px' },
          text:
            'この配列で現在のエンドポイントを丸ごと置き換えます。id は他アプリや GAS 側の識別子と' +
            '一致させてください。apiKey・controlToken を空にした場合は各タブで個別に入力できます。',
        }),
        h('div', { class: 'field' }, textarea, jsonErr),
        h('div', { class: 'row' }, applyBtn),
      )
    }

    function renderBody() {
      if (mode === 'json') renderJson()
      else renderForm()
    }

    const temperature = h('input', { type: 'number', step: '0.1', value: draft.temperature })
    const systemPrompt = h('textarea', { rows: '2' })
    systemPrompt.value = draft.systemPrompt

    renderTabs()
    renderBody()

    return h(
      'div',
      { class: 'modal' },
      h('h2', { text: '設定' }),
      tabs,
      body,
      h('div', { class: 'section', text: '共通' }),
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
