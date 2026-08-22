import { h } from '../lib/dom.js'

/** ドラッグ中の内容にファイルが含まれるか */
function hasFiles(e) {
  const types = e.dataTransfer?.types
  return types ? Array.from(types).includes('Files') : false
}

/**
 * 要素全体をドロップ領域にする。
 * dragenter / dragleave は子要素をまたぐたびに発火するため、
 * 深さを数えないとオーバーレイがちらつく。
 *
 * @param {HTMLElement} el
 * @param {(files: FileList) => void} onFiles
 */
export function setupDropzone(el, onFiles) {
  let depth = 0

  el.append(
    h(
      'div',
      { class: 'dropzone' },
      h('div', { class: 'dropzone-inner', text: 'ドロップして添付' }),
    ),
  )

  el.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return
    e.preventDefault()
    depth++
    el.classList.add('dragging')
  })

  el.addEventListener('dragover', (e) => {
    if (!hasFiles(e)) return
    e.preventDefault()
    e.dataTransfer.dropEffect = 'copy'
  })

  el.addEventListener('dragleave', () => {
    depth = Math.max(0, depth - 1)
    if (depth === 0) el.classList.remove('dragging')
  })

  el.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return
    e.preventDefault()
    depth = 0
    el.classList.remove('dragging')
    onFiles(e.dataTransfer.files)
  })

  // ドロップ領域の外に落とした場合にブラウザがファイルを開いてしまうのを防ぐ
  window.addEventListener('dragover', (e) => hasFiles(e) && e.preventDefault())
  window.addEventListener('drop', (e) => hasFiles(e) && e.preventDefault())
}
