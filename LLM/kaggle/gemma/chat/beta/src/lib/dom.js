/** 要素を組み立てる最小ヘルパー */
export function h(tag, props = {}, ...children) {
  const el = document.createElement(tag)
  for (const [k, v] of Object.entries(props)) {
    if (v === null || v === undefined || v === false) continue
    if (k === 'class') el.className = v
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v)
    else if (k.startsWith('on') && typeof v === 'function') {
      el.addEventListener(k.slice(2).toLowerCase(), v)
    } else if (k === 'text') el.textContent = v
    else el.setAttribute(k, v === true ? '' : v)
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue
    el.append(c instanceof Node ? c : document.createTextNode(String(c)))
  }
  return el
}

export function clear(el) {
  while (el.firstChild) el.removeChild(el.firstChild)
}

const loaded = new Map()

/** UMD 配布のライブラリを script タグで遅延読み込みし、グローバルを返す */
export function loadScript(url, globalName) {
  if (loaded.has(url)) return loaded.get(url)
  const p = new Promise((resolve, reject) => {
    const s = document.createElement('script')
    s.src = url
    s.onload = () => resolve(window[globalName])
    s.onerror = () => reject(new Error(`読み込みに失敗しました: ${url}`))
    document.head.append(s)
  })
  loaded.set(url, p)
  return p
}
