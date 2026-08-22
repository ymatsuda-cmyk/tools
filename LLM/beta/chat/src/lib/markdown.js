import { marked } from 'marked'
import DOMPurify from 'dompurify'

marked.setOptions({ breaks: true, gfm: true })

let hljs = null
let hljsLoading = null

function ensureHljs() {
  if (hljs) return Promise.resolve(hljs)
  if (!hljsLoading) {
    hljsLoading = import('highlight.js')
      .then((m) => {
        hljs = m.default ?? m
        return hljs
      })
      .catch(() => null)
  }
  return hljsLoading
}

/** Markdown を安全な HTML にして要素へ流し込む */
export function renderMarkdown(el, text) {
  el.innerHTML = DOMPurify.sanitize(marked.parse(text ?? ''))
  ensureHljs().then((lib) => {
    if (!lib) return
    for (const block of el.querySelectorAll('pre code')) {
      if (!block.dataset.highlighted) {
        lib.highlightElement(block)
        block.dataset.highlighted = '1'
      }
    }
  })
}

/**
 * ストリーミング中の再描画をスロットルする。
 * 毎トークン parse すると長文の後半で目に見えてカクつくため、
 * 既定 120ms に一度だけ描画し、最後は必ず描き切る。
 */
export function createThrottledRenderer(el, intervalMs = 120) {
  let pending = null
  let timer = null
  let last = 0

  const flush = () => {
    if (pending === null) return
    el.innerHTML = DOMPurify.sanitize(marked.parse(pending))
    pending = null
    last = Date.now()
  }

  return {
    update(text) {
      pending = text
      const wait = Math.max(0, intervalMs - (Date.now() - last))
      if (timer) return
      timer = setTimeout(() => {
        timer = null
        flush()
      }, wait)
    },
    finish(text) {
      if (timer) {
        clearTimeout(timer)
        timer = null
      }
      pending = null
      renderMarkdown(el, text)
    },
  }
}
