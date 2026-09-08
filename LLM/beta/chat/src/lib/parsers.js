import { estimateTokens } from './tokens.js'
import { loadScript } from './dom.js'

/** ペーストをカード化する閾値 */
export const PASTE_THRESHOLD = 2000

const PDFJS_VERSION = '4.10.38'
const MAMMOTH_URL = 'https://cdn.jsdelivr.net/npm/mammoth@1.9.0/mammoth.browser.min.js'
const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'

/** スキャンPDFを画像化するときの上限と解像度 */
const PDF_PAGE_IMAGE_LIMIT = 8
const PDF_PAGE_IMAGE_WIDTH = 1280

/** 1ページあたりこれ未満しか文字が取れなければ、文字情報を持たないPDFとみなす */
const PDF_TEXT_PER_PAGE = 30

const TEXT_EXT =
  /\.(txt|md|markdown|csv|tsv|json|ya?ml|log|ts|tsx|js|jsx|py|java|kt|sql|html|css|xml)$/i

function makeId() {
  return Math.random().toString(36).slice(2, 10)
}

function pack(name, kind, text) {
  return { id: makeId(), name, kind, chars: text.length, tokens: estimateTokens(text), text }
}

/**
 * 画像1枚あたりの概算トークン数。
 * Gemma 系は固定長のビジョントークンを使うため文字数からは推定できない。
 * 実測 usage.prompt_tokens と突き合わせて調整すること。
 */
export const IMAGE_TOKENS = 300

async function readDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result)
    r.onerror = () => reject(new Error('画像の読み込みに失敗しました'))
    r.readAsDataURL(file)
  })
}

async function packImage(file, name) {
  const dataUrl = await readDataUrl(file)
  return {
    id: makeId(),
    name,
    kind: 'image',
    chars: 0,
    tokens: IMAGE_TOKENS,
    text: '',
    dataUrl,
    bytes: file.size,
  }
}

async function openPdf(file) {
  const pdfjs = await import('pdfjs')
  pdfjs.GlobalWorkerOptions.workerSrc =
    `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.mjs`
  return pdfjs.getDocument({ data: await file.arrayBuffer() }).promise
}

async function pdfText(doc) {
  const pages = []
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i)
    const content = await page.getTextContent()
    const line = content.items
      .map((it) => ('str' in it ? it.str : ''))
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim()
    if (line) pages.push(`--- p.${i} ---\n${line}`)
  }
  return pages.join('\n\n')
}

/** スキャンPDFをビジョンモデルに読ませるため、ページを画像にする */
async function pdfPageImages(doc, name) {
  const out = []
  const last = Math.min(doc.numPages, PDF_PAGE_IMAGE_LIMIT)
  for (let i = 1; i <= last; i++) {
    const page = await doc.getPage(i)
    const base = page.getViewport({ scale: 1 })
    const viewport = page.getViewport({ scale: Math.min(2, PDF_PAGE_IMAGE_WIDTH / base.width) })
    const canvas = document.createElement('canvas')
    canvas.width = Math.round(viewport.width)
    canvas.height = Math.round(viewport.height)
    await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise
    out.push({
      id: makeId(),
      name: `${name} p.${i}`,
      kind: 'image',
      chars: 0,
      tokens: IMAGE_TOKENS,
      text: '',
      dataUrl: canvas.toDataURL('image/jpeg', 0.85),
      bytes: 0,
    })
  }
  return out
}

async function parsePdf(file, name) {
  const doc = await openPdf(file)
  const text = await pdfText(doc)
  const chars = text.replace(/---\s*p\.\d+\s*---/g, '').trim().length

  if (chars >= PDF_TEXT_PER_PAGE * doc.numPages) return [pack(name, 'pdf', text)]

  const images = await pdfPageImages(doc, name)
  if (!images.length) throw new Error(`内容を取り出せませんでした: ${name}`)

  const notes = []
  if (text.trim()) notes.push(pack(name, 'pdf', text))
  if (doc.numPages > PDF_PAGE_IMAGE_LIMIT) {
    notes.push(
      pack(
        `${name} の注記`,
        'text',
        `${name} は全${doc.numPages}ページです。文字データが埋め込まれていないため、` +
          `先頭${PDF_PAGE_IMAGE_LIMIT}ページのみ画像として読み込みました。`,
      ),
    )
  }
  return [...notes, ...images]
}

async function parseDocx(file) {
  const mammoth = await loadScript(MAMMOTH_URL, 'mammoth')
  const { value } = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() })
  return value.trim()
}

async function parseXlsx(file) {
  const XLSX = await loadScript(XLSX_URL, 'XLSX')
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' })
  const out = []
  for (const name of wb.SheetNames) {
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]).trim()
    if (csv) out.push(`--- ${name} ---\n${csv}`)
  }
  return out.join('\n\n')
}

/** 戻りは常に配列。スキャンPDFは1ファイルが複数のページ画像になる */
export async function parseFile(file) {
  const name = file.name || `image-${Date.now()}.png`
  if (file.type.startsWith('image/')) return [await packImage(file, name)]
  if (/\.pdf$/i.test(name)) return parsePdf(file, name)
  if (/\.docx$/i.test(name)) return [pack(name, 'docx', await parseDocx(file))]
  if (/\.(xlsx|xlsm|xls)$/i.test(name)) return [pack(name, 'xlsx', await parseXlsx(file))]
  if (TEXT_EXT.test(name) || file.type.startsWith('text/')) {
    return [pack(name, 'text', await file.text())]
  }
  throw new Error(`未対応の形式です: ${name}`)
}

export function packPaste(text, index) {
  return pack(`貼り付けテキスト ${index}`, 'paste', text)
}

export function attachmentTag(kind) {
  if (kind === 'pdf') return 'PDF'
  if (kind === 'docx') return 'DOC'
  if (kind === 'xlsx') return 'XLS'
  if (kind === 'image') return 'IMG'
  return 'TXT'
}

export function formatBytes(n) {
  if (!n) return ''
  if (n < 1024) return `${n} B`
  if (n < 1024 * 1024) return `${Math.round(n / 1024)} KB`
  return `${(n / 1024 / 1024).toFixed(1)} MB`
}
