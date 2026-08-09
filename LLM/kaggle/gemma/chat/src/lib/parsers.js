import { estimateTokens } from './tokens.js'
import { loadScript } from './dom.js'

/** ペーストをカード化する閾値 */
export const PASTE_THRESHOLD = 2000

const PDFJS_VERSION = '4.10.38'
const MAMMOTH_URL = 'https://cdn.jsdelivr.net/npm/mammoth@1.9.0/mammoth.browser.min.js'
const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'

const TEXT_EXT =
  /\.(txt|md|markdown|csv|tsv|json|ya?ml|log|ts|tsx|js|jsx|py|java|kt|sql|html|css|xml)$/i

function makeId() {
  return Math.random().toString(36).slice(2, 10)
}

function pack(name, kind, text) {
  return { id: makeId(), name, kind, chars: text.length, tokens: estimateTokens(text), text }
}

async function parsePdf(file) {
  const pdfjs = await import('pdfjs')
  pdfjs.GlobalWorkerOptions.workerSrc =
    `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.mjs`

  const doc = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise
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

export async function parseFile(file) {
  const name = file.name
  if (/\.pdf$/i.test(name)) return pack(name, 'pdf', await parsePdf(file))
  if (/\.docx$/i.test(name)) return pack(name, 'docx', await parseDocx(file))
  if (/\.(xlsx|xlsm|xls)$/i.test(name)) return pack(name, 'xlsx', await parseXlsx(file))
  if (TEXT_EXT.test(name) || file.type.startsWith('text/')) {
    return pack(name, 'text', await file.text())
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
  return 'TXT'
}
