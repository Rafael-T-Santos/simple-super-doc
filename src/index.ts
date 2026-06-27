import JSZip from 'jszip'
import { DocxParseError, type DocxDocument } from './types.js'
import { parseRelationships } from './parser/relationships.js'
import { parseStyles } from './parser/styles.js'
import { parseNumbering, emptyNumbering } from './parser/numbering.js'
import { parseDocument } from './parser/document.js'
import { render as renderHtml } from './renderer/html.js'

export type { DocxDocument, Block, ParagraphBlock, TableBlock, TableCell, TableRow, TextRun, ImageRun, Run, ComputedStyle, ListRef } from './types.js'
export { DocxParseError } from './types.js'

function parsePageSize(xml: string): DocxDocument['pageSize'] {
  const twipsToPx = (s: string) => Math.round(parseFloat(s) * 96 / 1440)
  const pgSz = /<w:pgSz\b([^/]*)\/?>/.exec(xml)
  if (!pgSz) return undefined
  const attrs = pgSz[1]
  const w = /\bw:w="([\d.]+)"/.exec(attrs)?.[1]
  const h = /\bw:h="([\d.]+)"/.exec(attrs)?.[1]
  if (!w || !h) return undefined

  const pgMar = /<w:pgMar\b([^/]*)\/?>/.exec(xml)
  const ma = pgMar?.[1] ?? ''
  const top    = /\bw:top="([\d.]+)"/.exec(ma)?.[1]    ?? '1440'
  const right  = /\bw:right="([\d.]+)"/.exec(ma)?.[1]  ?? '1440'
  const bottom = /\bw:bottom="([\d.]+)"/.exec(ma)?.[1] ?? '1440'
  const left   = /\bw:left="([\d.]+)"/.exec(ma)?.[1]   ?? '1440'

  return {
    widthPx:  twipsToPx(w),
    heightPx: twipsToPx(h),
    marginPx: {
      top:    twipsToPx(top),
      right:  twipsToPx(right),
      bottom: twipsToPx(bottom),
      left:   twipsToPx(left),
    },
  }
}

async function readEntry(zip: JSZip, path: string, required: true): Promise<string>
async function readEntry(zip: JSZip, path: string, required: false): Promise<string | null>
async function readEntry(zip: JSZip, path: string, required: boolean): Promise<string | null> {
  const entry = zip.file(path)
  if (!entry) {
    if (required) throw new DocxParseError(`Missing required entry: ${path}`, 'MISSING_ENTRY')
    return null
  }
  return entry.async('string')
}

export async function parse(buffer: ArrayBuffer): Promise<DocxDocument> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch {
    throw new DocxParseError('Failed to load zip archive', 'INVALID_ZIP')
  }

  const [documentXml, stylesXml, relsXml, numberingXml] = await Promise.all([
    readEntry(zip, 'word/document.xml', true),
    readEntry(zip, 'word/styles.xml', true),
    readEntry(zip, 'word/_rels/document.xml.rels', false),
    readEntry(zip, 'word/numbering.xml', false),
  ])

  const relationshipMap = relsXml ? parseRelationships(relsXml) : {}
  const { styleMap, docDefaults } = parseStyles(stylesXml)
  const { abstractNumMap, numMap } = numberingXml
    ? parseNumbering(numberingXml)
    : emptyNumbering()

  const blocks = await parseDocument(documentXml, {
    styleMap,
    docDefaults,
    abstractNumMap,
    numMap,
    relationshipMap,
    zip,
  })

  return { blocks, pageSize: parsePageSize(documentXml) }
}

export function render(doc: DocxDocument, container: HTMLElement): void {
  renderHtml(doc, container)
}
