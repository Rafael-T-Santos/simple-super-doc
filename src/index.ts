import JSZip from 'jszip'
import { DocxParseError, type DocxDocument, type NoteEntry, type RenderOptions, type Section, type Block, type ParagraphBlock, type PageSize } from './types.js'
import { parseRelationships } from './parser/relationships.js'
import { parseStyles } from './parser/styles.js'
import { parseNumbering, emptyNumbering } from './parser/numbering.js'
import { parseDocument, parseNotesXml, parseFooterXml, parseHeaderXml, type ParseContext } from './parser/document.js'
import { render as renderHtml } from './renderer/html.js'

export type { DocxDocument, Block, ParagraphBlock, TableBlock, TableCell, TableRow, TextRun, ImageRun, Run, ComputedStyle, ListRef, NoteEntry, RenderOptions, TabStop, Section, PageSize } from './types.js'
export { DocxParseError } from './types.js'

function parsePageSize(xml: string): DocxDocument['pageSize'] {
  const twipsToPx = (s: string) => Math.round(parseFloat(s) * 96 / 1440)
  // Use the LAST section's sectPr — the body-level <w:sectPr> (textually last),
  // which describes the final/only section. Earlier sectPrs sit in paragraph
  // pPr and describe earlier sections (handled per-section via doc.sections).
  const lastSect = xml.lastIndexOf('<w:sectPr')
  const scope = lastSect !== -1 ? xml.slice(lastSect) : xml
  const pgSz = /<w:pgSz\b([^/]*)\/?>/.exec(scope) ?? /<w:pgSz\b([^/]*)\/?>/.exec(xml)
  if (!pgSz) return undefined
  const attrs = pgSz[1]
  const w = /\bw:w="([\d.]+)"/.exec(attrs)?.[1]
  const h = /\bw:h="([\d.]+)"/.exec(attrs)?.[1]
  if (!w || !h) return undefined

  const pgMar = /<w:pgMar\b([^/]*)\/?>/.exec(scope) ?? /<w:pgMar\b([^/]*)\/?>/.exec(xml)
  const ma = pgMar?.[1] ?? ''
  const top    = /\bw:top="([\d.]+)"/.exec(ma)?.[1]    ?? '1440'
  const right  = /\bw:right="([\d.]+)"/.exec(ma)?.[1]  ?? '1440'
  const bottom = /\bw:bottom="([\d.]+)"/.exec(ma)?.[1] ?? '1440'
  const left   = /\bw:left="([\d.]+)"/.exec(ma)?.[1]   ?? '1440'
  const footer = /\bw:footer="([\d.]+)"/.exec(ma)?.[1]
  const header = /\bw:header="([\d.]+)"/.exec(ma)?.[1]

  return {
    widthPx:  twipsToPx(w),
    heightPx: twipsToPx(h),
    marginPx: {
      top:    twipsToPx(top),
      right:  twipsToPx(right),
      bottom: twipsToPx(bottom),
      left:   twipsToPx(left),
    },
    ...(footer ? { footerPx: twipsToPx(footer) } : {}),
    ...(header ? { headerPx: twipsToPx(header) } : {}),
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

  const [documentXml, stylesXml, relsXml, numberingXml, footnotesXml, endnotesXml] = await Promise.all([
    readEntry(zip, 'word/document.xml', true),
    readEntry(zip, 'word/styles.xml', true),
    readEntry(zip, 'word/_rels/document.xml.rels', false),
    readEntry(zip, 'word/numbering.xml', false),
    readEntry(zip, 'word/footnotes.xml', false),
    readEntry(zip, 'word/endnotes.xml', false),
  ])

  const relationshipMap = relsXml ? parseRelationships(relsXml) : {}
  const { styleMap, docDefaults } = parseStyles(stylesXml)
  const { abstractNumMap, numMap } = numberingXml
    ? parseNumbering(numberingXml)
    : emptyNumbering()

  const ctx: ParseContext = {
    styleMap,
    docDefaults,
    abstractNumMap,
    numMap,
    relationshipMap,
    zip,
    footnoteRefs: [],
    endnoteRefs: [],
  }

  // Parse the body first so footnote/endnote references are numbered in order.
  const blocks = await parseDocument(documentXml, ctx)

  // Resolve only the notes actually referenced, in reference order.
  const footnotes = await resolveNotes(footnotesXml, 'footnote', ctx.footnoteRefs, ctx)
  const endnotes = await resolveNotes(endnotesXml, 'endnote', ctx.endnoteRefs, ctx)

  // Default page footer / header (w:footerReference / w:headerReference type="default").
  const footer = await resolveFooter(documentXml, relationshipMap, zip, ctx)
  const header = await resolveHeader(documentXml, relationshipMap, zip, ctx)

  const pageSize = parsePageSize(documentXml)
  const sections = pageSize ? buildSections(blocks, pageSize) : []

  return {
    blocks,
    pageSize,
    ...(footnotes.length ? { footnotes } : {}),
    ...(endnotes.length ? { endnotes } : {}),
    ...(footer && footer.length ? { footer } : {}),
    ...(header && header.length ? { header } : {}),
    ...(sections.length > 1 ? { sections } : {}),
  }
}

// Split the flat block list into sections at paragraphs tagged with a
// sectionPageSize (a w:sectPr in their pPr). The trailing run of blocks forms
// the final section, sized by the body-level sectPr (bodyPageSize). The tag is
// transient and stripped here. Returns one section for a single-section doc.
function buildSections(blocks: Block[], bodyPageSize: PageSize): Section[] {
  const sections: Section[] = []
  let cur: Block[] = []
  for (const b of blocks) {
    cur.push(b)
    const tagged = b.type === 'paragraph' ? (b as ParagraphBlock).sectionPageSize : undefined
    if (tagged) {
      sections.push({ blocks: cur, pageSize: tagged })
      delete (b as ParagraphBlock).sectionPageSize
      cur = []
    }
  }
  if (cur.length) sections.push({ blocks: cur, pageSize: bodyPageSize })
  return sections
}

async function resolveFooter(
  documentXml: string,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
): Promise<DocxDocument['footer']> {
  return resolvePart(documentXml, relationshipMap, zip, ctx, 'footer', parseFooterXml)
}

async function resolveHeader(
  documentXml: string,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
): Promise<DocxDocument['header']> {
  return resolvePart(documentXml, relationshipMap, zip, ctx, 'header', parseHeaderXml)
}

// Resolve a default header/footer reference (w:headerReference / w:footerReference
// w:type="default") to its parsed blocks. Attribute order varies, so scan tags.
async function resolvePart(
  documentXml: string,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
  kind: 'header' | 'footer',
  parseXml: (xml: string, ctx: ParseContext) => Promise<DocxDocument['blocks']>,
): Promise<DocxDocument['blocks'] | undefined> {
  const ref = new RegExp(`<w:${kind}Reference\\b[^>]*>`, 'g')
  let m: RegExpExecArray | null
  let rId: string | undefined
  while ((m = ref.exec(documentXml)) !== null) {
    const tag = m[0]
    if (!/w:type="default"/.test(tag)) continue
    rId = /r:id="([^"]+)"/.exec(tag)?.[1]
    break
  }
  if (!rId || !relationshipMap[rId]) return undefined
  const xml = await readEntry(zip, `word/${relationshipMap[rId].target}`, false)
  return xml ? parseXml(xml, ctx) : undefined
}

async function resolveNotes(
  xml: string | null,
  kind: 'footnote' | 'endnote',
  refIds: string[],
  ctx: ParseContext,
): Promise<NoteEntry[]> {
  if (!xml || refIds.length === 0) return []
  const map = await parseNotesXml(xml, kind, ctx)
  const entries: NoteEntry[] = []
  refIds.forEach((id, i) => {
    const noteBlocks = map.get(id)
    if (noteBlocks) entries.push({ number: i + 1, blocks: noteBlocks })
  })
  return entries
}

export function render(doc: DocxDocument, container: HTMLElement, options?: RenderOptions): void {
  renderHtml(doc, container, options)
}
