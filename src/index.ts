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

  const [documentXml, stylesXml, relsXml, numberingXml, footnotesXml, endnotesXml, settingsXml] = await Promise.all([
    readEntry(zip, 'word/document.xml', true),
    readEntry(zip, 'word/styles.xml', true),
    readEntry(zip, 'word/_rels/document.xml.rels', false),
    readEntry(zip, 'word/numbering.xml', false),
    readEntry(zip, 'word/footnotes.xml', false),
    readEntry(zip, 'word/endnotes.xml', false),
    readEntry(zip, 'word/settings.xml', false),
  ])

  const relationshipMap = relsXml ? parseRelationships(relsXml) : {}
  const { styleMap, docDefaults, tableBorderMap } = parseStyles(stylesXml)
  const { abstractNumMap, numMap } = numberingXml
    ? parseNumbering(numberingXml)
    : emptyNumbering()

  const ctx: ParseContext = {
    styleMap,
    docDefaults,
    tableBorderMap,
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

  // Page footer / header. The default applies to every page; first-page and
  // even-page variants are used when w:titlePg (section) / w:evenAndOddHeaders
  // (settings) opt in. Letterhead, books and reports rely on these.
  const footer = await resolveFooter(documentXml, relationshipMap, zip, ctx)
  const header = await resolveHeader(documentXml, relationshipMap, zip, ctx)
  const headerFirst = await resolveHeader(documentXml, relationshipMap, zip, ctx, 'first')
  const headerEven = await resolveHeader(documentXml, relationshipMap, zip, ctx, 'even')
  const footerFirst = await resolveFooter(documentXml, relationshipMap, zip, ctx, 'first')
  const footerEven = await resolveFooter(documentXml, relationshipMap, zip, ctx, 'even')
  // w:titlePg lives in the body (last) sectPr; w:evenAndOddHeaders in settings.xml.
  const bodySectScope = documentXml.slice(documentXml.lastIndexOf('<w:sectPr'))
  const titlePg = /<w:titlePg\b[^>]*\/?>/.test(bodySectScope) && !/<w:titlePg\b[^>]*w:val="(?:0|false|off)"/.test(bodySectScope)
  const evenAndOddHeaders = !!settingsXml && /<w:evenAndOddHeaders\b[^>]*\/?>/.test(settingsXml) &&
    !/<w:evenAndOddHeaders\b[^>]*w:val="(?:0|false|off)"/.test(settingsXml)

  const pageSize = parsePageSize(documentXml)
  // Resolve a header/footer relationship id to its parsed blocks (per-section).
  const resolveRef = (rId: string | undefined, kind: 'header' | 'footer') =>
    resolveRid(rId, relationshipMap, zip, ctx, kind === 'header' ? parseHeaderXml : parseFooterXml)
  const sections = pageSize
    ? await buildSections(blocks, pageSize, bodySectPrRefs(documentXml), resolveRef, header, footer)
    : []

  return {
    blocks,
    pageSize,
    ...(footnotes.length ? { footnotes } : {}),
    ...(endnotes.length ? { endnotes } : {}),
    ...(footer && footer.length ? { footer } : {}),
    ...(header && header.length ? { header } : {}),
    ...(headerFirst && headerFirst.length ? { headerFirst } : {}),
    ...(headerEven && headerEven.length ? { headerEven } : {}),
    ...(footerFirst && footerFirst.length ? { footerFirst } : {}),
    ...(footerEven && footerEven.length ? { footerEven } : {}),
    ...(titlePg ? { titlePg } : {}),
    ...(evenAndOddHeaders ? { evenAndOddHeaders } : {}),
    ...(sections.length > 1 ? { sections } : {}),
  }
}

type SectionRefs = { headerRId?: string; footerRId?: string }

// Split the flat block list into sections at paragraphs tagged with a
// sectionPageSize (a w:sectPr in their pPr). The trailing run of blocks forms
// the final section, sized by the body-level sectPr (bodyPageSize, bodyRefs).
// Each section's header/footer references are resolved to blocks. A section
// that does not declare its own reference inherits the PREVIOUS section's
// (OOXML semantics), seeded by the document-level default (docHeader/docFooter)
// for leading sections that declare none. Transient tags are stripped here.
// Returns one section for a single-section doc.
async function buildSections(
  blocks: Block[],
  bodyPageSize: PageSize,
  bodyRefs: SectionRefs,
  resolveRef: (rId: string | undefined, kind: 'header' | 'footer') => Promise<Block[] | undefined>,
  docHeader: Block[] | undefined,
  docFooter: Block[] | undefined,
): Promise<Section[]> {
  const raw: { blocks: Block[]; pageSize: PageSize; refs: SectionRefs }[] = []
  let cur: Block[] = []
  for (const b of blocks) {
    cur.push(b)
    const pb = b.type === 'paragraph' ? (b as ParagraphBlock) : undefined
    const tagged = pb?.sectionPageSize
    if (tagged) {
      raw.push({ blocks: cur, pageSize: tagged, refs: pb!.sectionRefs ?? {} })
      delete pb!.sectionPageSize
      delete pb!.sectionRefs
      cur = []
    }
  }
  if (cur.length) raw.push({ blocks: cur, pageSize: bodyPageSize, refs: bodyRefs })

  // Walk sections in document order, carrying the last resolved header/footer
  // forward. A section with its own reference overrides the carry (and becomes
  // what later sections inherit); one without keeps the previous section's.
  const sections: Section[] = []
  let carriedHeader = docHeader
  let carriedFooter = docFooter
  for (const s of raw) {
    if (s.refs.headerRId) carriedHeader = (await resolveRef(s.refs.headerRId, 'header')) ?? carriedHeader
    if (s.refs.footerRId) carriedFooter = (await resolveRef(s.refs.footerRId, 'footer')) ?? carriedFooter
    const header = carriedHeader
    const footer = carriedFooter
    sections.push({
      blocks: s.blocks,
      pageSize: s.pageSize,
      ...(header && header.length ? { header } : {}),
      ...(footer && footer.length ? { footer } : {}),
    })
  }
  return sections
}

// Default header/footer relationship ids from the body-level sectPr (the
// textually last <w:sectPr>, which describes the final section).
function bodySectPrRefs(documentXml: string): SectionRefs {
  const lastSect = documentXml.lastIndexOf('<w:sectPr')
  if (lastSect === -1) return {}
  const end = documentXml.indexOf('</w:sectPr>', lastSect)
  const scope = end !== -1 ? documentXml.slice(lastSect, end) : documentXml.slice(lastSect)
  return {
    ...(defaultRefRId(scope, 'header') ? { headerRId: defaultRefRId(scope, 'header') } : {}),
    ...(defaultRefRId(scope, 'footer') ? { footerRId: defaultRefRId(scope, 'footer') } : {}),
  }
}

// First default-type header/footer reference rId within an XML scope.
function defaultRefRId(scope: string, kind: 'header' | 'footer'): string | undefined {
  const re = new RegExp(`<w:${kind}Reference\\b[^>]*>`, 'g')
  let m: RegExpExecArray | null
  while ((m = re.exec(scope)) !== null) {
    if (/w:type="default"/.test(m[0])) return /r:id="([^"]+)"/.exec(m[0])?.[1]
  }
  return undefined
}

async function resolveFooter(
  documentXml: string,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
  type: 'default' | 'first' | 'even' = 'default',
): Promise<DocxDocument['footer']> {
  return resolvePart(documentXml, relationshipMap, zip, ctx, 'footer', parseFooterXml, type)
}

async function resolveHeader(
  documentXml: string,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
  type: 'default' | 'first' | 'even' = 'default',
): Promise<DocxDocument['header']> {
  return resolvePart(documentXml, relationshipMap, zip, ctx, 'header', parseHeaderXml, type)
}

// Resolve a header/footer reference of the given w:type (default/first/even) to
// its parsed blocks. Attribute order varies, so scan tags.
async function resolvePart(
  documentXml: string,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
  kind: 'header' | 'footer',
  parseXml: (xml: string, ctx: ParseContext) => Promise<DocxDocument['blocks']>,
  type: 'default' | 'first' | 'even' = 'default',
): Promise<DocxDocument['blocks'] | undefined> {
  const ref = new RegExp(`<w:${kind}Reference\\b[^>]*>`, 'g')
  let m: RegExpExecArray | null
  let rId: string | undefined
  while ((m = ref.exec(documentXml)) !== null) {
    const tag = m[0]
    if (!new RegExp(`w:type="${type}"`).test(tag)) continue
    rId = /r:id="([^"]+)"/.exec(tag)?.[1]
    break
  }
  return resolveRid(rId, relationshipMap, zip, ctx, parseXml)
}

// Resolve a single header/footer relationship id to its parsed blocks.
async function resolveRid(
  rId: string | undefined,
  relationshipMap: ReturnType<typeof parseRelationships>,
  zip: JSZip,
  ctx: ParseContext,
  parseXml: (xml: string, ctx: ParseContext) => Promise<DocxDocument['blocks']>,
): Promise<DocxDocument['blocks'] | undefined> {
  if (!rId || !relationshipMap[rId]) return undefined
  const target = relationshipMap[rId].target
  const xml = await readEntry(zip, `word/${target}`, false)
  if (!xml) return undefined
  // A header/footer part has its OWN relationships file (e.g.
  // word/_rels/header1.xml.rels). Images and hyperlinks inside the part use that
  // rId namespace, not the document's, so parse the part with its own map
  // (falling back to the document's when the part declares none).
  const base = target.replace(/^.*\//, '')
  const partRelsXml = await readEntry(zip, `word/_rels/${base}.rels`, false)
  const partCtx: ParseContext = partRelsXml
    ? { ...ctx, relationshipMap: parseRelationships(partRelsXml) }
    : ctx
  return parseXml(xml, partCtx)
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
