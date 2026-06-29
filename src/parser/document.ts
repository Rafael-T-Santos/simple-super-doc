import { XMLParser } from 'fast-xml-parser'
import type JSZip from 'jszip'
import type {
  Block, ParagraphBlock, TableBlock, TableCell, TableRow,
  TextRun, ImageRun, Run, ComputedStyle, ListRef,
} from '../types.js'
import type { StyleMap } from './styles.js'
import { extractRPr, extractPPr, extractMarkRPr } from './styles.js'
import type { AbstractNumMap, NumMap } from './numbering.js'
import type { RelationshipMap } from './relationships.js'
import { resolveVMerge, type RawCell } from './table.js'
import { resolveImage } from './images.js'

const parser = new XMLParser({
  removeNSPrefix: true,
  attributeNamePrefix: '',
  ignoreAttributes: false,
  parseAttributeValue: false,
  // Preserve significant whitespace in <w:t xml:space="preserve"> runs. The
  // default (trimValues: true) strips leading/trailing spaces, which silently
  // joins words across runs ("WHEREAS the" -> "WHEREASthe").
  trimValues: false,
  isArray: (name) =>
    ['p', 'r', 'tbl', 'tr', 'tc', 'style', 'abstractNum', 'num', 'lvl', 'lvlOverride',
     'hyperlink', 'bookmarkStart', 'ins', 'del', 'footnote', 'endnote'].includes(name),
})

export type ParseContext = {
  styleMap: StyleMap
  docDefaults: ComputedStyle
  abstractNumMap: AbstractNumMap
  numMap: NumMap
  relationshipMap: RelationshipMap
  zip: JSZip
  // Footnote/endnote ids in document order; index+1 is the printed number.
  footnoteRefs: string[]
  endnoteRefs: string[]
}

function getVal(node: unknown): string | undefined {
  if (node == null) return undefined
  if (typeof node === 'object') return (node as Record<string, string>).val
  return String(node)
}

function resolveNamedStyle(
  styleId: string | undefined,
  styleMap: StyleMap,
  docDefaults: ComputedStyle,
): ComputedStyle {
  if (!styleId) return { ...docDefaults }
  const named = styleMap[styleId]
  if (!named) {
    console.warn(`[simple-super-doc] unknown style "${styleId}" — falling back to docDefaults`)
    return { ...docDefaults }
  }
  return Object.assign({}, docDefaults, named)
}

function resolveListRef(
  pPr: Record<string, unknown> | undefined,
  ctx: ParseContext,
): ListRef | undefined {
  if (!pPr) return undefined
  const numPr = pPr.numPr as Record<string, unknown> | undefined
  if (!numPr) return undefined

  const numId = getVal(numPr.numId)
  const ilvlRaw = getVal(numPr.ilvl)
  const ilvl = parseInt(ilvlRaw ?? '0', 10)

  if (!numId || numId === '0') return undefined

  const numEntry = ctx.numMap[numId]
  if (!numEntry) return undefined

  const levelMap = ctx.abstractNumMap[numEntry.abstractNumId]
  if (!levelMap) return undefined

  const levelInfo = levelMap[ilvl] ?? levelMap[0]
  if (!levelInfo) return undefined

  const start = numEntry.startOverride ?? levelInfo.start
  const ordered = levelInfo.format !== 'bullet' && levelInfo.format !== 'none'

  return { numId, ilvl, ordered, start, format: levelInfo.format }
}

async function parseRun(
  r: Record<string, unknown>,
  paraStyle: ComputedStyle,
  ctx: ParseContext,
  href?: string,
): Promise<Run | null> {
  const rPr = r.rPr as Record<string, unknown> | undefined

  // A PAGE field becomes a page-number marker (resolved per page at render time).
  if ('instrText' in r) {
    const instr = (typeof r.instrText === 'object' && r.instrText !== null
      ? String((r.instrText as Record<string, string>)['#text'] ?? '')
      : String(r.instrText)).trim().toUpperCase()
    if (instr === 'PAGE') {
      return { type: 'run', text: '', style: Object.assign({}, paraStyle, extractRPr(rPr)), pageNumber: true }
    }
    return null // other field codes are skipped
  }

  // Skip field characters and a note's own auto-number marker.
  if ('fldChar' in r) return null
  if ('footnoteRef' in r || 'endnoteRef' in r) return null

  // w:br: a soft line break becomes a <br>; a page break has no place in the
  // continuous flow model, so it's dropped (paragraph-level pageBreakBefore is
  // handled separately).
  let lineBreak = false
  if ('br' in r) {
    const br = r.br as Record<string, string> | string | undefined
    const brType = typeof br === 'object' && br !== null ? br.type : undefined
    if (brType === 'page' || brType === 'column') return null
    lineBreak = true
  }

  // Character style at cascade level 2 for this specific run
  const rStyleId = getVal(rPr?.rStyle as unknown)
  const charStyle = rStyleId ? (ctx.styleMap[rStyleId] ?? {}) : {}

  const runStyle = Object.assign({}, paraStyle, charStyle, extractRPr(rPr))

  // Footnote/endnote reference: emit a numbered marker run. The number is the
  // position of this reference among same-type references in document order.
  const noteEl = (r.footnoteReference ?? r.endnoteReference) as Record<string, string> | undefined
  if (noteEl !== undefined) {
    const isFoot = 'footnoteReference' in r
    const id = String(noteEl.id ?? '')
    const refs = isFoot ? ctx.footnoteRefs : ctx.endnoteRefs
    refs.push(id)
    const marker: TextRun = {
      type: 'run',
      text: String(refs.length),
      style: runStyle,
      noteRef: { type: isFoot ? 'footnote' : 'endnote', number: refs.length },
    }
    return marker
  }

  // Image via drawing
  if ('drawing' in r) {
    const drawing = r.drawing as Record<string, unknown>
    const anchor = drawing.anchor as Record<string, unknown> | undefined
    const inline = drawing.inline as Record<string, unknown> | undefined
    const isPageBackground = !inline && !!anchor && String(anchor.behindDoc ?? '0') === '1'

    const drawingEl = inline ?? anchor
    if (!drawingEl) return null

    const extent = drawingEl.extent as Record<string, string> | undefined
    const cx = parseInt(extent?.cx ?? '0', 10)
    const cy = parseInt(extent?.cy ?? '0', 10)

    // Search for a:blip r:embed anywhere in the drawing subtree
    const rId = findBlipEmbed(drawingEl)

    if (!rId) return null

    const resolved = await resolveImage(rId, ctx.relationshipMap, ctx.zip)
    if (!resolved) return null

    const img: ImageRun = {
      type: 'image',
      src: resolved.src,
      widthPx: Math.round(cx / 9525),
      heightPx: Math.round(cy / 9525),
      ...(isPageBackground ? { isPageBackground: true } : {}),
      ...(href ? { href } : {}),
    }
    return img
  }

  // Text run
  const tNode = r.t
  let text = ''
  if (typeof tNode === 'string') {
    text = tNode
  } else if (typeof tNode === 'object' && tNode !== null) {
    const t = tNode as Record<string, unknown>
    text = String(t['#text'] ?? t._ ?? '')
  } else if (tNode != null) {
    text = String(tNode)
  }

  const textRun: TextRun = {
    type: 'run', text, style: runStyle,
    ...(href ? { href } : {}),
    ...(lineBreak ? { lineBreak: true } : {}),
  }
  return textRun
}

// Resolve a <w:hyperlink>'s destination: an external URL via r:id (relationships)
// or an internal bookmark via w:anchor (rendered as a #fragment).
function resolveHyperlinkHref(
  hl: Record<string, unknown>,
  ctx: ParseContext,
): string | undefined {
  const rId = hl.id
  if (typeof rId === 'string') {
    const rel = ctx.relationshipMap[rId]
    if (rel?.target) return rel.target
  }
  const anchor = hl.anchor
  if (typeof anchor === 'string' && anchor) return `#${anchor}`
  return undefined
}

// Recursively search for a:blip embed attribute in a nested object
function findBlipEmbed(node: unknown): string | undefined {
  if (!node || typeof node !== 'object') return undefined
  const obj = node as Record<string, unknown>
  if ('embed' in obj && typeof obj.embed === 'string') return obj.embed
  for (const v of Object.values(obj)) {
    const found = findBlipEmbed(v)
    if (found) return found
  }
  return undefined
}

async function parseParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
  rawXml?: string,
): Promise<ParagraphBlock> {
  const pPr = p.pPr as Record<string, unknown> | undefined
  const pStyleId = getVal((pPr?.pStyle) as unknown)

  const namedStyle = resolveNamedStyle(pStyleId, ctx.styleMap, ctx.docDefaults)
  const pPrStyle = extractPPr(pPr)
  const paraStyle = Object.assign({}, namedStyle, pPrStyle)

  const list = resolveListRef(pPr, ctx)

  // w:pageBreakBefore forces a new page (unless explicitly disabled with val=0).
  let pageBreakBefore = false
  if (pPr && 'pageBreakBefore' in pPr) {
    const val = getVal(pPr.pageBreakBefore)
    pageBreakBefore = !(val === '0' || val === 'false' || val === 'off')
  }

  // Collect runs from: direct r[], w:ins child r[], w:hyperlink child r[],
  // interleaved in document order (recovered from rawXml) so a hyperlink or
  // insertion in the middle of a paragraph isn't reordered to the end. Hyperlink
  // runs carry the resolved href so the renderer can wrap them in <a>.
  type RunInput = { r: Record<string, unknown>; href?: string }
  const directRuns = (p.r ?? []) as Record<string, unknown>[]
  const insList = (p.ins ?? []) as Record<string, unknown>[]
  const hyperlinks = (p.hyperlink ?? []) as Record<string, unknown>[]

  const insInputs = (ins: Record<string, unknown>): RunInput[] => {
    const inner = ins.r
    if (Array.isArray(inner)) return inner.map(r => ({ r }))
    return inner ? [{ r: inner as Record<string, unknown> }] : []
  }
  const hlInputs = (hl: Record<string, unknown>): RunInput[] => {
    const href = resolveHyperlinkHref(hl, ctx)
    const inner = hl.r
    if (Array.isArray(inner)) return inner.map(r => ({ r, href }))
    return inner ? [{ r: inner as Record<string, unknown>, href }] : []
  }

  const order = rawXml ? getRunOrder(rawXml) : []
  // Use the recovered order only if it accounts for every grouped child; else
  // fall back to the safe (append-grouped) order.
  const counts = { r: 0, hyperlink: 0, ins: 0 }
  for (const t of order) counts[t]++
  const orderOk = order.length > 0 &&
    counts.r === directRuns.length && counts.hyperlink === hyperlinks.length && counts.ins === insList.length

  const inputs: RunInput[] = []
  if (orderOk) {
    let ri = 0, hi = 0, ii = 0
    for (const t of order) {
      if (t === 'r') inputs.push({ r: directRuns[ri++] })
      else if (t === 'hyperlink') inputs.push(...hlInputs(hyperlinks[hi++]))
      else inputs.push(...insInputs(insList[ii++]))
    }
  } else {
    for (const r of directRuns) inputs.push({ r })
    for (const ins of insList) inputs.push(...insInputs(ins))
    for (const hl of hyperlinks) inputs.push(...hlInputs(hl))
  }

  const runs: Run[] = []
  for (const { r, href } of inputs) {
    const run = await parseRun(r, paraStyle, ctx, href)
    if (run !== null) runs.push(run)
  }

  // For an empty paragraph the paragraph-mark run properties (pPr > rPr) set the
  // line's font/metrics; for a paragraph with text they apply only to the mark
  // glyph and must NOT bold/resize the runs, so we fold them into the block
  // style only when there is no visible run.
  const hasVisible = runs.some(r => r.type === 'image' || (r as TextRun).text.trim().length > 0)
  const blockStyle = hasVisible ? paraStyle : Object.assign({}, paraStyle, extractMarkRPr(pPr))

  return {
    type: 'paragraph',
    style: blockStyle,
    runs,
    ...(list ? { list } : {}),
    ...(pageBreakBefore ? { pageBreakBefore: true } : {}),
  }
}

async function parseTable(
  tbl: Record<string, unknown>,
  ctx: ParseContext,
): Promise<TableBlock> {
  const rows = (tbl.tr ?? []) as Record<string, unknown>[]

  // Pass 1: collect raw cells
  const rawGrid: RawCell[][] = []
  for (const tr of rows) {
    const cells = (tr.tc ?? []) as Record<string, unknown>[]
    const rawRow: RawCell[] = []
    for (const tc of cells) {
      const tcPr = tc.tcPr as Record<string, unknown> | undefined
      const colSpan = parseInt(getVal(tcPr?.gridSpan as unknown) ?? '1', 10) || 1

      const vMergeNode = tcPr?.vMerge
      let vMerge: 'restart' | 'continue' | 'none' = 'none'
      if (vMergeNode !== undefined) {
        const val = getVal(vMergeNode)
        vMerge = val === 'restart' ? 'restart' : 'continue'
      }

      const tcShd = tcPr?.shd as Record<string, string> | undefined
      const cellBg = tcShd?.fill && tcShd.fill !== 'auto' ? tcShd.fill : undefined

      rawRow.push({ colSpan, vMerge, rawData: tc, backgroundColor: cellBg })
    }
    rawGrid.push(rawRow)
  }

  // Pass 2: resolve vMerge
  const resolved = resolveVMerge(rawGrid)

  const irRows: TableRow[] = []
  for (const resolvedRow of resolved) {
    const irCells: TableCell[] = []
    for (const cell of resolvedRow) {
      const tc = cell.rawData as Record<string, unknown>
      const cellBlocks = await parseBlockContainer(tc, ctx)
      const irCell: TableCell = { rowSpan: cell.rowSpan, colSpan: cell.colSpan, blocks: cellBlocks }
      if (cell.backgroundColor) irCell.backgroundColor = cell.backgroundColor
      irCells.push(irCell)
    }
    irRows.push({ cells: irCells })
  }

  // Column widths from w:tblGrid (so columns match the document instead of being
  // sized by content, which would change wrapping and how many rows fit a page).
  const grid = tbl.tblGrid as Record<string, unknown> | undefined
  const colsNode = grid?.gridCol
  const colArr = Array.isArray(colsNode) ? colsNode : colsNode ? [colsNode] : []
  const columnWidths = colArr.map(c => {
    const w = (c as Record<string, string>).w
    return w != null ? Math.round((parseFloat(w) * 96) / 1440) : 0
  })

  // Cell padding from w:tblCellMar (table default) or the first cell's w:tcMar.
  const tblPr = tbl.tblPr as Record<string, unknown> | undefined
  const firstTc = ((rows[0]?.tc ?? []) as Record<string, unknown>[])[0]
  const firstTcPr = firstTc?.tcPr as Record<string, unknown> | undefined
  const marNode = (tblPr?.tblCellMar ?? firstTcPr?.tcMar) as Record<string, unknown> | undefined
  const cellPadding = marginToPx(marNode)

  return {
    type: 'table',
    rows: irRows,
    ...(columnWidths.some(w => w > 0) ? { columnWidths } : {}),
    ...(cellPadding ? { cellPadding } : {}),
  }
}

// Convert an OOXML margin node (top/right/bottom/left with w:w in twips) to px.
function marginToPx(
  node: Record<string, unknown> | undefined,
): { top: number; right: number; bottom: number; left: number } | undefined {
  if (!node) return undefined
  const side = (s: unknown): number => {
    const w = (s as Record<string, string> | undefined)?.w
    return w != null ? Math.round((parseFloat(w) * 96) / 1440) : 0
  }
  return { top: side(node.top), right: side(node.right), bottom: side(node.bottom), left: side(node.left) }
}

// Raw XML of each top-level (non-table-nested) paragraph, in document order.
// Used to recover run order within a paragraph (see getRunOrder); aligns 1:1
// with the parser's grouped body.p[] array.
function extractParagraphChunks(xml: string): string[] {
  const bodyStart = xml.indexOf('<w:body>')
  const bodyEnd = xml.lastIndexOf('</w:body>')
  if (bodyStart === -1 || bodyEnd === -1) return []
  const body = xml.slice(bodyStart + 8, bodyEnd)

  const chunks: string[] = []
  const re = /<(\/?)w:(p|tbl)[\s>\/]/g
  let tblDepth = 0
  let pStart = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(body)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2]
    if (tag === 'tbl') {
      if (isClose) tblDepth = Math.max(0, tblDepth - 1)
      else tblDepth++
      continue
    }
    if (tblDepth !== 0) continue // paragraph inside a table — not top-level
    if (!isClose) {
      const gt = body.indexOf('>', m.index)
      if (gt > 0 && body[gt - 1] === '/') chunks.push(body.slice(m.index, gt + 1)) // <w:p/>
      else if (pStart === -1) pStart = m.index
    } else if (pStart !== -1) {
      const gt = body.indexOf('>', m.index)
      chunks.push(body.slice(pStart, gt + 1))
      pStart = -1
    }
  }
  return chunks
}

// Document order of a paragraph's content children (runs, hyperlinks, tracked
// insertions). fast-xml-parser groups these by tag, so a hyperlink/insertion in
// the MIDDLE of a paragraph would otherwise be reordered to the end.
function getRunOrder(paraXml: string): Array<'r' | 'hyperlink' | 'ins'> {
  let body = paraXml
  const pprEnd = body.indexOf('</w:pPr>')
  if (pprEnd !== -1) body = body.slice(pprEnd + 8) // skip paragraph properties

  const order: Array<'r' | 'hyperlink' | 'ins'> = []
  const re = /<(\/?)w:(r|hyperlink|ins)[\s>\/]/g
  let depth = 0    // inside a hyperlink/ins wrapper
  let inRun = false // inside a <w:r> (skip its rPr-change ins markers)
  let m: RegExpExecArray | null
  while ((m = re.exec(body)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2] as 'r' | 'hyperlink' | 'ins'
    if (tag === 'r') {
      if (isClose) inRun = false
      else {
        if (depth === 0) order.push('r')
        inRun = true
      }
      continue
    }
    if (inRun) continue // an <w:ins> inside a run is an rPr change, not content
    if (isClose) depth = Math.max(0, depth - 1)
    else {
      if (depth === 0) order.push(tag)
      depth++
    }
  }
  return order
}

// Scan raw body XML to determine document order of top-level p and tbl elements.
// fast-xml-parser groups elements by tag name, losing cross-type ordering.
function getBodyBlockOrder(xml: string): Array<'p' | 'tbl'> {
  const bodyStart = xml.indexOf('<w:body>')
  const bodyEnd = xml.lastIndexOf('</w:body>')
  if (bodyStart === -1 || bodyEnd === -1) return []
  const body = xml.slice(bodyStart + 8, bodyEnd)

  const order: Array<'p' | 'tbl'> = []
  // Match opening/closing w:p and w:tbl tags ([\s>\/] excludes w:pPr, w:pStyle etc.)
  const re = /<(\/?)w:(p|tbl)[\s>\/]/g
  let tblDepth = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(body)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2] as 'p' | 'tbl'
    if (isClose) {
      if (tag === 'tbl') tblDepth = Math.max(0, tblDepth - 1)
    } else {
      if (tblDepth === 0) order.push(tag)
      if (tag === 'tbl') tblDepth++
    }
  }
  return order
}

// Parse all paragraph and table children of a container node (body, cell, etc.)
// paraXmls (when provided, aligned 1:1 with container.p[]) lets paragraphs
// recover their run order from raw XML.
async function parseBlockContainer(
  container: Record<string, unknown>,
  ctx: ParseContext,
  order?: Array<'p' | 'tbl'>,
  paraXmls?: string[],
): Promise<Block[]> {
  const ps = (container.p ?? []) as Record<string, unknown>[]
  const tbls = (container.tbl ?? []) as Record<string, unknown>[]
  const xmls = paraXmls && paraXmls.length === ps.length ? paraXmls : undefined

  if (!order || order.length === 0) {
    const blocks: Block[] = []
    for (let i = 0; i < ps.length; i++) blocks.push(await parseParagraph(ps[i], ctx, xmls?.[i]))
    for (const tbl of tbls) blocks.push(await parseTable(tbl, ctx))
    return blocks
  }

  const blocks: Block[] = []
  let pIdx = 0, tblIdx = 0
  for (const type of order) {
    if (type === 'p' && pIdx < ps.length) {
      blocks.push(await parseParagraph(ps[pIdx], ctx, xmls?.[pIdx]))
      pIdx++
    } else if (type === 'tbl' && tblIdx < tbls.length) {
      blocks.push(await parseTable(tbls[tblIdx++], ctx))
    }
  }
  return blocks
}

export async function parseDocument(xml: string, ctx: ParseContext): Promise<Block[]> {
  const doc = parser.parse(xml) as Record<string, unknown>
  const body = (doc?.document as Record<string, unknown>)?.body as Record<string, unknown>
  if (!body) return []
  const order = getBodyBlockOrder(xml)
  const paraXmls = extractParagraphChunks(xml)
  return parseBlockContainer(body, ctx, order, paraXmls)
}

// Parse a footer part (footerN.xml, root <w:ftr>) into content blocks.
export async function parseFooterXml(xml: string, ctx: ParseContext): Promise<Block[]> {
  const doc = parser.parse(xml) as Record<string, unknown>
  const ftr = doc?.ftr as Record<string, unknown> | undefined
  if (!ftr) return []
  return parseBlockContainer(ftr, ctx)
}

// Parse footnotes.xml / endnotes.xml into a map of note id -> content blocks.
// `kind` is 'footnote' or 'endnote'; the root/child element names follow it.
export async function parseNotesXml(
  xml: string,
  kind: 'footnote' | 'endnote',
  ctx: ParseContext,
): Promise<Map<string, Block[]>> {
  const map = new Map<string, Block[]>()
  const doc = parser.parse(xml) as Record<string, unknown>
  const root = doc?.[`${kind}s`] as Record<string, unknown> | undefined
  const notes = (root?.[kind] ?? []) as Record<string, unknown>[]
  for (const note of notes) {
    const id = String((note as Record<string, string>).id ?? '')
    // Skip the separator/continuation pseudo-notes (type set, no real content).
    const type = (note as Record<string, string>).type
    if (type === 'separator' || type === 'continuationSeparator') continue
    map.set(id, await parseBlockContainer(note, ctx))
  }
  return map
}
