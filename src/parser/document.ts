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

  // Field instructions and field characters are handled by the field state
  // machine in parseParagraph (so PAGE/NUMPAGES resolve live and their cached
  // result is suppressed). They never reach parseRun, but guard defensively.
  if ('instrText' in r) return null

  // Skip field characters and a note's own auto-number marker.
  if ('fldChar' in r) return null
  if ('footnoteRef' in r || 'endnoteRef' in r) return null

  // w:br: a soft line break becomes a <br>; an explicit page break becomes a
  // transient pageBreak marker (parseParagraph splits the paragraph there onto a
  // new page). A column break has no place in the single-column flow — dropped.
  let lineBreak = false
  if ('br' in r) {
    const br = r.br as Record<string, string> | string | undefined
    const brType = typeof br === 'object' && br !== null ? br.type : undefined
    if (brType === 'page') {
      return { type: 'run', text: '', style: Object.assign({}, paraStyle, extractRPr(rPr)), pageBreak: true }
    }
    if (brType === 'column') return null
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

  // Text run. A tracked-deletion run carries its text in <w:delText> instead of
  // <w:t>, so fall back to it.
  const tNode = r.t ?? r.delText
  let text = ''
  if (typeof tNode === 'string') {
    text = tNode
  } else if (typeof tNode === 'object' && tNode !== null) {
    const t = tNode as Record<string, unknown>
    text = String(t['#text'] ?? t._ ?? '')
  } else if (tNode != null) {
    text = String(tNode)
  }

  // Leading tab(s) in the run (w:tab) — rendered as spacers.
  const tabNode = r.tab
  const tabs = Array.isArray(tabNode) ? tabNode.length : 'tab' in r ? 1 : 0

  const textRun: TextRun = {
    type: 'run', text, style: runStyle,
    ...(href ? { href } : {}),
    ...(lineBreak ? { lineBreak: true } : {}),
    ...(tabs ? { tabs } : {}),
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
): Promise<ParagraphBlock[]> {
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
  type RunInput = { r: Record<string, unknown>; href?: string; deleted?: boolean; inserted?: boolean }
  const directRuns = (p.r ?? []) as Record<string, unknown>[]
  const insList = (p.ins ?? []) as Record<string, unknown>[]
  const delList = (p.del ?? []) as Record<string, unknown>[]
  const hyperlinks = (p.hyperlink ?? []) as Record<string, unknown>[]

  const insInputs = (ins: Record<string, unknown>): RunInput[] => {
    const inner = ins.r
    if (Array.isArray(inner)) return inner.map(r => ({ r, inserted: true }))
    return inner ? [{ r: inner as Record<string, unknown>, inserted: true }] : []
  }
  const delInputs = (del: Record<string, unknown>): RunInput[] => {
    const inner = del.r
    if (Array.isArray(inner)) return inner.map(r => ({ r, deleted: true }))
    return inner ? [{ r: inner as Record<string, unknown>, deleted: true }] : []
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
  const counts = { r: 0, hyperlink: 0, ins: 0, del: 0 }
  for (const t of order) counts[t]++
  const orderOk = order.length > 0 &&
    counts.r === directRuns.length && counts.hyperlink === hyperlinks.length &&
    counts.ins === insList.length && counts.del === delList.length

  const inputs: RunInput[] = []
  if (orderOk) {
    let ri = 0, hi = 0, ii = 0, di = 0
    for (const t of order) {
      if (t === 'r') inputs.push({ r: directRuns[ri++] })
      else if (t === 'hyperlink') inputs.push(...hlInputs(hyperlinks[hi++]))
      else if (t === 'ins') inputs.push(...insInputs(insList[ii++]))
      else inputs.push(...delInputs(delList[di++]))
    }
  } else {
    for (const r of directRuns) inputs.push({ r })
    for (const ins of insList) inputs.push(...insInputs(ins))
    for (const del of delList) inputs.push(...delInputs(del))
    for (const hl of hyperlinks) inputs.push(...hlInputs(hl))
  }

  // Field state machine. A complex field is begin → instrText(code) → separate →
  // cached-result-runs → end. PAGE / NUMPAGES are resolved live (a marker is
  // emitted at the field code), so their cached result runs are suppressed to
  // avoid rendering the live value AND a stale cached number side by side.
  const LIVE_FIELDS = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES'])
  const fieldStack: { name: string; inResult: boolean }[] = []
  const runs: Run[] = []
  for (const { r, href, deleted, inserted } of inputs) {
    if ('fldChar' in r) {
      const t = (r.fldChar as Record<string, string>)?.fldCharType
      if (t === 'begin') fieldStack.push({ name: '', inResult: false })
      else if (t === 'separate') { if (fieldStack.length) fieldStack[fieldStack.length - 1].inResult = true }
      else if (t === 'end') fieldStack.pop()
      continue
    }
    if ('instrText' in r) {
      const instr = (typeof r.instrText === 'object' && r.instrText !== null
        ? String((r.instrText as Record<string, string>)['#text'] ?? '')
        : String(r.instrText)).trim().toUpperCase()
      const field = instr.split(/\s+/)[0]
      if (fieldStack.length) fieldStack[fieldStack.length - 1].name = field
      const mStyle = Object.assign({}, paraStyle, extractRPr(r.rPr as Record<string, unknown> | undefined))
      if (field === 'PAGE') runs.push({ type: 'run', text: '', style: mStyle, pageNumber: true })
      else if (field === 'NUMPAGES' || field === 'SECTIONPAGES') runs.push({ type: 'run', text: '', style: mStyle, totalPages: true })
      continue
    }
    // Suppress the cached result of a live (PAGE/NUMPAGES) field.
    if (fieldStack.some(f => f.inResult && LIVE_FIELDS.has(f.name))) continue
    const run = await parseRun(r, paraStyle, ctx, href)
    if (run !== null) {
      if (deleted && run.type === 'run') (run as TextRun).deleted = true
      if (inserted && run.type === 'run') (run as TextRun).inserted = true
      runs.push(run)
    }
  }

  // An explicit page break (w:br w:type="page") splits the paragraph into
  // segments: the run sequence is divided at each pageBreak marker (the marker
  // itself is dropped), and every segment after the first forces a new page.
  const segments: Run[][] = [[]]
  for (const run of runs) {
    if (run.type === 'run' && (run as TextRun).pageBreak) segments.push([])
    else segments[segments.length - 1].push(run)
  }

  const makeBlock = (segRuns: Run[], forceBreak: boolean): ParagraphBlock => {
    // For an empty paragraph the paragraph-mark run properties (pPr > rPr) set the
    // line's font/metrics; for a paragraph with text they apply only to the mark
    // glyph and must NOT bold/resize the runs, so we fold them into the block
    // style only when there is no visible run.
    const hasVisible = segRuns.some(r => r.type === 'image' || (r as TextRun).text.trim().length > 0)
    const blockStyle = hasVisible ? paraStyle : Object.assign({}, paraStyle, extractMarkRPr(pPr))
    return {
      type: 'paragraph',
      style: blockStyle,
      runs: segRuns,
      ...(list ? { list } : {}),
      ...(forceBreak ? { pageBreakBefore: true } : {}),
    }
  }

  return segments.map((seg, i) => makeBlock(seg, i === 0 ? pageBreakBefore : true))
}

async function parseTable(
  tbl: Record<string, unknown>,
  ctx: ParseContext,
  rawTableXml?: string,
): Promise<TableBlock> {
  const rows = (tbl.tr ?? []) as Record<string, unknown>[]

  // Map each cell node to its raw inner XML so cell paragraphs can recover run
  // order (a mid-cell hyperlink/insertion) like body paragraphs do. Built only
  // when the row/cell counts align with the parsed grid (else fall back safely).
  const tcXmlMap = new Map<Record<string, unknown>, string>()
  if (rawTableXml) {
    const cellInners = extractRowCellInners(rawTableXml)
    for (let ri = 0; ri < rows.length; ri++) {
      const cells = (rows[ri].tc ?? []) as Record<string, unknown>[]
      const rawRowInners = cellInners[ri]
      if (rawRowInners && rawRowInners.length === cells.length) {
        for (let ci = 0; ci < cells.length; ci++) tcXmlMap.set(cells[ci], rawRowInners[ci])
      }
    }
  }

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
      const cellXml = tcXmlMap.get(tc)
      const cellOrder = cellXml ? blockOrderOf(cellXml) : undefined
      const cellParaXmls = cellXml ? paragraphChunksOf(cellXml) : undefined
      const cellTableXmls = cellXml ? tableChunksOf(cellXml) : undefined
      const cellBlocks = await parseBlockContainer(tc, ctx, cellOrder, cellParaXmls, cellTableXmls)
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

// Inner XML of <w:body> (between its tags), or '' if absent.
function bodyInner(xml: string): string {
  const bodyStart = xml.indexOf('<w:body>')
  const bodyEnd = xml.lastIndexOf('</w:body>')
  if (bodyStart === -1 || bodyEnd === -1) return ''
  return xml.slice(bodyStart + 8, bodyEnd)
}

// Raw XML of each direct-child paragraph of a container's inner XML (paragraphs
// nested inside a table are skipped), in document order. Aligns 1:1 with the
// parser's grouped container.p[] array. Used to recover intra-paragraph run
// order (see getRunOrder) for the body AND for table cells.
function paragraphChunksOf(inner: string): string[] {
  const chunks: string[] = []
  const re = /<(\/?)w:(p|tbl)[\s>\/]/g
  let tblDepth = 0
  let pStart = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(inner)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2]
    if (tag === 'tbl') {
      if (isClose) tblDepth = Math.max(0, tblDepth - 1)
      else tblDepth++
      continue
    }
    if (tblDepth !== 0) continue // paragraph inside a nested table — not direct
    if (!isClose) {
      const gt = inner.indexOf('>', m.index)
      if (gt > 0 && inner[gt - 1] === '/') chunks.push(inner.slice(m.index, gt + 1)) // <w:p/>
      else if (pStart === -1) pStart = m.index
    } else if (pStart !== -1) {
      const gt = inner.indexOf('>', m.index)
      chunks.push(inner.slice(pStart, gt + 1))
      pStart = -1
    }
  }
  return chunks
}

function extractParagraphChunks(xml: string): string[] {
  return paragraphChunksOf(bodyInner(xml))
}

// Raw XML of each direct-child table of a container's inner XML, in document
// order. Nested tables are part of their parent cell's XML, so this aligns 1:1
// with the parser's grouped container.tbl[] array (body or a cell).
function tableChunksOf(inner: string): string[] {
  const chunks: string[] = []
  const re = /<(\/?)w:tbl[\s>\/]/g
  let tblDepth = 0
  let start = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(inner)) !== null) {
    const isClose = m[1] === '/'
    if (!isClose) {
      if (tblDepth === 0) start = m.index
      tblDepth++
    } else {
      tblDepth = Math.max(0, tblDepth - 1)
      if (tblDepth === 0 && start !== -1) {
        const gt = inner.indexOf('>', m.index)
        chunks.push(inner.slice(start, gt + 1))
        start = -1
      }
    }
  }
  return chunks
}

function extractTableChunks(xml: string): string[] {
  return tableChunksOf(bodyInner(xml))
}

// Inner XML of each direct cell of a table, grouped by row. Only the outer
// table's rows/cells are captured (nested tables stay inside their cell's XML),
// so the result aligns 1:1 with the parser's tbl.tr[].tc[]. Lets cell paragraphs
// recover their run order just like body paragraphs.
function extractRowCellInners(tableXml: string): string[][] {
  const rows: string[][] = []
  let tblDepth = 0
  let curRow: string[] | null = null
  let tcStart = -1
  const re = /<(\/?)w:(tbl|tr|tc)[\s>\/]/g
  let m: RegExpExecArray | null
  while ((m = re.exec(tableXml)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2]
    if (tag === 'tbl') {
      if (isClose) tblDepth = Math.max(0, tblDepth - 1)
      else tblDepth++
      continue
    }
    if (tblDepth !== 1) continue // only the outer table delimits rows/cells
    if (tag === 'tr') {
      if (isClose) { if (curRow) { rows.push(curRow); curRow = null } }
      else curRow = []
      continue
    }
    // tag === 'tc'
    if (isClose) {
      if (curRow && tcStart !== -1) { curRow.push(tableXml.slice(tcStart, m.index)); tcStart = -1 }
    } else {
      const gt = tableXml.indexOf('>', m.index)
      if (gt > 0 && tableXml[gt - 1] === '/') { if (curRow) curRow.push('') } // <w:tc/>
      else tcStart = gt + 1
    }
  }
  return rows
}

// Document order of a paragraph's content children (runs, hyperlinks, tracked
// insertions and deletions). fast-xml-parser groups these by tag, so a
// hyperlink/insertion/deletion in the MIDDLE of a paragraph would otherwise be
// reordered to the end.
function getRunOrder(paraXml: string): Array<'r' | 'hyperlink' | 'ins' | 'del'> {
  let body = paraXml
  const pprEnd = body.indexOf('</w:pPr>')
  if (pprEnd !== -1) body = body.slice(pprEnd + 8) // skip paragraph properties

  const order: Array<'r' | 'hyperlink' | 'ins' | 'del'> = []
  const re = /<(\/?)w:(r|hyperlink|ins|del)[\s>\/]/g
  let depth = 0    // inside a hyperlink/ins/del wrapper
  let inRun = false // inside a <w:r> (skip its rPr-change ins/del markers)
  let m: RegExpExecArray | null
  while ((m = re.exec(body)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2] as 'r' | 'hyperlink' | 'ins' | 'del'
    if (tag === 'r') {
      if (isClose) inRun = false
      else {
        if (depth === 0) order.push('r')
        inRun = true
      }
      continue
    }
    if (inRun) continue // an <w:ins>/<w:del> inside a run is an rPr change, not content
    if (isClose) depth = Math.max(0, depth - 1)
    else {
      if (depth === 0) order.push(tag)
      depth++
    }
  }
  return order
}

// Document order of the direct-child p and tbl elements of a container's inner
// XML. fast-xml-parser groups elements by tag name, losing cross-type ordering.
function blockOrderOf(inner: string): Array<'p' | 'tbl'> {
  const order: Array<'p' | 'tbl'> = []
  // Match opening/closing w:p and w:tbl tags ([\s>\/] excludes w:pPr, w:pStyle etc.)
  const re = /<(\/?)w:(p|tbl)[\s>\/]/g
  let tblDepth = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(inner)) !== null) {
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

function getBodyBlockOrder(xml: string): Array<'p' | 'tbl'> {
  return blockOrderOf(bodyInner(xml))
}

// Parse all paragraph and table children of a container node (body, cell, etc.)
// paraXmls (when provided, aligned 1:1 with container.p[]) lets paragraphs
// recover their run order from raw XML.
async function parseBlockContainer(
  container: Record<string, unknown>,
  ctx: ParseContext,
  order?: Array<'p' | 'tbl'>,
  paraXmls?: string[],
  tableXmls?: string[],
): Promise<Block[]> {
  const ps = (container.p ?? []) as Record<string, unknown>[]
  const tbls = (container.tbl ?? []) as Record<string, unknown>[]
  const xmls = paraXmls && paraXmls.length === ps.length ? paraXmls : undefined
  const tXmls = tableXmls && tableXmls.length === tbls.length ? tableXmls : undefined

  if (!order || order.length === 0) {
    const blocks: Block[] = []
    for (let i = 0; i < ps.length; i++) blocks.push(...await parseParagraph(ps[i], ctx, xmls?.[i]))
    for (let i = 0; i < tbls.length; i++) blocks.push(await parseTable(tbls[i], ctx, tXmls?.[i]))
    return blocks
  }

  const blocks: Block[] = []
  let pIdx = 0, tblIdx = 0
  for (const type of order) {
    if (type === 'p' && pIdx < ps.length) {
      blocks.push(...await parseParagraph(ps[pIdx], ctx, xmls?.[pIdx]))
      pIdx++
    } else if (type === 'tbl' && tblIdx < tbls.length) {
      blocks.push(await parseTable(tbls[tblIdx], ctx, tXmls?.[tblIdx]))
      tblIdx++
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
  const tableXmls = extractTableChunks(xml)
  return parseBlockContainer(body, ctx, order, paraXmls, tableXmls)
}

// Parse a footer part (footerN.xml, root <w:ftr>) into content blocks.
export async function parseFooterXml(xml: string, ctx: ParseContext): Promise<Block[]> {
  const doc = parser.parse(xml) as Record<string, unknown>
  const ftr = doc?.ftr as Record<string, unknown> | undefined
  if (!ftr) return []
  return parseBlockContainer(ftr, ctx)
}

// Parse a header part (headerN.xml, root <w:hdr>) into content blocks.
export async function parseHeaderXml(xml: string, ctx: ParseContext): Promise<Block[]> {
  const doc = parser.parse(xml) as Record<string, unknown>
  const hdr = doc?.hdr as Record<string, unknown> | undefined
  if (!hdr) return []
  return parseBlockContainer(hdr, ctx)
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
