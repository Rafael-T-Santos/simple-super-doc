import { XMLParser } from 'fast-xml-parser'
import type JSZip from 'jszip'
import type {
  Block, ParagraphBlock, TableBlock, TableCell, TableRow,
  TextRun, ImageRun, Run, ComputedStyle, ListRef, PageSize, CellBorders,
} from '../types.js'
import type { StyleMap, TableBorderMap, RawBorders } from './styles.js'
import { extractRPr, extractPPr, extractMarkRPr, extractBorders } from './styles.js'
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
     'hyperlink', 'bookmarkStart', 'ins', 'del', 'moveTo', 'moveFrom', 'fldSimple',
     'footnote', 'endnote'].includes(name),
})

export type ParseContext = {
  styleMap: StyleMap
  docDefaults: ComputedStyle
  tableBorderMap: TableBorderMap
  abstractNumMap: AbstractNumMap
  numMap: NumMap
  relationshipMap: RelationshipMap
  zip: JSZip
  // Footnote/endnote ids in document order; index+1 is the printed number.
  footnoteRefs: string[]
  endnoteRefs: string[]
}

// An EMPTY run (<w:r/> or <w:r></w:r>, which Word emits routinely) parses to the
// string "" rather than an object, and the `in` checks used to inspect a run
// throw a TypeError on a non-object — which failed the WHOLE document. Normalize
// to an empty object so an empty run flows through as the empty run it is.
function asRunObject(node: unknown): Record<string, unknown> {
  return (typeof node === 'object' && node !== null ? node : {}) as Record<string, unknown>
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

  return {
    numId, ilvl, ordered, start, format: levelInfo.format,
    ...(levelInfo.text !== undefined ? { bulletText: levelInfo.text } : {}),
  }
}

async function parseRun(
  r: Record<string, unknown>,
  paraStyle: ComputedStyle,
  ctx: ParseContext,
  href?: string,
  runXml?: string,
): Promise<Run | Run[] | null> {
  const rPr = r.rPr as Record<string, unknown> | undefined

  // Field instructions and field characters are handled by the field state
  // machine in parseParagraph (so PAGE/NUMPAGES resolve live and their cached
  // result is suppressed). They never reach parseRun, but guard defensively.
  if ('instrText' in r) return null

  // Skip field characters and a note's own auto-number marker.
  if ('fldChar' in r) return null
  if ('footnoteRef' in r || 'endnoteRef' in r) return null

  // The run's inline children, in document order, recovered from raw XML. Used
  // only when the recovered parts account for EVERY grouped child, so a scan
  // that disagrees with the parsed object falls back to the positional
  // heuristics below rather than dropping content.
  const collectText = (node: unknown): string[] => {
    if (node == null) return []
    if (Array.isArray(node)) return node.flatMap(collectText)
    if (typeof node === 'string') return [node]
    if (typeof node === 'object') {
      const t = node as Record<string, unknown>
      return [String(t['#text'] ?? t._ ?? '')]
    }
    return [String(node)]
  }
  const asArray = (n: unknown): unknown[] => (Array.isArray(n) ? n : n != null ? [n] : [])
  const rawSegments = collectText(r.t ?? r.delText)
  const tabCount = Array.isArray(r.tab) ? r.tab.length : 'tab' in r ? 1 : 0
  const symList = asArray(r.sym)
  const drawingList = asArray(r.drawing)
  const pictList = asArray(r.pict)
  const scanned = runXml ? getRunChildOrder(runXml) : undefined
  const countOf = (k: RunPart['kind']): number => (scanned ?? []).filter(p => p.kind === k).length
  const orderedParts =
    scanned &&
    countOf('text') === rawSegments.length &&
    countOf('tab') === tabCount &&
    countOf('sym') === symList.length &&
    countOf('noBreakHyphen') === asArray(r.noBreakHyphen).length &&
    countOf('softHyphen') === asArray(r.softHyphen).length &&
    countOf('drawing') === drawingList.length &&
    countOf('pict') === pictList.length
      ? scanned
      : undefined

  // w:br: a soft line break becomes a <br>; an explicit page break becomes a
  // transient pageBreak marker (parseParagraph splits the paragraph there onto a
  // new page). A column break has no place in the single-column flow — dropped.
  let lineBreak = false
  if ('br' in r) {
    const br = r.br as Record<string, string> | string | undefined
    // Multiple <w:br/> in one run make this an array; the fallback below can only
    // act on a single kind, so it reads the first (the ordered path handles each
    // break individually and ignores this).
    const first = Array.isArray(br) ? br[0] : br
    const brType = typeof first === 'object' && first !== null
      ? (first as Record<string, string>).type
      : undefined
    // Without recovered order, a page/column break can only be treated as the
    // whole run — which drops any text sharing it. With order, fall through and
    // let the ordered path place the break BETWEEN the texts instead.
    if (!orderedParts) {
      if (brType === 'page') {
        return { type: 'run', text: '', style: Object.assign({}, paraStyle, extractRPr(rPr)), pageBreak: true }
      }
      if (brType === 'column') return null
    }
    lineBreak = true
  }
  // w:cr — a carriage return is a soft line break, same as <w:br/> with no type.
  if ('cr' in r) lineBreak = true

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

  // Image via drawing. Built per occurrence so the ordered path below can place
  // an image BETWEEN the texts of the same run instead of replacing the run.
  const buildDrawing = async (drawing: Record<string, unknown>): Promise<ImageRun | null> => {
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

    // Anchor offset (EMU → px) so a positioned header/footer logo can be placed
    // (letterhead). Body floating layout stays out of scope; the renderer only
    // consults these inside a header/footer.
    const offsetPx = (pos: unknown): number | undefined => {
      const off = (pos as Record<string, unknown> | undefined)?.posOffset
      const n = off != null ? parseInt(String((off as Record<string, unknown>)['#text'] ?? off), 10) : NaN
      return Number.isNaN(n) ? undefined : Math.round(n / 9525)
    }
    const anchorXPx = anchor ? offsetPx(anchor.positionH) : undefined
    const anchorYPx = anchor ? offsetPx(anchor.positionV) : undefined

    const img: ImageRun = {
      type: 'image',
      src: resolved.src,
      widthPx: Math.round(cx / 9525),
      heightPx: Math.round(cy / 9525),
      ...(isPageBackground ? { isPageBackground: true } : {}),
      ...(href ? { href } : {}),
      ...(anchorXPx !== undefined ? { anchorXPx } : {}),
      ...(anchorYPx !== undefined ? { anchorYPx } : {}),
    }
    return img
  }

  // VML image: <w:pict><v:shape style="width:..pt;height:..pt"><v:imagedata
  // r:id="..."/></v:shape></w:pict>. The legacy image form (older Word, some
  // producers). A pict with no imagedata is a VML text box — its text is
  // recovered separately (collectTxbxContent), so this run renders nothing.
  const buildPict = async (pict: unknown): Promise<ImageRun | null> => {
    const vml = findVmlImage(pict)
    if (!vml) return null
    const resolved = await resolveImage(vml.id, ctx.relationshipMap, ctx.zip)
    if (!resolved) return null
    const ptToPx = (pt: string | undefined): number =>
      pt ? Math.round((parseFloat(pt) * 96) / 72) : 0
    return {
      type: 'image',
      src: resolved.src,
      widthPx: ptToPx(/width:([\d.]+)pt/.exec(vml.style ?? '')?.[1]),
      heightPx: ptToPx(/height:([\d.]+)pt/.exec(vml.style ?? '')?.[1]),
      ...(href ? { href } : {}),
    }
  }

  // Without recovered order an image has to stand for the whole run, which drops
  // any text sharing it. With order, fall through so the ordered path can place
  // the image between the texts.
  if (!orderedParts) {
    if ('drawing' in r) return await buildDrawing(asRunObject(r.drawing))
    if ('pict' in r) return await buildPict(r.pict)
  }

  // Text run(s). A tracked-deletion run carries its text in <w:delText> instead
  // of <w:t>. A single run may also hold MULTIPLE <w:t> segments interleaved
  // with <w:tab/> and breaks — Google Docs packs "text <w:tab/> text" into one
  // <w:r>, and Word puts a Shift+Enter/Ctrl+Enter break inside the run holding
  // the surrounding text.
  //
  // Preferred path: emit runs in TRUE document order. A run is shaped
  // [leading tabs][text][trailing <br>], so any part that cannot extend the
  // current run closes it and opens the next — which is what puts a break
  // between two texts instead of after both, keeps N breaks as N, and stops a
  // page break from swallowing the text that shares its run.
  if (orderedParts) {
    const makeOrdered = (text: string, tabs: number, br: boolean): TextRun => ({
      type: 'run', text, style: runStyle,
      ...(href ? { href } : {}),
      ...(br ? { lineBreak: true } : {}),
      ...(tabs ? { tabs } : {}),
    })
    const out: Run[] = []
    let text = ''
    let tabs = 0
    let started = false // a text/sym/hyphen part has opened the current run
    let ti = 0, syi = 0, di = 0, pi = 0
    let font: string | undefined
    const flush = (withBreak: boolean): void => {
      if (!started && !withBreak && tabs === 0) return
      const run = makeOrdered(text, tabs, withBreak)
      // font is set only when a symbol OPENED this run (a text part always
      // flushes first), so the run is symbol-only and keeps the symbol font.
      if (font) run.style = Object.assign({}, runStyle, { fontFamily: font })
      out.push(run)
      text = ''; tabs = 0; started = false; font = undefined
    }
    for (const p of orderedParts) {
      switch (p.kind) {
        case 'text':
          if (started) flush(false) // adjacent <w:t> stay separate runs, as before
          text = rawSegments[ti++] ?? ''
          started = true
          break
        case 'sym': {
          const sy = symList[syi++] as Record<string, string> | undefined
          const code = parseInt(String(sy?.char ?? ''), 16)
          // Guard the valid Unicode range so a malformed w:char can't throw.
          if (code >= 0 && code <= 0x10ffff) text += String.fromCodePoint(code)
          if (!started) font = sy?.font // symbol-only run keeps the symbol font
          started = true
          break
        }
        case 'noBreakHyphen': text += '‑'; started = true; break
        case 'softHyphen': text += '­'; started = true; break
        case 'tab':
          if (started) flush(false) // tabs are LEADING — they belong to the next run
          tabs++
          break
        case 'lineBreak': flush(true); break
        case 'pageBreak':
          flush(false)
          // Transient marker; parseParagraph splits the paragraph here.
          out.push({ type: 'run', text: '', style: runStyle, pageBreak: true })
          break
        case 'columnBreak': break // no place in the single-column flow — dropped
        case 'drawing': {
          flush(false)
          const img = await buildDrawing(asRunObject(drawingList[di++]))
          if (img) out.push(img)
          break
        }
        case 'pict': {
          flush(false)
          const img = await buildPict(pictList[pi++])
          if (img) out.push(img)
          break
        }
      }
    }
    flush(false)
    if (out.length === 0) {
      // Nothing survived. A column break has no place in a single-column flow,
      // and a drawing/pict that resolved to no image (an unresolvable rId, or a
      // VML text box whose text is recovered separately) renders nothing — drop
      // the run entirely rather than leaving an empty one behind.
      const dropped = orderedParts.some(
        p => p.kind === 'columnBreak' || p.kind === 'drawing' || p.kind === 'pict')
      if (dropped) return null
      return { type: 'run', text: '', style: runStyle, ...(href ? { href } : {}) }
    }
    return out.length === 1 ? out[0] : out
  }

  // Fallback: no raw XML for this run (or the scan disagreed with the parsed
  // object), so positions have to be guessed. Split into ordered runs by the two
  // dominant tab patterns: one tab between each segment, else all tabs between
  // the first two (keeps a footer's "V.4 <tabs> CONFIDENCIAL" split so the
  // trailing part right-aligns instead of getting dragged along the tabs).
  const segments = rawSegments.slice()

  // Inline character-producing elements a run may carry alongside <w:t>: symbol
  // glyphs (Insert Symbol, in a symbol font), non-breaking (U+2011) and soft
  // (U+00AD) hyphens. A symbol is usually its own run; hyphens sit between words.
  // The XML grouping loses their position, so weave hyphens between the first two
  // text segments (the common "word<nbh>word" packing) and merge a symbol into
  // the run text (carrying its symbol font when the run is symbol-only).
  let symText = ''
  let symFont: string | undefined
  for (const sy of symList) {
    const code = parseInt(String((sy as Record<string, string>)?.char ?? ''), 16)
    // Guard the valid Unicode range so a malformed w:char can't throw a RangeError.
    if (code >= 0 && code <= 0x10ffff) symText += String.fromCodePoint(code)
    symFont ??= (sy as Record<string, string>)?.font
  }
  const hyphens =
    '‑'.repeat(asArray(r.noBreakHyphen).length) + '­'.repeat(asArray(r.softHyphen).length)
  if (hyphens) {
    if (segments.length >= 2) segments[1] = hyphens + segments[1]
    else segments[0] = (segments[0] ?? '') + hyphens
  }
  if (symText) {
    if (segments.length === 0) {
      const style = symFont ? Object.assign({}, runStyle, { fontFamily: symFont }) : runStyle
      return {
        type: 'run', text: symText, style,
        ...(href ? { href } : {}),
        ...(lineBreak ? { lineBreak: true } : {}),
      }
    }
    segments[segments.length - 1] += symText
  }

  const makeRun = (text: string, tabs: number, isLast: boolean): TextRun => ({
    type: 'run', text, style: runStyle,
    ...(href ? { href } : {}),
    ...(isLast && lineBreak ? { lineBreak: true } : {}),
    ...(tabs ? { tabs } : {}),
  })

  // Simple case: 0 or 1 text segment with leading tabs.
  if (segments.length <= 1) return makeRun(segments[0] ?? '', tabCount, true)

  const interleave = tabCount === segments.length - 1
  return segments.map((seg, i) =>
    makeRun(seg, i === 0 ? 0 : interleave ? 1 : i === 1 ? tabCount : 0, i === segments.length - 1),
  )
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
  // a:blip carries r:embed (packaged image) or r:link (external/linked image).
  if ('embed' in obj && typeof obj.embed === 'string') return obj.embed
  if ('link' in obj && typeof obj.link === 'string') return obj.link
  for (const v of Object.values(obj)) {
    const found = findBlipEmbed(v)
    if (found) return found
  }
  return undefined
}

// Find a VML image: a <v:imagedata r:id="..."/> anywhere under a <w:pict>,
// returning its relationship id and the carrying <v:shape>'s style (which holds
// the width/height, e.g. "width:150pt;height:120pt"). A pict with no imagedata
// (a VML text box, handled separately) yields undefined.
function findVmlImage(node: unknown): { id: string; style?: string } | undefined {
  if (!node || typeof node !== 'object') return undefined
  if (Array.isArray(node)) {
    for (const x of node) { const f = findVmlImage(x); if (f) return f }
    return undefined
  }
  const obj = node as Record<string, unknown>
  if ('imagedata' in obj) {
    const imgs = (Array.isArray(obj.imagedata) ? obj.imagedata : [obj.imagedata]) as Record<string, string>[]
    for (const im of imgs) {
      // The parser's isArray covers 'style' (for <w:style>), so the v:shape's
      // style ATTRIBUTE arrives wrapped in an array — unwrap it.
      const styleRaw = Array.isArray(obj.style) ? obj.style[0] : obj.style
      if (im?.id) return { id: im.id, style: typeof styleRaw === 'string' ? styleRaw : undefined }
    }
  }
  for (const v of Object.values(obj)) { const f = findVmlImage(v); if (f) return f }
  return undefined
}

async function parseParagraph(
  p: Record<string, unknown>,
  ctx: ParseContext,
  rawXml?: string,
): Promise<Block[]> {
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
  // fldSimple is a self-contained field: <w:fldSimple w:instr=" PAGE "> wraps
  // its own cached-result runs. r is optional on those inputs (the field is
  // resolved from the instr attribute, not from a run).
  type RunInput = { r?: Record<string, unknown>; href?: string; deleted?: boolean; inserted?: boolean; fldSimple?: Record<string, unknown>; xml?: string }
  const directRuns = (p.r ?? []) as Record<string, unknown>[]
  const insList = (p.ins ?? []) as Record<string, unknown>[]
  const delList = (p.del ?? []) as Record<string, unknown>[]
  // Tracked MOVES: w:moveTo is the text at its NEW location (kept in the final
  // view, like an insertion); w:moveFrom is the same text at its OLD location
  // (removed in the final view, like a deletion). Treating them as ins/del keeps
  // the moved text exactly once and supports showRevisions, instead of dropping
  // both copies (the runs live inside moveTo/moveFrom, not directly in the para).
  const moveToList = (p.moveTo ?? []) as Record<string, unknown>[]
  const moveFromList = (p.moveFrom ?? []) as Record<string, unknown>[]
  const hyperlinks = (p.hyperlink ?? []) as Record<string, unknown>[]
  const fldSimpleList = (p.fldSimple ?? []) as Record<string, unknown>[]

  // xmls (when present) are the raw <w:r> slices of the group being expanded,
  // aligned 1:1 with its parsed runs, so each run can recover its own child order.
  const runChildInputs = (node: Record<string, unknown>, flag: 'inserted' | 'deleted', xmls?: string[]): RunInput[] => {
    const inner = node.r
    const arr = Array.isArray(inner) ? inner : inner ? [inner] : []
    const ok = xmls && xmls.length === arr.length
    return (arr as Record<string, unknown>[]).map((r, i) => ({ r, [flag]: true, ...(ok ? { xml: xmls[i] } : {}) }))
  }
  const insInputs = (ins: Record<string, unknown>, xmls?: string[]): RunInput[] => runChildInputs(ins, 'inserted', xmls)
  const delInputs = (del: Record<string, unknown>, xmls?: string[]): RunInput[] => runChildInputs(del, 'deleted', xmls)
  const hlInputs = (hl: Record<string, unknown>, xmls?: string[]): RunInput[] => {
    const href = resolveHyperlinkHref(hl, ctx)
    const inner = hl.r
    const arr = (Array.isArray(inner) ? inner : inner ? [inner] : []) as Record<string, unknown>[]
    const ok = xmls && xmls.length === arr.length
    return arr.map((r, i) => ({ r, href, ...(ok ? { xml: xmls[i] } : {}) }))
  }

  const order = rawXml ? getRunOrder(rawXml) : []
  // Use the recovered order only if it accounts for every grouped child; else
  // fall back to the safe (append-grouped) order.
  const counts = { r: 0, hyperlink: 0, ins: 0, del: 0, moveTo: 0, moveFrom: 0, fldSimple: 0 }
  for (const e of order) counts[e.tag]++
  const orderOk = order.length > 0 &&
    counts.r === directRuns.length && counts.hyperlink === hyperlinks.length &&
    counts.ins === insList.length && counts.del === delList.length &&
    counts.moveTo === moveToList.length && counts.moveFrom === moveFromList.length &&
    counts.fldSimple === fldSimpleList.length

  const inputs: RunInput[] = []
  if (orderOk) {
    let ri = 0, hi = 0, ii = 0, di = 0, mti = 0, mfi = 0, fi = 0
    for (const e of order) {
      const t = e.tag
      if (t === 'r') inputs.push({ r: directRuns[ri++], xml: e.runXmls[0] })
      else if (t === 'hyperlink') inputs.push(...hlInputs(hyperlinks[hi++], e.runXmls))
      else if (t === 'ins') inputs.push(...insInputs(insList[ii++], e.runXmls))
      else if (t === 'del') inputs.push(...delInputs(delList[di++], e.runXmls))
      else if (t === 'moveTo') inputs.push(...runChildInputs(moveToList[mti++], 'inserted', e.runXmls))
      else if (t === 'moveFrom') inputs.push(...runChildInputs(moveFromList[mfi++], 'deleted', e.runXmls))
      else inputs.push({ fldSimple: fldSimpleList[fi++] })
    }
  } else {
    for (const r of directRuns) inputs.push({ r })
    for (const ins of insList) inputs.push(...insInputs(ins))
    for (const del of delList) inputs.push(...delInputs(del))
    for (const mt of moveToList) inputs.push(...runChildInputs(mt, 'inserted'))
    for (const mf of moveFromList) inputs.push(...runChildInputs(mf, 'deleted'))
    for (const hl of hyperlinks) inputs.push(...hlInputs(hl))
    for (const fs of fldSimpleList) inputs.push({ fldSimple: fs })
  }

  // Field state machine. A complex field is begin → instrText(code) → separate →
  // cached-result-runs → end. PAGE / NUMPAGES are resolved live (a marker is
  // emitted at the field code), so their cached result runs are suppressed to
  // avoid rendering the live value AND a stale cached number side by side.
  const LIVE_FIELDS = new Set(['PAGE', 'NUMPAGES', 'SECTIONPAGES'])
  const fieldStack: { name: string; inResult: boolean }[] = []
  const runs: Run[] = []
  for (const inp of inputs) {
    // w:fldSimple — the compact field form. The field code lives in the instr
    // attribute and the cached result is the element's child runs. Resolve
    // PAGE/NUMPAGES live (suppressing the cached runs); for any other field
    // render the cached result runs as-is.
    if (inp.fldSimple) {
      const fs = inp.fldSimple
      const field = String((fs as Record<string, string>).instr ?? '').trim().toUpperCase().split(/\s+/)[0]
      const innerRuns = (Array.isArray(fs.r) ? fs.r : fs.r ? [fs.r] : []) as Record<string, unknown>[]
      const mStyle = Object.assign({}, paraStyle, extractRPr(innerRuns[0]?.rPr as Record<string, unknown> | undefined))
      if (field === 'PAGE') runs.push({ type: 'run', text: '', style: mStyle, pageNumber: true })
      else if (field === 'NUMPAGES' || field === 'SECTIONPAGES') runs.push({ type: 'run', text: '', style: mStyle, totalPages: true })
      else {
        for (const ir of innerRuns) {
          const run = await parseRun(asRunObject(ir), paraStyle, ctx)
          if (run !== null) runs.push(...(Array.isArray(run) ? run : [run]))
        }
      }
      continue
    }
    const r = asRunObject(inp.r)
    const { href, deleted, inserted, xml } = inp

    // Field components may be split one-per-run (begin | code | separate | result
    // | end) OR packed into a SINGLE run — Google Docs exports the whole field in
    // one <w:r> (begin + instrText + separate + end together), which makes
    // r.fldChar an array. Process a run's field parts in canonical order
    // (begin → code → separate → end) so both layouts resolve the same.
    const fldChars = ('fldChar' in r)
      ? (Array.isArray(r.fldChar) ? r.fldChar : [r.fldChar]) as Record<string, string>[]
      : []
    const hasInstr = 'instrText' in r
    if (fldChars.length || hasInstr) {
      for (const fc of fldChars) if (fc?.fldCharType === 'begin') fieldStack.push({ name: '', inResult: false })
      if (hasInstr) {
        const instr = (typeof r.instrText === 'object' && r.instrText !== null
          ? String((r.instrText as Record<string, string>)['#text'] ?? '')
          : String(r.instrText)).trim().toUpperCase()
        const field = instr.split(/\s+/)[0]
        if (fieldStack.length) fieldStack[fieldStack.length - 1].name = field
        const mStyle = Object.assign({}, paraStyle, extractRPr(r.rPr as Record<string, unknown> | undefined))
        if (field === 'PAGE') runs.push({ type: 'run', text: '', style: mStyle, pageNumber: true })
        else if (field === 'NUMPAGES' || field === 'SECTIONPAGES') runs.push({ type: 'run', text: '', style: mStyle, totalPages: true })
      }
      for (const fc of fldChars) if (fc?.fldCharType === 'separate') { if (fieldStack.length) fieldStack[fieldStack.length - 1].inResult = true }
      for (const fc of fldChars) if (fc?.fldCharType === 'end') fieldStack.pop()
      continue
    }
    // Suppress the cached result of a live (PAGE/NUMPAGES) field.
    if (fieldStack.some(f => f.inResult && LIVE_FIELDS.has(f.name))) continue
    const run = await parseRun(r, paraStyle, ctx, href, xml)
    if (run !== null) {
      for (const rn of Array.isArray(run) ? run : [run]) {
        if (deleted && rn.type === 'run') (rn as TextRun).deleted = true
        if (inserted && rn.type === 'run') (rn as TextRun).inserted = true
        runs.push(rn)
      }
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

  const blocks = segments.map((seg, i) => makeBlock(seg, i === 0 ? pageBreakBefore : true))

  // A section break: this paragraph's pPr carried a w:sectPr, so it is the LAST
  // paragraph of a section. Tag the last block with that section's page size so
  // the document can be split into sections (see buildSections).
  const sectPr = pPr?.sectPr as Record<string, unknown> | undefined
  const sectionPageSize = pageSizeFromSectPr(sectPr)
  if (sectionPageSize) {
    const last = blocks[blocks.length - 1]
    last.sectionPageSize = sectionPageSize
    const refs = refsFromSectPr(sectPr)
    if (refs) last.sectionRefs = refs
  }

  // Text boxes (w:txbxContent inside a w:drawing / w:pict shape) carry block
  // content the run path can't represent. Recover it into the flow right after
  // this paragraph so the text is never silently dropped. (Floating placement
  // is out of scope — see README; this is minimum-viable text recovery.)
  const textBoxes = await extractTextBoxes(p, ctx)
  return textBoxes.length ? [...blocks, ...textBoxes] : blocks
}

// Parse every text box reachable from a paragraph into flow blocks. Word emits
// a text box twice inside mc:AlternateContent (DrawingML Choice + VML Fallback),
// so only one branch is followed to avoid duplicating the content.
async function extractTextBoxes(p: Record<string, unknown>, ctx: ParseContext): Promise<Block[]> {
  const contents: Record<string, unknown>[] = []
  collectTxbxContent(p, contents)
  const blocks: Block[] = []
  for (const tc of contents) blocks.push(...await parseBlockContainer(tc, ctx))
  return blocks
}

// Collect all w:txbxContent nodes in a parsed subtree. On AlternateContent,
// descend only into the Choice (DrawingML) and ignore the Fallback (VML) so the
// same text box isn't collected twice. A found txbxContent is not recursed into
// here — any text box nested inside it is recovered when its own paragraphs are
// parsed (parseParagraph -> extractTextBoxes again).
function collectTxbxContent(node: unknown, acc: Record<string, unknown>[]): void {
  if (!node || typeof node !== 'object') return
  if (Array.isArray(node)) {
    for (const item of node) collectTxbxContent(item, acc)
    return
  }
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key === 'txbxContent') {
      for (const tc of (Array.isArray(value) ? value : [value])) {
        if (tc && typeof tc === 'object') acc.push(tc as Record<string, unknown>)
      }
    } else if (key === 'AlternateContent') {
      for (const ac of (Array.isArray(value) ? value : [value])) {
        const a = ac as Record<string, unknown>
        collectTxbxContent(a.Choice ?? a.Fallback, acc)
      }
    } else {
      collectTxbxContent(value, acc)
    }
  }
}

// Extract the default header/footer relationship ids from a parsed sectPr.
// w:headerReference / w:footerReference carry w:type (default/even/first) and
// r:id; only the default type is used (matching the doc-level resolver). The
// ids are resolved to blocks when building DocxDocument.sections.
function refsFromSectPr(
  sectPr: Record<string, unknown> | undefined,
): { headerRId?: string; footerRId?: string } | undefined {
  if (!sectPr) return undefined
  const defaultRId = (node: unknown): string | undefined => {
    const refs = (Array.isArray(node) ? node : node ? [node] : []) as Record<string, string>[]
    const def = refs.find(r => r.type === 'default') ?? refs[0]
    return def?.id
  }
  const headerRId = defaultRId(sectPr.headerReference)
  const footerRId = defaultRId(sectPr.footerReference)
  if (!headerRId && !footerRId) return undefined
  return { ...(headerRId ? { headerRId } : {}), ...(footerRId ? { footerRId } : {}) }
}

async function parseTable(
  tbl: Record<string, unknown>,
  ctx: ParseContext,
  rawTableXml?: string,
): Promise<TableBlock> {
  const rows = (tbl.tr ?? []) as Record<string, unknown>[]

  // Effective table borders: the table style's borders (resolved through its
  // basedOn chain, e.g. TableGrid) overlaid with any direct w:tblBorders. The
  // grid look comes from a uniform per-cell border (the inside borders, falling
  // back to an outer side); 'none' means explicitly off. Per-cell w:tcBorders
  // override this below. This recovers borders that come only from a table
  // style (no explicit tblBorders/tcBorders on the table).
  const tblPrForBorders = tbl.tblPr as Record<string, unknown> | undefined
  const tblStyleId = getVal(tblPrForBorders?.tblStyle as unknown)
  const styleBorders: RawBorders = (tblStyleId && ctx.tableBorderMap[tblStyleId]) || {}
  const tableBorders: RawBorders = {
    ...styleBorders,
    ...extractBorders(tblPrForBorders?.tblBorders as Record<string, unknown> | undefined),
  }
  const definedSide = (...vals: (string | undefined)[]) => vals.find(v => v && v !== 'none')
  const uniformBorder = definedSide(
    tableBorders.insideH, tableBorders.insideV,
    tableBorders.top, tableBorders.bottom, tableBorders.left, tableBorders.right,
  )

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
  for (let ri = 0; ri < resolved.length; ri++) {
    const resolvedRow = resolved[ri]
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
      const border = resolveCellBorder((tc.tcPr as Record<string, unknown> | undefined), uniformBorder)
      if (border) irCell.border = border
      irCells.push(irCell)
    }
    const irRow: TableRow = { cells: irCells }
    // Row height from w:trPr > w:trHeight (twips). hRule "exact" fixes the height;
    // the default "atLeast" is a minimum the content can grow past.
    const trH = (rows[ri]?.trPr as Record<string, unknown> | undefined)?.trHeight as Record<string, string> | undefined
    if (trH?.val != null) {
      const px = Math.round((parseFloat(trH.val) * 96) / 1440)
      if (px > 0) {
        irRow.heightPx = px
        if (trH.hRule === 'exact') irRow.heightExact = true
      }
    }
    irRows.push(irRow)
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

  // Table alignment (w:jc): a table narrower than the page is placed left/center/
  // right instead of being stretched to full width.
  const jc = getVal(tblPr?.jc as unknown)
  const align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : undefined

  return {
    type: 'table',
    rows: irRows,
    ...(columnWidths.some(w => w > 0) ? { columnWidths } : {}),
    ...(cellPadding ? { cellPadding } : {}),
    ...(align ? { align } : {}),
  }
}

// Per-cell border: each side uses the cell's own w:tcBorders when present
// ('none' = explicitly off, suppressing the uniform border), otherwise the
// table's uniform border. Returns undefined when no side has a border.
function resolveCellBorder(
  tcPr: Record<string, unknown> | undefined,
  uniform: string | undefined,
): CellBorders | undefined {
  const tcb = extractBorders(tcPr?.tcBorders as Record<string, unknown> | undefined)
  const side = (s?: string): string | undefined => (s === 'none' ? undefined : (s ?? uniform))
  const b: CellBorders = {}
  const top = side(tcb.top), right = side(tcb.right), bottom = side(tcb.bottom), left = side(tcb.left)
  if (top) b.top = top
  if (right) b.right = right
  if (bottom) b.bottom = bottom
  if (left) b.left = left
  return (b.top || b.right || b.bottom || b.left) ? b : undefined
}

// Page size/margins from a parsed <w:sectPr> object (pgSz + pgMar), in px.
// Mirrors index.ts parsePageSize but works on the already-parsed object so a
// per-section sectPr in a paragraph's pPr can be read. Returns undefined when
// there is no pgSz (e.g. a section that only changes columns).
function pageSizeFromSectPr(sectPr: Record<string, unknown> | undefined): PageSize | undefined {
  if (!sectPr) return undefined
  const twips = (v: unknown): number => Math.round((parseFloat(String(v)) * 96) / 1440)
  const pgSz = sectPr.pgSz as Record<string, string> | undefined
  if (!pgSz || pgSz.w == null || pgSz.h == null) return undefined
  const pgMar = (sectPr.pgMar ?? {}) as Record<string, string>
  const m = (v: string | undefined, d: number) => (v != null ? twips(v) : d)
  const size: PageSize = {
    widthPx: twips(pgSz.w),
    heightPx: twips(pgSz.h),
    marginPx: {
      top: m(pgMar.top, 96), right: m(pgMar.right, 96),
      bottom: m(pgMar.bottom, 96), left: m(pgMar.left, 96),
    },
  }
  if (pgMar.footer != null) size.footerPx = twips(pgMar.footer)
  if (pgMar.header != null) size.headerPx = twips(pgMar.header)
  return size
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
  return rootInner(xml, 'w:body')
}

// Inner XML of a part's root element. The body is <w:body>, but a header is
// <w:hdr>, a footer <w:ftr> and a note <w:footnote>/<w:endnote> — all of which
// need the same raw XML to recover document order (run order within a paragraph,
// and a run's own child order).
function rootInner(xml: string, tag: string): string {
  const open = xml.indexOf(`<${tag}`)
  if (open === -1) return ''
  const gt = xml.indexOf('>', open)
  const close = xml.lastIndexOf(`</${tag}>`)
  if (gt === -1 || close === -1 || close < gt) return ''
  return xml.slice(gt + 1, close)
}

// Raw XML of each direct-child <tag> element of a container, in document order.
// Used to align notes (w:footnote/w:endnote) with their parsed counterparts.
function elementChunksOf(inner: string, tag: string): string[] {
  const chunks: string[] = []
  const re = new RegExp(`<(/?)${tag}[\\s>/]`, 'g')
  let depth = 0
  let start = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(inner)) !== null) {
    if (m[1] === '/') {
      depth = Math.max(0, depth - 1)
      if (depth === 0 && start >= 0) {
        const gt = inner.indexOf('>', m.index)
        chunks.push(inner.slice(start, gt === -1 ? m.index : gt + 1))
        start = -1
      }
      continue
    }
    const gt = inner.indexOf('>', m.index)
    const selfClose = gt > 0 && inner[gt - 1] === '/'
    if (selfClose) {
      if (depth === 0) chunks.push(inner.slice(m.index, gt + 1))
      continue
    }
    if (depth === 0) start = m.index
    depth++
  }
  return chunks
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
type RunOrderTag = 'r' | 'hyperlink' | 'ins' | 'del' | 'moveTo' | 'moveFrom' | 'fldSimple'
// Each entry also carries the raw XML of the <w:r> elements it contributes (the
// run itself for 'r', the wrapped runs for a hyperlink/ins/del/...), so parseRun
// can recover the order of a run's OWN children (text vs. br vs. tab) the same
// way — see getRunChildOrder.
type RunOrderEntry = { tag: RunOrderTag; runXmls: string[] }
function getRunOrder(paraXml: string): RunOrderEntry[] {
  let body = paraXml
  const pprEnd = body.indexOf('</w:pPr>')
  if (pprEnd !== -1) body = body.slice(pprEnd + 8) // skip paragraph properties

  const order: RunOrderEntry[] = []
  // moveTo/moveFrom are tracked-move wrappers (like ins/del). The trailing
  // [\s>\/] excludes their range markers (moveToRangeStart/End etc.), which carry
  // no runs.
  const re = /<(\/?)w:(r|hyperlink|ins|del|moveTo|moveFrom|fldSimple)[\s>\/]/g
  let depth = 0    // inside a hyperlink/ins/del/moveTo/moveFrom/fldSimple wrapper
  let runDepth = 0 // inside a <w:r> (skip its rPr-change ins/del markers)
  let runStart = -1 // offset of the outermost open <w:r>, for its raw slice
  let m: RegExpExecArray | null
  while ((m = re.exec(body)) !== null) {
    const isClose = m[1] === '/'
    const tag = m[2] as RunOrderTag
    if (tag === 'r') {
      if (isClose) {
        runDepth = Math.max(0, runDepth - 1)
        // Closing the OUTERMOST run: capture its full XML on the entry that owns
        // it (the 'r' entry just pushed, or the enclosing wrapper's entry).
        if (runDepth === 0 && runStart >= 0) {
          const gt = body.indexOf('>', m.index)
          order[order.length - 1]?.runXmls.push(body.slice(runStart, gt === -1 ? m.index : gt + 1))
          runStart = -1
        }
        continue
      }
      // A run nested in a drawing/pict text box is not a paragraph-level child.
      if (runDepth === 0) {
        if (depth === 0) order.push({ tag: 'r', runXmls: [] })
        const gt = body.indexOf('>', m.index)
        const selfClose = gt > 0 && body[gt - 1] === '/'
        if (selfClose) continue // <w:r/> — no children, no slice to capture
        runStart = m.index
      }
      runDepth++
      continue
    }
    if (runDepth > 0) continue // an <w:ins>/<w:del> inside a run is an rPr change, not content
    if (isClose) { depth = Math.max(0, depth - 1); continue }
    // Opening wrapper. A self-closed element (e.g. <w:fldSimple .../>) has no
    // matching close tag, so it must not push the depth.
    const gt = body.indexOf('>', m.index)
    const selfClose = gt > 0 && body[gt - 1] === '/'
    if (depth === 0) order.push({ tag, runXmls: [] })
    if (!selfClose) depth++
  }
  return order
}

// Document order of a run's own inline children. fast-xml-parser groups these by
// tag too, so "<w:t>A</w:t><w:br/><w:t>B</w:t>" arrives as { t: [A,B], br: '' } —
// the COUNT survives but the POSITION does not. Recovering it from raw XML is
// what lets a break land between the two texts instead of after both, and what
// keeps a page break from swallowing the text sharing its run.
type RunPart =
  | { kind: 'text' } | { kind: 'tab' } | { kind: 'sym' }
  | { kind: 'noBreakHyphen' } | { kind: 'softHyphen' }
  | { kind: 'lineBreak' } | { kind: 'pageBreak' } | { kind: 'columnBreak' }
  | { kind: 'drawing' } | { kind: 'pict' }

// End offset of the element opened at `start`, so its whole subtree can be
// skipped. A drawing/pict can nest entire paragraphs (a text box), and their
// <w:t> must not be counted as the run's own text. Returns -1 if unterminated.
function skipElement(xml: string, start: number, tag: string): number {
  const gt = xml.indexOf('>', start)
  if (gt === -1) return -1
  if (xml[gt - 1] === '/') return gt + 1 // self-closed, no subtree
  const open = new RegExp(`<${tag}[\\s>/]`, 'g')
  const close = new RegExp(`</${tag}\\s*>`, 'g')
  let depth = 1
  let i = gt + 1
  while (depth > 0) {
    open.lastIndex = i
    close.lastIndex = i
    const o = open.exec(xml)
    const c = close.exec(xml)
    if (!c) return -1 // unterminated — caller declines the whole scan
    if (o && o.index < c.index) { depth++; i = o.index + 1 }
    else { depth--; i = c.index + c[0].length }
  }
  return i
}

function getRunChildOrder(runXml: string): RunPart[] | undefined {
  let body = runXml
  const rprEnd = body.indexOf('</w:rPr>')
  if (rprEnd !== -1) body = body.slice(rprEnd + 8) // skip run properties

  const parts: RunPart[] = []
  const re = /<w:(t|delText|br|cr|tab|sym|noBreakHyphen|softHyphen|drawing|pict)([\s\/>])([^>]*)?/g
  let m: RegExpExecArray | null
  while ((m = re.exec(body)) !== null) {
    const tag = m[1]
    if (tag === 'drawing' || tag === 'pict') {
      const end = skipElement(body, m.index, `w:${tag}`)
      if (end < 0) return undefined // malformed subtree — fall back
      parts.push({ kind: tag })
      re.lastIndex = end
      continue
    }
    if (tag === 't' || tag === 'delText') parts.push({ kind: 'text' })
    else if (tag === 'cr') parts.push({ kind: 'lineBreak' })
    else if (tag === 'br') {
      const type = /w:type="([^"]*)"/.exec(m[3] ?? '')?.[1]
      parts.push({ kind: type === 'page' ? 'pageBreak' : type === 'column' ? 'columnBreak' : 'lineBreak' })
    } else if (tag === 'tab') parts.push({ kind: 'tab' })
    else if (tag === 'sym') parts.push({ kind: 'sym' })
    else if (tag === 'noBreakHyphen') parts.push({ kind: 'noBreakHyphen' })
    else parts.push({ kind: 'softHyphen' })
  }
  return parts
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

// Several OOXML elements are TRANSPARENT WRAPPERS: their content should render as
// if the wrapper were not there. fast-xml-parser groups them apart from the
// surrounding <w:p>/<w:r>, so without unwrapping their content (block or inline,
// possibly nested) is dropped:
//   - w:sdt        content controls / structured document tags (forms, templates)
//   - w:smartTag   auto-recognized entities (dates, names) — legacy Word
//   - w:customXml  custom-XML data-binding wrappers
// Strip the wrapper tags from the raw XML before parsing so the inner content
// becomes normal in-place content — for both the parsed tree AND the raw-order
// scanners (getRunOrder etc.). The *Pr property elements (sdtPr/sdtEndPr/
// smartTagPr/customXmlPr) hold control metadata, not document text, so remove
// them entirely; the wrappers themselves are unwrapped (content kept, tags dropped).
// OMML math (m:oMath) stores its text in <m:t> elements, not <w:t>, so the run
// parser drops equations entirely. There is no 2D math layout here (a non-goal),
// but the content must not be lost: replace each <m:oMath> with a normal run
// carrying its concatenated math text, in place and in document order, so an
// equation like the area formula renders as its linear text ("A=πr2"). Display
// equations sit inside <m:oMathPara>, which is unwrapped so the run lands in the
// paragraph. The extracted text keeps its source XML-entity encoding.
function inlineOmmlText(xml: string): string {
  xml = xml.replace(/<m:oMath\b[\s\S]*?<\/m:oMath>/g, (frag) => {
    const text = (frag.match(/<m:t\b[^>]*>([\s\S]*?)<\/m:t>/g) || [])
      .map(t => t.replace(/<[^>]+>/g, '')).join('')
    return text ? `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>` : ''
  })
  return xml
    .replace(/<m:oMathParaPr\b[^>]*\/>/g, '')
    .replace(/<m:oMathParaPr\b[\s\S]*?<\/m:oMathParaPr>/g, '')
    .replace(/<\/?m:oMathPara\b[^>]*>/g, '')
}

function stripTransparentWrappers(xml: string): string {
  return xml
    .replace(/<w:sdtPr\b[^>]*\/>/g, '')
    .replace(/<w:sdtPr\b[\s\S]*?<\/w:sdtPr>/g, '')
    .replace(/<w:sdtEndPr\b[^>]*\/>/g, '')
    .replace(/<w:sdtEndPr\b[\s\S]*?<\/w:sdtEndPr>/g, '')
    .replace(/<w:smartTagPr\b[^>]*\/>/g, '')
    .replace(/<w:smartTagPr\b[\s\S]*?<\/w:smartTagPr>/g, '')
    .replace(/<w:customXmlPr\b[^>]*\/>/g, '')
    .replace(/<w:customXmlPr\b[\s\S]*?<\/w:customXmlPr>/g, '')
    .replace(/<\/?w:sdtContent\b[^>]*>/g, '')
    .replace(/<\/?w:sdt\b[^>]*>/g, '')
    .replace(/<\/?w:smartTag\b[^>]*>/g, '')
    .replace(/<\/?w:customXml\b[^>]*>/g, '')
}

export async function parseDocument(xml: string, ctx: ParseContext): Promise<Block[]> {
  xml = stripTransparentWrappers(inlineOmmlText(xml))
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
  xml = stripTransparentWrappers(inlineOmmlText(xml))
  const doc = parser.parse(xml) as Record<string, unknown>
  const ftr = doc?.ftr as Record<string, unknown> | undefined
  if (!ftr) return []
  const inner = rootInner(xml, 'w:ftr')
  return parseBlockContainer(ftr, ctx, blockOrderOf(inner), paragraphChunksOf(inner), tableChunksOf(inner))
}

// Parse a header part (headerN.xml, root <w:hdr>) into content blocks.
export async function parseHeaderXml(xml: string, ctx: ParseContext): Promise<Block[]> {
  xml = stripTransparentWrappers(inlineOmmlText(xml))
  const doc = parser.parse(xml) as Record<string, unknown>
  const hdr = doc?.hdr as Record<string, unknown> | undefined
  if (!hdr) return []
  const inner = rootInner(xml, 'w:hdr')
  return parseBlockContainer(hdr, ctx, blockOrderOf(inner), paragraphChunksOf(inner), tableChunksOf(inner))
}

// Parse footnotes.xml / endnotes.xml into a map of note id -> content blocks.
// `kind` is 'footnote' or 'endnote'; the root/child element names follow it.
export async function parseNotesXml(
  xml: string,
  kind: 'footnote' | 'endnote',
  ctx: ParseContext,
): Promise<Map<string, Block[]>> {
  const map = new Map<string, Block[]>()
  xml = stripTransparentWrappers(inlineOmmlText(xml))
  const doc = parser.parse(xml) as Record<string, unknown>
  const root = doc?.[`${kind}s`] as Record<string, unknown> | undefined
  const notes = (root?.[kind] ?? []) as Record<string, unknown>[]
  // Raw XML per note, aligned 1:1 with the parsed notes, so a note's paragraphs
  // recover document order the same way the body's do.
  const rawNotes = elementChunksOf(rootInner(xml, `w:${kind}s`), `w:${kind}`)
  const aligned = rawNotes.length === notes.length ? rawNotes : undefined
  for (let i = 0; i < notes.length; i++) {
    const note = notes[i]
    const id = String((note as Record<string, string>).id ?? '')
    // Skip the separator/continuation pseudo-notes (type set, no real content).
    const type = (note as Record<string, string>).type
    if (type === 'separator' || type === 'continuationSeparator') continue
    const inner = aligned ? rootInner(aligned[i], `w:${kind}`) : ''
    map.set(id, inner
      ? await parseBlockContainer(note, ctx, blockOrderOf(inner), paragraphChunksOf(inner), tableChunksOf(inner))
      : await parseBlockContainer(note, ctx))
  }
  return map
}
