import { XMLParser } from 'fast-xml-parser'
import type { ComputedStyle } from '../types.js'

export type StyleMap = Record<string, ComputedStyle>

const parser = new XMLParser({
  removeNSPrefix: true,
  attributeNamePrefix: '',
  ignoreAttributes: false,
  parseAttributeValue: false,
  isArray: (name) => name === 'style',
})

// OOXML w:highlight named colors → CSS colors.
const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: 'yellow', green: 'lime', cyan: 'cyan', magenta: 'magenta', red: 'red',
  blue: 'blue', white: 'white', black: 'black', darkBlue: 'darkblue',
  darkCyan: 'darkcyan', darkGreen: 'darkgreen', darkMagenta: 'darkmagenta',
  darkRed: 'darkred', darkYellow: '#808000', darkGray: '#a9a9a9', lightGray: '#d3d3d3',
}

// Extract ComputedStyle from a raw rPr/pPr node (already namespace-stripped).
export function extractRPr(rPr: Record<string, unknown> | undefined): ComputedStyle {
  if (!rPr) return {}
  const s: ComputedStyle = {}

  // bold
  if ('b' in rPr) {
    const b = rPr.b as Record<string, string> | string | boolean
    const val = typeof b === 'object' && b !== null ? (b as Record<string, string>).val : undefined
    s.bold = val === '0' || val === 'false' || val === 'off' ? false : true
  }

  // italic
  if ('i' in rPr) {
    const i = rPr.i as Record<string, string> | string | boolean
    const val = typeof i === 'object' && i !== null ? (i as Record<string, string>).val : undefined
    s.italic = val === '0' || val === 'false' || val === 'off' ? false : true
  }

  // underline
  if ('u' in rPr) {
    const u = rPr.u as Record<string, string>
    const val = typeof u === 'object' && u !== null ? u.val : undefined
    if (val === 'none' || val === '0' || val === 'false' || val === 'off') {
      s.underline = false
    } else {
      s.underline = true
    }
  }

  // strikethrough
  if ('strike' in rPr) {
    const st = rPr.strike as Record<string, string> | undefined
    const val = typeof st === 'object' && st !== null ? st.val : undefined
    s.strike = !(val === '0' || val === 'false' || val === 'off')
  }

  // super/subscript
  if ('vertAlign' in rPr) {
    const va = (rPr.vertAlign as Record<string, string>)?.val
    if (va === 'superscript') s.vertAlign = 'super'
    else if (va === 'subscript') s.vertAlign = 'sub'
  }

  // named highlight color (distinct from w:shd fill)
  if ('highlight' in rPr) {
    const val = (rPr.highlight as Record<string, string>)?.val
    const css = HIGHLIGHT_COLORS[val ?? '']
    if (css) s.highlight = css
  }

  // font size: w:sz stores HALF-POINTS
  if ('sz' in rPr) {
    const sz = rPr.sz as Record<string, string>
    const raw = typeof sz === 'object' && sz !== null ? sz.val : String(sz)
    const half = parseInt(raw, 10)
    if (!isNaN(half)) s.fontSize = half / 2
  }

  // font family: prefer w:ascii, fallback w:hAnsi
  if ('rFonts' in rPr) {
    const fonts = rPr.rFonts as Record<string, string>
    const family = fonts.ascii ?? fonts.hAnsi
    if (family) s.fontFamily = family
  }

  // color: skip "auto"
  if ('color' in rPr) {
    const color = rPr.color as Record<string, string>
    const val = typeof color === 'object' && color !== null ? color.val : String(color)
    if (val && val !== 'auto') s.color = val
  }

  // character-level shading (background highlight)
  if ('shd' in rPr) {
    const shd = rPr.shd as Record<string, string>
    const fill = typeof shd === 'object' && shd !== null ? shd.fill : undefined
    if (fill && fill !== 'auto') s.backgroundColor = fill
  }

  return s
}

const twipsToPx = (twips: string): number => Math.round((parseFloat(twips) * 96) / 1440)

export function extractPPr(pPr: Record<string, unknown> | undefined): Partial<ComputedStyle> {
  if (!pPr) return {}
  const s: Partial<ComputedStyle> = {}

  // alignment. OOXML uses w:val="both" (and "distribute") for justified text,
  // not "justify"; "end"/"start" are the bidi-aware right/left.
  if ('jc' in pPr) {
    const jc = pPr.jc as Record<string, string>
    const val = typeof jc === 'object' && jc !== null ? jc.val : String(jc)
    if (val === 'center') s.alignment = 'center'
    else if (val === 'right' || val === 'end') s.alignment = 'right'
    else if (val === 'both' || val === 'distribute' || val === 'justify') s.alignment = 'justify'
    else s.alignment = 'left'
  }

  // paragraph borders (w:pBdr): top/bottom/left/right with sz (eighths of a pt),
  // val (line style) and color.
  if ('pBdr' in pPr) {
    const bdr = pPr.pBdr as Record<string, unknown>
    const sides = [['top', 'borderTop'], ['bottom', 'borderBottom'], ['left', 'borderLeft'], ['right', 'borderRight']] as const
    for (const [side, key] of sides) {
      const b = bdr[side] as Record<string, string> | undefined
      if (!b || typeof b !== 'object') continue
      const val = b.val
      if (!val || val === 'nil' || val === 'none') continue
      const px = Math.max(1, Math.round((b.sz ? parseFloat(b.sz) / 8 : 0.5) * 96 / 72))
      const lineStyle = val === 'double' ? 'double' : val === 'dashed' ? 'dashed' : val === 'dotted' ? 'dotted' : 'solid'
      const color = b.color && b.color !== 'auto' ? `#${b.color}` : '#000'
      s[key] = `${px}px ${lineStyle} ${color}`
    }
  }

  // indentation (w:ind): left/right and first-line/hanging, all in twips
  if ('ind' in pPr) {
    const ind = pPr.ind as Record<string, string>
    if (ind && typeof ind === 'object') {
      if (ind.left != null) s.indentLeft = twipsToPx(ind.left)
      if (ind.right != null) s.indentRight = twipsToPx(ind.right)
      if (ind.hanging != null) s.indentHanging = twipsToPx(ind.hanging)
      else if (ind.firstLine != null) s.indentFirstLine = twipsToPx(ind.firstLine)
    }
  }

  // paragraph spacing: before/after (twips) and line spacing (w:line + w:lineRule)
  if ('spacing' in pPr) {
    const sp = pPr.spacing as Record<string, string>
    if (sp && typeof sp === 'object') {
      if (sp.before != null) s.spaceBefore = twipsToPx(sp.before)
      if (sp.after != null) s.spaceAfter = twipsToPx(sp.after)
      if (sp.line != null) {
        if (sp.lineRule === 'atLeast' || sp.lineRule === 'exact') {
          s.lineHeightPx = twipsToPx(sp.line) // line is in twips for these rules
        } else {
          s.lineHeight = parseFloat(sp.line) / 240 // "auto": line is in 240ths of a line
        }
      }
    }
  }

  // NOTE: pPr can also contain an <w:rPr> — but that is the PARAGRAPH MARK's run
  // properties (the ¶ glyph), which must NOT cascade onto the paragraph's text
  // runs (e.g. a bold mark would wrongly bold the whole line). It is read
  // separately by the caller (see extractMarkRPr) and applied only to the mark
  // / empty paragraphs.

  return s
}

// The paragraph mark's run properties (pPr > rPr). Used for an empty paragraph's
// line metrics; NOT inherited by the paragraph's text runs.
export function extractMarkRPr(pPr: Record<string, unknown> | undefined): ComputedStyle {
  if (!pPr || !('rPr' in pPr)) return {}
  return extractRPr(pPr.rPr as Record<string, unknown>)
}

function resolveStyleChain(
  styleId: string,
  rawStyles: Record<string, Record<string, unknown>>,
  visited: Set<string>,
): ComputedStyle {
  if (visited.has(styleId)) {
    console.warn(`[simple-super-doc] basedOn cycle detected at styleId "${styleId}"`)
    return {}
  }
  visited.add(styleId)

  const raw = rawStyles[styleId]
  if (!raw) return {}

  let base: ComputedStyle = {}

  const basedOn = raw.basedOn as Record<string, string> | undefined
  if (basedOn?.val) {
    base = resolveStyleChain(basedOn.val, rawStyles, visited)
  }

  const pPrStyle = extractPPr(raw.pPr as Record<string, unknown> | undefined)
  const rPrStyle = extractRPr(raw.rPr as Record<string, unknown> | undefined)

  return Object.assign({}, base, pPrStyle, rPrStyle)
}

export function parseStyles(xml: string): { styleMap: StyleMap; docDefaults: ComputedStyle } {
  const doc = parser.parse(xml)
  const styles = doc?.styles ?? {}

  // Parse docDefaults
  const docDefaults: ComputedStyle = {}
  const rPrDefault = styles.docDefaults?.rPrDefault?.rPr
  if (rPrDefault) Object.assign(docDefaults, extractRPr(rPrDefault))
  const pPrDefault = styles.docDefaults?.pPrDefault?.pPr
  if (pPrDefault) Object.assign(docDefaults, extractPPr(pPrDefault))

  // Collect raw styles by styleId
  const rawStyles: Record<string, Record<string, unknown>> = {}
  const styleList: unknown[] = styles.style ?? []
  for (const s of styleList) {
    const style = s as Record<string, unknown>
    const id = style.styleId as string
    if (id) rawStyles[id] = style
  }

  // Resolve each style's full chain
  const styleMap: StyleMap = {}
  for (const id of Object.keys(rawStyles)) {
    styleMap[id] = resolveStyleChain(id, rawStyles, new Set<string>())
  }

  return { styleMap, docDefaults }
}
