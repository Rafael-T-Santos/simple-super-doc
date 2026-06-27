import { XMLParser } from 'fast-xml-parser'
import type JSZip from 'jszip'
import type {
  Block, ParagraphBlock, TableBlock, TableCell, TableRow,
  TextRun, ImageRun, Run, ComputedStyle, ListRef,
} from '../types.js'
import type { StyleMap } from './styles.js'
import { extractRPr, extractPPr } from './styles.js'
import type { AbstractNumMap, NumMap } from './numbering.js'
import type { RelationshipMap } from './relationships.js'
import { resolveVMerge, type RawCell } from './table.js'
import { resolveImage } from './images.js'

const parser = new XMLParser({
  removeNSPrefix: true,
  attributeNamePrefix: '',
  ignoreAttributes: false,
  parseAttributeValue: false,
  isArray: (name) =>
    ['p', 'r', 'tr', 'tc', 'style', 'abstractNum', 'num', 'lvl', 'lvlOverride',
     'hyperlink', 'bookmarkStart', 'ins', 'del'].includes(name),
})

export type ParseContext = {
  styleMap: StyleMap
  docDefaults: ComputedStyle
  abstractNumMap: AbstractNumMap
  numMap: NumMap
  relationshipMap: RelationshipMap
  zip: JSZip
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

  return { numId, ilvl, ordered, start }
}

async function parseRun(
  r: Record<string, unknown>,
  paraStyle: ComputedStyle,
  ctx: ParseContext,
): Promise<Run | null> {
  const rPr = r.rPr as Record<string, unknown> | undefined

  // Skip field characters and line breaks silently
  if ('fldChar' in r || 'instrText' in r || 'br' in r) return null

  // Character style at cascade level 2 for this specific run
  const rStyleId = getVal(rPr?.rStyle as unknown)
  const charStyle = rStyleId ? (ctx.styleMap[rStyleId] ?? {}) : {}

  const runStyle = Object.assign({}, paraStyle, charStyle, extractRPr(rPr))

  // Image via drawing
  if ('drawing' in r) {
    const drawing = r.drawing as Record<string, unknown>
    const drawingEl =
      (drawing.inline as Record<string, unknown> | undefined) ??
      (drawing.anchor as Record<string, unknown> | undefined)
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

  const textRun: TextRun = { type: 'run', text, style: runStyle }
  return textRun
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
): Promise<ParagraphBlock> {
  const pPr = p.pPr as Record<string, unknown> | undefined
  const pStyleId = getVal((pPr?.pStyle) as unknown)

  const namedStyle = resolveNamedStyle(pStyleId, ctx.styleMap, ctx.docDefaults)
  const pPrStyle = extractPPr(pPr)
  const paraStyle = Object.assign({}, namedStyle, pPrStyle)

  const list = resolveListRef(pPr, ctx)

  // Collect runs from: direct r[], w:ins child r[], w:hyperlink child r[]
  const directRuns = (p.r ?? []) as Record<string, unknown>[]
  const insRuns: Record<string, unknown>[] = []
  for (const ins of ((p.ins ?? []) as Record<string, unknown>[])) {
    const inner = ins.r
    if (Array.isArray(inner)) insRuns.push(...inner)
    else if (inner) insRuns.push(inner as Record<string, unknown>)
  }
  const hyperlinkRuns: Record<string, unknown>[] = []
  for (const hl of ((p.hyperlink ?? []) as Record<string, unknown>[])) {
    const inner = hl.r
    if (Array.isArray(inner)) hyperlinkRuns.push(...inner)
    else if (inner) hyperlinkRuns.push(inner as Record<string, unknown>)
  }

  const allRuns = [...directRuns, ...insRuns, ...hyperlinkRuns]
  const runs: Run[] = []
  for (const r of allRuns) {
    const run = await parseRun(r, paraStyle, ctx)
    if (run !== null) runs.push(run)
  }

  return { type: 'paragraph', style: paraStyle, runs, ...(list ? { list } : {}) }
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

      rawRow.push({ colSpan, vMerge, rawData: tc })
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
      irCells.push({ rowSpan: cell.rowSpan, colSpan: cell.colSpan, blocks: cellBlocks })
    }
    irRows.push({ cells: irCells })
  }

  return { type: 'table', rows: irRows }
}

// Parse all paragraph and table children of a container node (body, cell, etc.)
async function parseBlockContainer(
  container: Record<string, unknown>,
  ctx: ParseContext,
): Promise<Block[]> {
  const blocks: Block[] = []
  for (const p of ((container.p ?? []) as Record<string, unknown>[])) {
    blocks.push(await parseParagraph(p, ctx))
  }
  for (const tbl of ((container.tbl ?? []) as Record<string, unknown>[])) {
    blocks.push(await parseTable(tbl, ctx))
  }
  return blocks
}

export async function parseDocument(xml: string, ctx: ParseContext): Promise<Block[]> {
  const doc = parser.parse(xml) as Record<string, unknown>
  const body = (doc?.document as Record<string, unknown>)?.body as Record<string, unknown>
  if (!body) return []
  return parseBlockContainer(body, ctx)
}
