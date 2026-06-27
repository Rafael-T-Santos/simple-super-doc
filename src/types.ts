export type DocxDocument = {
  blocks: Block[]
  pageSize?: { widthPx: number; heightPx: number; marginPx: { top: number; right: number; bottom: number; left: number } }
}

export type Block = ParagraphBlock | TableBlock

export type ParagraphBlock = {
  type: 'paragraph'
  style: ComputedStyle
  runs: Run[]
  list?: ListRef
}

export type Run = TextRun | ImageRun

export type TextRun = {
  type: 'run'
  text: string
  style: ComputedStyle
}

export type ImageRun = {
  type: 'image'
  src: string        // base64 data URL
  widthPx: number    // converted from EMU: Math.round(cx / 9525)
  heightPx: number
  isPageBackground?: boolean  // true when wp:anchor behindDoc="1" (full-page background)
}

export type TableBlock = {
  type: 'table'
  rows: TableRow[]
}

export type TableRow = {
  cells: TableCell[]
}

export type TableCell = {
  rowSpan: number   // 1 = no merge
  colSpan: number   // 1 = no merge
  blocks: Block[]
  backgroundColor?: string  // hex from w:shd fill (e.g. "ff6109"), no # prefix
}

// Counter increment is the renderer's responsibility.
// start is resolved at parse time from abstractNum + any lvlOverride.
export type ListRef = {
  numId: string
  ilvl: number      // only 0 supported in v0.1.0
  ordered: boolean  // true = numbered, false = bullet
  start: number
}

export type ComputedStyle = {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number      // in points (pt); w:sz stores half-points → divide by 2
  fontFamily?: string    // from w:rFonts w:ascii (fallback w:hAnsi)
  color?: string         // hex e.g. "FF0000"; "auto" is filtered out at parse time
  alignment?: 'left' | 'center' | 'right' | 'justify'
  backgroundColor?: string  // hex from w:shd fill (e.g. "ff6109"), no # prefix
}

export class DocxParseError extends Error {
  code: 'MISSING_ENTRY' | 'INVALID_ZIP'
  constructor(message: string, code: 'MISSING_ENTRY' | 'INVALID_ZIP') {
    super(message)
    this.name = 'DocxParseError'
    this.code = code
  }
}
