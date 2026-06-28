export type DocxDocument = {
  blocks: Block[]
  pageSize?: { widthPx: number; heightPx: number; marginPx: { top: number; right: number; bottom: number; left: number } }
  footnotes?: NoteEntry[]  // referenced footnotes, in document order (number 1..n)
  endnotes?: NoteEntry[]   // referenced endnotes, in document order (number 1..n)
}

// A footnote/endnote's resolved content. number matches the in-text marker.
export type NoteEntry = {
  number: number
  blocks: Block[]
}

export type Block = ParagraphBlock | TableBlock

export type ParagraphBlock = {
  type: 'paragraph'
  style: ComputedStyle
  runs: Run[]
  list?: ListRef
  pageBreakBefore?: boolean  // w:pageBreakBefore — forces a new page in paginated render
}

export type Run = TextRun | ImageRun

export type TextRun = {
  type: 'run'
  text: string
  style: ComputedStyle
  href?: string  // set when the run is inside a w:hyperlink (external URL or #bookmark)
  noteRef?: { type: 'footnote' | 'endnote'; number: number }  // a footnote/endnote marker
}

export type ImageRun = {
  type: 'image'
  src: string        // base64 data URL
  widthPx: number    // converted from EMU: Math.round(cx / 9525)
  heightPx: number
  isPageBackground?: boolean  // true when wp:anchor behindDoc="1" (full-page background)
  href?: string  // set when the image is inside a w:hyperlink
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
  ilvl: number      // nesting level (0 = top); nested levels render as nested lists
  ordered: boolean  // true = numbered, false = bullet
  start: number
  format: string    // OOXML w:numFmt val (decimal, lowerLetter, lowerRoman, bullet, ...)
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
  // Paragraph spacing (from w:spacing). Run-level styles ignore these.
  spaceBefore?: number   // px, from w:spacing w:before (twips)
  spaceAfter?: number    // px, from w:spacing w:after (twips)
  lineHeight?: number    // unitless multiplier, from w:spacing w:line (auto rule)
  lineHeightPx?: number  // fixed px, from w:spacing w:line (atLeast/exact rule)
  // Paragraph indentation (from w:ind), in px.
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number  // positive: first-line indent
  indentHanging?: number    // positive: hanging indent (first line out-dented)
}

export class DocxParseError extends Error {
  code: 'MISSING_ENTRY' | 'INVALID_ZIP'
  constructor(message: string, code: 'MISSING_ENTRY' | 'INVALID_ZIP') {
    super(message)
    this.name = 'DocxParseError'
    this.code = code
  }
}
