export type DocxDocument = {
  blocks: Block[]
  pageSize?: {
    widthPx: number
    heightPx: number
    marginPx: { top: number; right: number; bottom: number; left: number }
    footerPx?: number  // distance of the footer from the page bottom edge (w:footer)
    headerPx?: number  // distance of the header from the page top edge (w:header)
  }
  footnotes?: NoteEntry[]  // referenced footnotes, in document order (number 1..n)
  endnotes?: NoteEntry[]   // referenced endnotes, in document order (number 1..n)
  footer?: Block[]         // default page footer (w:footerReference); PAGE fields become page numbers
  header?: Block[]         // default page header (w:headerReference); PAGE fields become page numbers
  // Distinct first-page / even-page header & footer (w:type="first"/"even").
  // Used only when the corresponding flag below is set.
  headerFirst?: Block[]
  headerEven?: Block[]
  footerFirst?: Block[]
  footerEven?: Block[]
  titlePg?: boolean            // w:titlePg in the section: page 1 uses the "first" header/footer
  evenAndOddHeaders?: boolean  // w:evenAndOddHeaders in settings: even pages use the "even" header/footer
  // Document sections, each with its own page size/orientation (w:sectPr).
  // Present only when the document has MORE THAN ONE section; a single-section
  // document uses `blocks` + `pageSize` directly. blocks[] across sections
  // concatenate to `blocks`.
  sections?: Section[]
}

export type PageSize = NonNullable<DocxDocument['pageSize']>

export type Section = {
  blocks: Block[]
  pageSize: PageSize
  header?: Block[]  // this section's page header (its sectPr's w:headerReference); inherits the previous section's when it declares none
  footer?: Block[]  // this section's page footer (its sectPr's w:footerReference); inherits the previous section's when it declares none
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
  // Transient: set on the LAST paragraph of a section (its pPr held a w:sectPr)
  // to that section's page size. Consumed when building DocxDocument.sections;
  // stripped before the document is returned.
  sectionPageSize?: PageSize
  // Transient: the default header/footer relationship ids from the same sectPr,
  // resolved to blocks when building sections, then stripped.
  sectionRefs?: { headerRId?: string; footerRId?: string }
}

export type Run = TextRun | ImageRun

export type TextRun = {
  type: 'run'
  text: string
  style: ComputedStyle
  href?: string  // set when the run is inside a w:hyperlink (external URL or #bookmark)
  noteRef?: { type: 'footnote' | 'endnote'; number: number }  // a footnote/endnote marker
  pageNumber?: boolean  // a PAGE field — rendered as the current page number
  totalPages?: boolean  // a NUMPAGES field — rendered as the total page count
  lineBreak?: boolean   // a w:br (soft line break) — rendered as <br> after the text
  pageBreak?: boolean   // a w:br w:type="page" — splits the paragraph onto a new page (transient marker; never rendered)
  tabs?: number         // count of leading w:tab elements — rendered as spacers
  deleted?: boolean     // a tracked deletion (w:del/w:delText) — hidden unless showRevisions
  inserted?: boolean    // a tracked insertion (w:ins) — styled when showRevisions
}

// Options for rendering a DocxDocument to HTML.
export type RenderOptions = {
  // Show tracked changes: deletions struck through, insertions underlined.
  // Default false = the accepted/final view (deletions removed, insertions kept).
  showRevisions?: boolean
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
  columnWidths?: number[]  // px per column, from w:tblGrid gridCol widths
  cellPadding?: { top: number; right: number; bottom: number; left: number }  // px, from w:tblCellMar / w:tcMar
}

export type TableRow = {
  cells: TableCell[]
}

export type TableCell = {
  rowSpan: number   // 1 = no merge
  colSpan: number   // 1 = no merge
  blocks: Block[]
  backgroundColor?: string  // hex from w:shd fill (e.g. "ff6109"), no # prefix
  border?: CellBorders      // effective cell borders (table style ∪ tblBorders ∪ tcBorders)
}

// Per-side CSS border shorthands (e.g. "1px solid #000"). A side is absent when
// no border applies there. Resolved by cascade: table style → tblBorders → tcBorders.
export type CellBorders = {
  top?: string
  right?: string
  bottom?: string
  left?: string
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
  strike?: boolean       // w:strike — strikethrough
  vertAlign?: 'super' | 'sub'  // w:vertAlign — super/subscript
  fontSize?: number      // in points (pt); w:sz stores half-points → divide by 2
  fontFamily?: string    // from w:rFonts w:ascii (fallback w:hAnsi)
  color?: string         // hex e.g. "FF0000"; "auto" is filtered out at parse time
  alignment?: 'left' | 'center' | 'right' | 'justify'
  rtl?: boolean          // right-to-left: w:bidi (paragraph) or w:rtl (run) → dir="rtl"
  backgroundColor?: string  // hex from w:shd fill (e.g. "ff6109"), no # prefix
  highlight?: string     // CSS color name from w:highlight (e.g. "yellow")
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
  // Paragraph borders (from w:pBdr), as ready-to-use CSS border shorthands.
  borderTop?: string
  borderBottom?: string
  borderLeft?: string
  borderRight?: string
  // Tab stops (from w:tabs), sorted by position. Drives tab rendering — most
  // importantly right-aligned stops with dot leaders (tables of contents).
  tabStops?: TabStop[]
}

export type TabStop = {
  posPx: number  // position from the paragraph's left edge, px (w:pos twips → px)
  val: 'left' | 'right' | 'center' | 'decimal' | 'bar'
  leader: 'none' | 'dot' | 'hyphen' | 'underscore'
}

export class DocxParseError extends Error {
  code: 'MISSING_ENTRY' | 'INVALID_ZIP'
  constructor(message: string, code: 'MISSING_ENTRY' | 'INVALID_ZIP') {
    super(message)
    this.name = 'DocxParseError'
    this.code = code
  }
}
