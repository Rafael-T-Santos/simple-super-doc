import type { DocxDocument, Block, ParagraphBlock, TableBlock, TextRun, ImageRun, Run, ComputedStyle, NoteEntry, TabStop, RenderOptions } from '../types.js'
import {
  EMPTY_LINE_EM, LINE_HEIGHT,
  extractPageBackground, isBlockVisible, isHeadingBlock, isIconOnly,
  fullPageImage, flowOnly, watermarksOf,
} from './layout.js'

function styleToCss(s: ComputedStyle): string {
  const parts: string[] = []
  if (s.bold) parts.push('font-weight:bold')
  if (s.italic) parts.push('font-style:italic')
  // underline and strikethrough combine into one text-decoration; a double
  // strikethrough (w:dstrike) uses the double line style.
  const deco = [s.underline ? 'underline' : '', s.strike || s.doubleStrike ? 'line-through' : ''].filter(Boolean)
  if (deco.length) parts.push(`text-decoration:${deco.join(' ')}${s.doubleStrike && !s.underline ? ' double' : ''}`)
  if (s.caps) parts.push('text-transform:uppercase')
  if (s.smallCaps) parts.push('font-variant:small-caps')
  if (s.vertAlign) {
    parts.push(`vertical-align:${s.vertAlign}`)
    parts.push('font-size:0.83em') // browsers shrink super/subscript text
  } else if (s.fontSize != null) {
    parts.push(`font-size:${s.fontSize}pt`)
  }
  if (s.fontFamily) parts.push(`font-family:${s.fontFamily},sans-serif`)
  if (s.color) parts.push(`color:#${s.color}`)
  if (s.alignment) parts.push(`text-align:${s.alignment}`)
  if (s.backgroundColor) parts.push(`background-color:#${s.backgroundColor}`)
  if (s.highlight) parts.push(`background-color:${s.highlight}`)
  return parts.join(';')
}

function renderRun(run: Run, parent: HTMLElement, skipLeadingTabs = false): void {
  // A run inside a hyperlink renders into an <a> wrapping the run's content.
  const href = (run as TextRun | ImageRun).href
  let target = parent
  if (href) {
    const a = document.createElement('a')
    // SECURITY: href is from the document's relationships; only allow safe URL
    // schemes (http/https/mailto/relative/#fragment), never javascript:.
    a.setAttribute('href', sanitizeHref(href))
    a.target = '_blank'
    a.rel = 'noopener noreferrer'
    parent.appendChild(a)
    target = a
  }

  if (run.type === 'image') {
    const ir = run as ImageRun
    const img = document.createElement('img')
    setImageSrc(img, ir.src)
    img.width = ir.widthPx
    img.height = ir.heightPx
    img.style.display = 'inline-block'
    img.style.maxWidth = '100%'
    // A positioned header/footer logo (letterhead): honor the anchor offset so
    // it lands where Word puts it (e.g. top-right) instead of flowing inline.
    if (inHeaderFooter && ir.anchorXPx !== undefined) {
      img.style.position = 'absolute'
      img.style.left = `${ir.anchorXPx}px`
      img.style.top = `${ir.anchorYPx ?? 0}px`
      img.style.maxWidth = 'none'
    }
    target.appendChild(img)
    return
  }

  const textRun = run as TextRun

  // Hidden text (w:vanish) is not shown in the final view (like Word/PDF export).
  if (textRun.style.hidden) return

  // Tracked deletion: hidden in the accepted/final view; shown struck through
  // (inside an <del>) when revisions are displayed.
  if (textRun.deleted) {
    if (!showRevisions) return
    const del = document.createElement('del')
    del.style.cssText = 'color:#c0392b'
    target.appendChild(del)
    target = del
  } else if (textRun.inserted && showRevisions) {
    // Tracked insertion: underlined and colored when revisions are displayed.
    const ins = document.createElement('ins')
    ins.style.cssText = 'color:#1d6f42;text-decoration:underline'
    target.appendChild(ins)
    target = ins
  }

  // PAGE field: render the current page number (set per page in renderFooter).
  if (textRun.pageNumber) {
    target.appendChild(document.createTextNode(String(currentPageNumber)))
    return
  }

  // NUMPAGES field: a placeholder filled with the total page count once the
  // document is fully paginated (see fillTotalPages).
  if (textRun.totalPages) {
    const span = document.createElement('span')
    span.dataset.ssdNumpages = '1'
    span.textContent = '1'
    target.appendChild(span)
    return
  }

  // Footnote/endnote marker: a superscript number linking to the notes section.
  if (textRun.noteRef) {
    const { type, number } = textRun.noteRef
    const prefix = type === 'footnote' ? 'fn' : 'en'
    const sup = document.createElement('sup')
    sup.id = `${prefix}ref-${number}`
    const a = document.createElement('a')
    a.setAttribute('href', `#${prefix}-${number}`)
    a.textContent = String(number)
    sup.appendChild(a)
    target.appendChild(sup)
    return
  }

  // Leading tab spacers (no tab-stop math, just visible separation). Skipped
  // when a tab-stop layout has already consumed the leading tabs as separators.
  for (let t = 0; !skipLeadingTabs && t < (textRun.tabs ?? 0); t++) {
    const sp = document.createElement('span')
    sp.style.cssText = 'display:inline-block;min-width:2.5em'
    target.appendChild(sp)
  }

  const css = styleToCss(textRun.style)

  if (textRun.text) {
    // A right-to-left run (w:rtl) needs dir="rtl" so a span is forced even with
    // no other styling.
    if (css || textRun.style.rtl) {
      const span = document.createElement('span')
      if (css) span.style.cssText = css
      if (textRun.style.rtl) span.dir = 'rtl'
      // SECURITY: use textContent, never innerHTML
      span.textContent = textRun.text
      target.appendChild(span)
    } else {
      target.appendChild(document.createTextNode(textRun.text))
    }
  }

  // A soft line break (w:br) renders after the run's text.
  if (textRun.lineBreak) target.appendChild(document.createElement('br'))
}

// Render the footnote and/or endnote sections at the end of the document. Each
// note is an <li> (numbered to match its in-text marker) with a back-reference.
function renderNotes(
  doc: DocxDocument,
  container: HTMLElement,
  opts: { footnotes?: boolean; endnotes?: boolean } = { footnotes: true, endnotes: true },
): void {
  const all: Array<{ prefix: string; label: string; notes: DocxDocument['footnotes'] }> = []
  if (opts.footnotes) all.push({ prefix: 'fn', label: 'Footnotes', notes: doc.footnotes })
  if (opts.endnotes) all.push({ prefix: 'en', label: 'Endnotes', notes: doc.endnotes })
  for (const { prefix, label, notes } of all) {
    if (!notes || notes.length === 0) continue
    const section = document.createElement('section')
    section.className = `ssd-${prefix === 'fn' ? 'footnotes' : 'endnotes'}`
    const hr = document.createElement('hr')
    section.appendChild(hr)
    const heading = document.createElement('h2')
    heading.textContent = label
    heading.style.cssText = 'font-size:1em'
    section.appendChild(heading)
    const ol = document.createElement('ol')
    for (const note of notes) {
      const li = document.createElement('li')
      li.id = `${prefix}-${note.number}`
      renderBlocks(note.blocks, li)
      const back = document.createElement('a')
      back.setAttribute('href', `#${prefix}ref-${note.number}`)
      back.textContent = ' ↩'
      li.appendChild(back)
      ol.appendChild(li)
    }
    section.appendChild(ol)
    container.appendChild(section)
  }
}

// Only allow safe URL schemes; neutralize javascript:/data: and other unsafe
// schemes to "#" so a malicious .docx can't inject a script URL.
function sanitizeHref(href: string): string {
  const trimmed = href.trim()
  // Relative URLs and #fragments are safe.
  if (/^(#|\/|\.|[^:]*$)/.test(trimmed)) return trimmed
  if (/^(https?:|mailto:|tel:)/i.test(trimmed)) return trimmed
  return '#'
}

// Only allow safe image sources: embedded data:image/* URLs, external http(s)
// images, and relative URLs. Anything else (javascript:, data:text/html, …) is
// rejected to an empty string so it is never set as an <img src>. Image src does
// not execute these schemes in browsers, but this is defense-in-depth: it keeps a
// malicious .docx from smuggling an unexpected URL scheme through an image.
function sanitizeImageSrc(src: string): string {
  const trimmed = src.trim()
  if (/^data:image\//i.test(trimmed)) return trimmed
  if (/^https?:\/\//i.test(trimmed)) return trimmed
  if (/^(\/|\.|[^:]*$)/.test(trimmed)) return trimmed // relative
  return ''
}

// Set an <img>'s src only when it passes sanitizeImageSrc; an unsafe/empty src is
// left unset (never src="") so the browser does not re-request the page URL.
function setImageSrc(img: HTMLImageElement, src: string): void {
  const safe = sanitizeImageSrc(src)
  if (safe) img.src = safe
}

// Empty-paragraph line height, set per render(): template covers (hasPageBg)
// need a taller empty line to match Word's layout (EMPTY_LINE_EM); plain
// documents use a normal single line (LINE_HEIGHT) so blank lines don't bloat
// the page. LINE_HEIGHT and EMPTY_LINE_EM live in ./layout.
let emptyLineEm = LINE_HEIGHT

// When true, tracked changes are shown (deletions struck through, insertions
// underlined); when false, the accepted/final view is rendered (deletions
// removed, insertions kept as plain text). Set per render() from RenderOptions.
let showRevisions = false

function ensureLineBox(el: HTMLElement): void {
  // Empty runs still create (empty) text nodes, so check for visible content
  // rather than child count. Images carry their own height.
  if (!el.textContent && !el.querySelector('img')) {
    el.style.minHeight = `${emptyLineEm}em`
  }
}

// Set while rendering inside a table cell: the cell's padding provides the
// spacing, so paragraph before/after margins are suppressed to keep rows compact.
let inTableCell = false

// Set while rendering a header/footer: those parts carry implicit default tab
// stops (a centered stop mid-page and a right stop at the margin) that Word uses
// for the classic "left <tab> center <tab> right" layout, so a tabbed paragraph
// with no explicit stops still right-aligns its trailing segment.
let inHeaderFooter = false

// The page number substituted for PAGE fields, set per page by renderFooter.
let currentPageNumber = 0

// Footnote-rule height budget (separator rule + gap) when reserving page space.
const FOOTNOTE_SEPARATOR_H = 12

// Footnote numbers referenced inside a rendered element (from <sup id="fnref-N">).
function footnoteNumbersIn(el: HTMLElement): number[] {
  return Array.from(el.querySelectorAll('sup[id^="fnref-"]'))
    .map(s => parseInt((s as HTMLElement).id.slice('fnref-'.length), 10))
    .filter(n => !Number.isNaN(n))
}

// One footnote at the bottom of a page: "N <content>" with a back-reference.
function buildFootnoteItem(fn: NoteEntry): HTMLElement {
  const item = document.createElement('div')
  item.style.cssText = 'display:flex;gap:6px;align-items:baseline'
  const num = document.createElement('span')
  num.id = `fn-${fn.number}`
  num.textContent = String(fn.number)
  num.style.flex = '0 0 auto'
  const content = document.createElement('div')
  renderBlocks(fn.blocks, content)
  const back = document.createElement('a')
  back.setAttribute('href', `#fnref-${fn.number}`)
  back.textContent = ' ↩'
  content.appendChild(back)
  item.append(num, content)
  return item
}

// Render the given footnotes at the bottom of a page (above the footer), with a
// short separator rule, matching Word's per-page footnote placement.
function renderPageFootnotes(footnotes: NoteEntry[], pageDiv: HTMLElement, numbers: number[], pm: PageMargins): void {
  if (numbers.length === 0) return
  const box = document.createElement('div')
  box.className = 'ssd-footnotes'
  box.style.cssText =
    `position:absolute;left:${pm.left}px;right:${pm.right}px;bottom:${pm.bottom}px;font-size:9pt;line-height:1.3`
  const sep = document.createElement('div')
  sep.style.cssText = 'border-top:1px solid #999;width:33%;margin-bottom:4px'
  box.appendChild(sep)
  for (const n of numbers) {
    const fn = footnotes.find(f => f.number === n)
    if (fn) box.appendChild(buildFootnoteItem(fn))
  }
  pageDiv.appendChild(box)
}

// Render the document's footer at the bottom of a page box, resolving PAGE
// fields to the given page number.
// Pick the header/footer for a page: the first-page variant on page 1 (titlePg),
// the even-page variant on even pages (evenAndOddHeaders), else the default.
function pickPart(
  doc: DocxDocument, pageNum: number,
  def: Block[] | undefined, first: Block[] | undefined, even: Block[] | undefined,
): Block[] | undefined {
  if (doc.titlePg && pageNum === 1 && first) return first
  if (doc.evenAndOddHeaders && pageNum % 2 === 0 && even) return even
  return def
}

function renderFooter(doc: DocxDocument, pageDiv: HTMLElement, pageNum: number, pm: PageMargins, footerPx: number): void {
  const footer = pickPart(doc, pageNum, doc.footer, doc.footerFirst, doc.footerEven)
  if (!footer || footer.length === 0) return
  const prev = currentPageNumber
  currentPageNumber = pageNum
  const el = document.createElement('div')
  el.className = 'ssd-footer'
  el.style.cssText = `position:absolute;left:${pm.left}px;right:${pm.right}px;bottom:${footerPx}px`
  const prevHF = inHeaderFooter
  inHeaderFooter = true
  renderBlocks(footer, el)
  inHeaderFooter = prevHF
  pageDiv.appendChild(el)
  currentPageNumber = prev
}

// Replace every NUMPAGES placeholder (rendered by renderRun) with the final
// total page count, once the document has been fully paginated.
function fillTotalPages(container: HTMLElement, total: number): void {
  container.querySelectorAll('[data-ssd-numpages]').forEach(el => { el.textContent = String(total) })
}

// Render the document's header at the top of a page box, resolving PAGE fields
// to the given page number. headerPx is the header's distance from the top edge.
function renderHeader(doc: DocxDocument, pageDiv: HTMLElement, pageNum: number, pm: PageMargins, headerPx: number): void {
  const header = pickPart(doc, pageNum, doc.header, doc.headerFirst, doc.headerEven)
  if (!header || header.length === 0) return
  const prev = currentPageNumber
  currentPageNumber = pageNum
  const el = document.createElement('div')
  el.className = 'ssd-header'
  el.style.cssText = `position:absolute;left:${pm.left}px;right:${pm.right}px;top:${headerPx}px`
  const prevHF = inHeaderFooter
  inHeaderFooter = true
  renderBlocks(header, el)
  inHeaderFooter = prevHF
  pageDiv.appendChild(el)
  currentPageNumber = prev
}

// CSS for a paragraph / list item: the document's w:spacing before/after as
// margins, its line spacing, indentation, and the run-level style. Replaces the
// browser's default 1em margins so the vertical rhythm matches the document.
// For list items, indentation is handled by the list's padding (skipIndent).
function paragraphCss(style: ComputedStyle, skipIndent = false): string {
  const mt = inTableCell ? 0 : style.spaceBefore ?? 0
  const mb = inTableCell ? 0 : style.spaceAfter ?? 0
  // Single spacing is the font's own line box — `normal`, not a number — and is
  // checked before the multiplier so an explicit single wins over an inherited
  // one. A document that says nothing falls back to LINE_HEIGHT as before.
  const lh =
    style.lineHeightPx != null ? `${style.lineHeightPx}px`
    : style.lineHeightSingle ? 'normal'
    : `${style.lineHeight ?? LINE_HEIGHT}`
  let css = `margin:${mt}px 0 ${mb}px;line-height:${lh}`
  if (!skipIndent) {
    if (style.indentLeft) css += `;padding-left:${style.indentLeft}px`
    if (style.indentRight) css += `;padding-right:${style.indentRight}px`
    if (style.indentFirstLine) css += `;text-indent:${style.indentFirstLine}px`
    else if (style.indentHanging) css += `;text-indent:${-style.indentHanging}px`
  }
  if (style.borderTop) css += `;border-top:${style.borderTop}`
  if (style.borderBottom) css += `;border-bottom:${style.borderBottom}`
  if (style.borderLeft) css += `;border-left:${style.borderLeft}`
  if (style.borderRight) css += `;border-right:${style.borderRight}`
  const inline = styleToCss(style)
  if (inline) css += ';' + inline
  return css
}

function renderParagraph(block: ParagraphBlock): HTMLElement {
  const p = document.createElement('p')
  p.style.cssText = paragraphCss(block.style)
  // Right-to-left paragraph (w:bidi): the browser then flows and right-aligns it.
  if (block.style.rtl) p.dir = 'rtl'

  // A right/center/decimal tab stop with leading tabs (a table-of-contents row:
  // "Title.....12") is laid out with a flex leader instead of a blank spacer.
  let stops = block.style.tabStops
  const hasTabRun = block.runs.some(r => r.type === 'run' && ((r as TextRun).tabs ?? 0) > 0)
  // In a header/footer, a tabbed paragraph with no explicit stops gets Word's
  // implicit center + right stops so "V.4 <tabs> CONFIDENCIAL" splits left/right
  // instead of drifting to the middle.
  if (inHeaderFooter && hasTabRun && !stops?.some(s => s.val === 'right' || s.val === 'center' || s.val === 'decimal')) {
    stops = [
      { posPx: 0, val: 'center', leader: 'none' },
      { posPx: 0, val: 'right', leader: 'none' },
    ]
  }
  const alignStop = stops?.some(s => s.val === 'right' || s.val === 'center' || s.val === 'decimal')
  if (hasTabRun && stops && alignStop) {
    renderTabbedParagraph(block, p, stops)
  } else {
    for (const run of block.runs) renderRun(run, p)
  }

  ensureLineBox(p)
  if (block.pageBreakBefore) p.dataset.ssdBreak = '1'
  return p
}

// Lay out a paragraph whose tabs align to explicit tab stops. The run sequence is
// split at each leading tab into segments; between segments a "leader" fills the
// gap (flex-grow for right/center/decimal stops, a fixed spacer for left stops),
// with optional dot/hyphen/underscore leader styling. This produces the dotted
// right-aligned page numbers of a table of contents.
function renderTabbedParagraph(block: ParagraphBlock, p: HTMLElement, stops: TabStop[]): void {
  p.style.display = 'flex'
  p.style.alignItems = 'baseline'
  p.style.width = '100%'

  const segments: HTMLElement[] = []
  const seps: TabStop[] = []
  let stopIdx = 0
  let current = document.createElement('span')
  current.style.cssText = 'flex:0 1 auto;min-width:0'
  segments.push(current)

  for (const run of block.runs) {
    let leadTabs = run.type === 'run' ? ((run as TextRun).tabs ?? 0) : 0
    while (leadTabs-- > 0) {
      seps.push(stops[Math.min(stopIdx, stops.length - 1)])
      stopIdx++
      current = document.createElement('span')
      current.style.cssText = 'flex:0 1 auto;min-width:0'
      segments.push(current)
    }
    renderRun(run, current, /* skipLeadingTabs */ true)
  }

  // The last segment (e.g. the page number) should never wrap or shrink.
  const last = segments[segments.length - 1]
  if (segments.length > 1) last.style.cssText = 'flex:0 0 auto;white-space:nowrap'

  p.appendChild(segments[0])
  for (let i = 1; i < segments.length; i++) {
    p.appendChild(makeLeader(seps[i - 1]))
    p.appendChild(segments[i])
  }
}

// The filler between two tab-stop segments. Right/center/decimal stops grow to
// push the next segment toward the stop; left/bar stops use a fixed spacer.
function makeLeader(stop: TabStop): HTMLElement {
  const el = document.createElement('span')
  const grows = stop.val === 'right' || stop.val === 'center' || stop.val === 'decimal'
  const border =
    stop.leader === 'dot' ? 'border-bottom:2px dotted currentColor'
    : stop.leader === 'hyphen' ? 'border-bottom:1px dashed currentColor'
    : stop.leader === 'underscore' ? 'border-bottom:1px solid currentColor'
    : ''
  if (grows) {
    el.style.cssText = `flex:1 1 0;margin:0 4px;align-self:flex-end;transform:translateY(-0.35em);${border}`
  } else {
    el.style.cssText = 'display:inline-block;flex:0 0 auto;min-width:2.5em'
  }
  return el
}

function renderTable(block: TableBlock, container: HTMLElement): void {
  const table = document.createElement('table')
  table.style.borderCollapse = 'collapse'
  table.style.maxWidth = '100%'

  // Apply the document's column widths (w:tblGrid) with a fixed layout so columns
  // — and therefore text wrapping and row heights — match the document. Use the
  // ABSOLUTE widths (not 100%) so a table narrower than the page keeps its real
  // width and alignment (w:jc) instead of being stretched full-width; maxWidth
  // 100% still caps a table wider than its container. Falls back to content-based
  // sizing when the document gives no grid.
  if (block.columnWidths && block.columnWidths.length > 0) {
    const total = block.columnWidths.reduce((a, b) => a + b, 0)
    if (total > 0) {
      table.style.tableLayout = 'fixed'
      table.style.width = `${total}px`
      if (block.align === 'center') table.style.margin = '0 auto'
      else if (block.align === 'right') table.style.marginLeft = 'auto'
      const colgroup = document.createElement('colgroup')
      for (const w of block.columnWidths) {
        const col = document.createElement('col')
        col.style.width = `${w}px`
        colgroup.appendChild(col)
      }
      table.appendChild(colgroup)
    }
  }

  const pad = block.cellPadding
  for (const row of block.rows) {
    const tr = document.createElement('tr')
    // w:tblHeader: marked so a continuation piece can repeat it (see
    // buildContinuationTable). A data attribute rather than <thead> keeps the
    // existing flat row structure the paginator walks.
    if (row.isHeader) tr.dataset.ssdHeader = '1'
    // w:trHeight: `height` on a table row acts as a MINIMUM (the row still grows
    // to fit content), which matches the common "atLeast" rule.
    if (row.heightPx) tr.style.height = `${row.heightPx}px`
    for (const cell of row.cells) {
      const td = document.createElement('td')
      if (cell.colSpan > 1) td.colSpan = cell.colSpan
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan
      if (cell.backgroundColor) td.style.backgroundColor = `#${cell.backgroundColor}`
      // Cell borders resolved by cascade (table style → tblBorders → tcBorders).
      if (cell.border) {
        if (cell.border.top) td.style.borderTop = cell.border.top
        if (cell.border.right) td.style.borderRight = cell.border.right
        if (cell.border.bottom) td.style.borderBottom = cell.border.bottom
        if (cell.border.left) td.style.borderLeft = cell.border.left
      }
      // Cell padding from the document's w:tcMar (keeps rows as compact as Word).
      if (pad) td.style.padding = `${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px`
      // w:vAlign; Word's default is top, which is also what an unset cell gets.
      td.style.verticalAlign =
        cell.verticalAlign === 'center' ? 'middle' : cell.verticalAlign === 'bottom' ? 'bottom' : 'top'
      // Let long template tokens wrap instead of forcing overflow.
      td.style.overflowWrap = 'break-word'
      td.style.wordBreak = 'break-word'
      // w:textDirection: writing-mode goes on the CELL, not on a wrapper, so the
      // table sizing algorithm measures the rotated text and the row grows to fit
      // it (a wrapper would be sized by a row height that does not know about it).
      // btLr reads bottom-to-top, which is vertical-rl turned 180°; that rotation
      // has to live on an inner element, because rotating the cell itself would
      // also flip its background and swap its top/bottom borders.
      let target: HTMLElement = td
      if (cell.textDirection) {
        td.style.writingMode = 'vertical-rl'
        if (cell.textDirection === 'btLr') {
          const rot = document.createElement('div')
          rot.style.transform = 'rotate(180deg)'
          td.appendChild(rot)
          target = rot
        }
      }
      // Cell padding already provides the spacing, so don't add the paragraph's
      // before/after margins inside cells (Word keeps cell content tight).
      const prev = inTableCell
      inTableCell = true
      renderBlocks(cell.blocks, target)
      inTableCell = prev
      tr.appendChild(td)
    }
    table.appendChild(tr)
  }
  container.appendChild(table)
}

// OOXML w:numFmt -> CSS list-style-type so nested levels show a/b/c, i/ii/iii,
// etc. instead of every <ol> defaulting to decimal.
const LIST_STYLE: Record<string, string> = {
  decimal: 'decimal',
  decimalZero: 'decimal-leading-zero',
  lowerLetter: 'lower-alpha',
  upperLetter: 'upper-alpha',
  lowerRoman: 'lower-roman',
  upperRoman: 'upper-roman',
  bullet: 'disc',
  none: 'none',
}

// The CSS list-style-type for a list level. For bullets, honor the document's
// literal w:lvlText (e.g. "-" or "*") when it's plain printable ASCII, so a
// hyphen bullet renders as "-" and not a generic "•". Symbol-font glyphs
// (Wingdings, U+2022, …) are non-ASCII and need their own font, so they fall
// back to the generic `disc` rather than rendering as tofu.
function listStyleTypeFor(format: string, bulletText?: string): string | undefined {
  if (format === 'bullet' && bulletText && /^[\x20-\x7E]+$/.test(bulletText)) {
    return `"${bulletText.replace(/[\\"]/g, '\\$&')}"`
  }
  return LIST_STYLE[format]
}

// One open list level while rendering nested lists. The stack's depth tracks
// the current w:ilvl; index i holds the <ol>/<ul> for level i.
type ListFrame = { el: HTMLElement; numId: string; lastLi: HTMLElement | null }

function renderBlocks(blocks: Block[], container: HTMLElement): void {
  // The list stack: stack[i] is the open list at ilvl i. stack[0] is the root
  // list (appended to container when the whole group closes); deeper levels are
  // nested inside the parent level's most recent <li>.
  let stack: ListFrame[] = []

  const closeLists = (): void => {
    if (stack.length > 0) container.appendChild(stack[0].el)
    stack = []
  }

  for (const block of blocks) {
    if (block.type === 'paragraph' && block.list) {
      const { numId, ordered, start, ilvl, format, bulletText } = block.list

      // A different list (different numId) at the base ends the previous one.
      if (stack.length > 0 && stack[0].numId !== numId) closeLists()

      // Close deeper levels until the stack is exactly ilvl+1 deep or shorter.
      while (stack.length > ilvl + 1) stack.pop()

      // Open nested lists until the stack reaches ilvl+1 deep. Intermediate
      // levels (when an item jumps more than one level) inherit this item's
      // format/start, which is rare and good enough.
      while (stack.length < ilvl + 1) {
        const listEl = document.createElement(ordered ? 'ol' : 'ul')
        if (ordered) (listEl as HTMLOListElement).start = start
        const styleType = listStyleTypeFor(format, bulletText)
        if (styleType) listEl.style.listStyleType = styleType
        // Indent from the item's w:ind (the marker hangs in this padding);
        // fall back to a modest default when the document doesn't specify one.
        const indent = block.style.indentLeft ?? 24
        listEl.style.margin = '0'
        listEl.style.paddingLeft = `${indent}px`
        const parent = stack[stack.length - 1]
        if (parent) (parent.lastLi ?? parent.el).appendChild(listEl)
        stack.push({ el: listEl, numId, lastLi: null })
      }

      const frame = stack[stack.length - 1]
      const li = document.createElement('li')
      li.style.cssText = paragraphCss(block.style, true) // indent handled by the list padding
      for (const run of block.runs) renderRun(run, li)
      frame.el.appendChild(li)
      frame.lastLi = li
    } else {
      closeLists()

      if (block.type === 'paragraph') {
        container.appendChild(renderParagraph(block))
      } else if (block.type === 'table') {
        renderTable(block, container)
      }
    }
  }

  closeLists()
}

// A block the paginator can cut in two when it does not fit the space left on a
// page: a table splits between rows, a list between items, and a paragraph
// between lines. A table CELL is deliberately not split — a tall single row
// still moves whole, because cutting inside a cell means rebuilding the row.
function splitterFor(el: HTMLElement): ((el: HTMLElement, availH: number) => HTMLElement | null) | null {
  if (el.tagName === 'TABLE') return (e, h) => splitTableRows(e as HTMLTableElement, h)
  if (el.tagName === 'OL' || el.tagName === 'UL') return splitListItems
  if (el.tagName === 'P') return splitInlineAtHeight
  return null
}

// getClientRects returns one rect per inline FRAGMENT, not per line: a line
// broken across three <span>s yields three rects, and a <sup> on that line adds
// a fourth with a different top. Count real line boxes by merging rects that
// overlap vertically. Counting rects instead reported 15 lines for a 44px
// element, which made the widow/orphan guard meaningless.
function countLines(rects: DOMRectList): number {
  const boxes = Array.from(rects).filter(r => r.height > 0).sort((a, b) => a.top - b.top)
  if (boxes.length === 0) return 0
  let lines = 1
  let lineBottom = boxes[0].bottom
  for (let i = 1; i < boxes.length; i++) {
    if (boxes[i].top < lineBottom - 1) {
      lineBottom = Math.max(lineBottom, boxes[i].bottom) // same line, taller run
      continue
    }
    lines++
    lineBottom = boxes[i].bottom
  }
  return lines
}

// Word's default widow/orphan control: never strand fewer than this many lines
// of a paragraph alone on either side of a page break. Below that, moving the
// whole block reads better than the split.
const MIN_LINES_PER_PIECE = 2

// Every text node inside `el`, in document order, with the running character
// offset each one starts at. The offsets let a binary search address a position
// in the element's text as a single number.
function textPositions(el: HTMLElement): { nodes: Text[]; starts: number[]; total: number } {
  const nodes: Text[] = []
  const starts: number[] = []
  let total = 0
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const t = n as Text
    nodes.push(t)
    starts.push(total)
    total += t.data.length
  }
  return { nodes, starts, total }
}

// Place a Range boundary at global character offset `at`.
function boundaryAt(
  pos: { nodes: Text[]; starts: number[] },
  at: number,
): { node: Text; offset: number } | null {
  if (pos.nodes.length === 0) return null
  // The last node whose start is <= at; the offset is the remainder inside it.
  let lo = 0, hi = pos.nodes.length - 1
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1
    if (pos.starts[mid] <= at) lo = mid
    else hi = mid - 1
  }
  return { node: pos.nodes[lo], offset: Math.min(at - pos.starts[lo], pos.nodes[lo].data.length) }
}

// Split a block's inline content at the last line box that fits within availH:
// the lines that fit stay in `el`, the rest move into a clone that is returned.
// This is what lets a 100-line list item continue on the next page instead of
// jumping whole and leaving most of a page blank.
//
// The tree is cut with Range.extractContents, which rebuilds the ancestor chain
// on the far side of the cut — so a split in the middle of a <span>, an <a> or a
// tracked-change <ins> keeps its formatting on both pieces. Doing this by hand
// would mean reimplementing that.
//
// Returns null when the split is not worth making: no line fits, nothing would
// move, or either piece would be left under MIN_LINES_PER_PIECE. The caller
// then moves the whole block to a fresh page, exactly as before.
function splitInlineAtHeight(el: HTMLElement, availH: number): HTMLElement | null {
  const pos = textPositions(el)
  if (pos.total === 0) return null
  const top = el.getBoundingClientRect().top
  // The lines are not the whole box: padding, border and the margin below still
  // have to fit, or the piece that stays overflows the page by exactly that much
  // (measured: 11px, the list item's bottom margin).
  const cs = getComputedStyle(el)
  const below =
    parseFloat(cs.paddingBottom) + parseFloat(cs.borderBottomWidth) + parseFloat(cs.marginBottom)
  // A Range's rects cover the TEXT box, which is shorter than the LINE box
  // whenever line-height exceeds the font size. Laying out against the text
  // bottom therefore cuts a few px too late and the page overflows by the
  // leading (measured: 3px). Take the difference between the element's line
  // height and its tallest text rect as that leading.
  const whole = document.createRange()
  whole.selectNodeContents(el)
  const allRects = whole.getClientRects()
  // Reduce rather than Math.max(...spread): a 100-line item runs to thousands of
  // inline fragments, and spreading that many arguments is how you get a
  // RangeError on a document that is merely long.
  let textH = 0
  for (let i = 0; i < allRects.length; i++) textH = Math.max(textH, allRects[i].height)
  const lineH = parseFloat(cs.lineHeight)
  const leading = Number.isFinite(lineH) ? Math.max(0, lineH - textH) : 0
  availH -= below + leading
  if (availH <= 0) return null

  const probe = document.createRange()
  // Height of the content up to global offset `at`, measured from the block's
  // own top so the block's padding and the first line's leading are included.
  const heightUpTo = (at: number): number => {
    const b = boundaryAt(pos, at)
    if (!b) return Infinity
    probe.setStart(pos.nodes[0], 0)
    probe.setEnd(b.node, b.offset)
    const rects = probe.getClientRects()
    if (rects.length === 0) return 0
    // The LOWEST rect, not the last one: getClientRects is in document order, so
    // a subscript or an inline image that hangs below the baseline can sit
    // earlier in the list than the text that follows it. Taking the last rect
    // would under-report the height there and cut a line too late.
    let bottom = -Infinity
    for (let i = 0; i < rects.length; i++) bottom = Math.max(bottom, rects[i].bottom)
    return bottom - top
  }

  // Largest offset that still fits. Binary search: line boxes grow monotonically
  // with the offset, so the predicate is monotone even though the text is not.
  let lo = 0, hi = pos.total
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1
    if (heightUpTo(mid) <= availH) lo = mid
    else hi = mid - 1
  }
  if (lo <= 0 || lo >= pos.total) return null

  // Back up to a word boundary so a word is never cut in half. The whitespace
  // itself goes with the first piece, matching how a line break consumes it.
  let cut = lo
  while (cut > 0 && !/\s/.test(textAt(pos, cut - 1))) cut--
  if (cut <= 0) return null

  const b = boundaryAt(pos, cut)
  if (!b) return null

  // Widow/orphan guard, measured in real line boxes on both sides of the cut.
  probe.setStart(pos.nodes[0], 0)
  probe.setEnd(b.node, b.offset)
  const linesKept = countLines(probe.getClientRects())
  probe.setStart(b.node, b.offset)
  probe.setEnd(pos.nodes[pos.nodes.length - 1], pos.nodes[pos.nodes.length - 1].data.length)
  const linesMoved = countLines(probe.getClientRects())
  if (linesKept < MIN_LINES_PER_PIECE || linesMoved < MIN_LINES_PER_PIECE) return null

  const rest = el.cloneNode(false) as HTMLElement
  const cutRange = document.createRange()
  cutRange.setStart(b.node, b.offset)
  cutRange.setEndAfter(el.lastChild!)
  rest.appendChild(cutRange.extractContents())
  if (!(rest.textContent ?? '').trim()) return null
  // A continued paragraph must not repeat the space above it at the top of the
  // next page, and a continued list item must not repeat its marker.
  rest.style.marginTop = '0'
  if (rest.tagName === 'LI') rest.style.listStyle = 'none'
  return rest
}

// The character at a global offset, or '' past the end.
function textAt(pos: { nodes: Text[]; starts: number[] }, at: number): string {
  const b = boundaryAt(pos, at)
  return b ? b.node.data.charAt(b.offset) : ''
}

// Split a list's items so the ones that fit within availH stay in the original
// list; the overflow moves into a continuation list that is returned. Mirrors
// splitTableRows, including its contract: null when nothing useful can be split
// off (not even one item fits, or everything already fits), and the list must be
// in the DOM so item heights can be measured.
function splitListItems(list: HTMLElement, availH: number): HTMLElement | null {
  const kids = Array.from(list.children) as HTMLElement[]
  const listTop = list.getBoundingClientRect().top
  let splitAt = 0 // index the continuation starts at; 0 means nothing can be cut
  for (let i = 0; i < kids.length; i++) {
    // Measure from the list's own top edge rather than summing offsetHeight:
    // offsetHeight excludes an item's margins, and paragraphCss gives list items
    // real margins, so summing underestimates the list and overfills the page
    // (pages visibly grew past their min-height before this was measured from
    // rects). Table rows have no margins, which is why splitTableRows can sum.
    // Include the item's own bottom margin. It collapses out of the list's box,
    // so the rect stops short of it, but the space is still taken on the page —
    // cutting after this item and ignoring it overflowed the page by exactly
    // that margin (7px, and the page silently stretched to absorb it).
    const consumed =
      kids[i].getBoundingClientRect().bottom - listTop +
      parseFloat(getComputedStyle(kids[i]).marginBottom)
    // No "the first item always fits" exemption, unlike splitTableRows. When not
    // even one item fits the space left, returning null lets the caller move the
    // whole list to a fresh page and try again with a full page of budget;
    // keeping the item here would silently overflow the page box instead. This
    // still terminates: on that fresh page the caller accepts the overflow
    // (used === 0) when one item is taller than a whole page.
    if (consumed > availH) break
    // Only cut where the NEXT child opens a new item. renderBlocks attaches a
    // nested list to `parent.lastLi ?? parent.el`, so a document whose list
    // starts at a deeper level puts a bare <ol> directly under this one; that
    // nested list has to travel with the item it belongs to, never alone.
    if (kids[i + 1] === undefined || kids[i + 1].tagName === 'LI') splitAt = i + 1
  }
  if (splitAt >= kids.length) return null // everything already fits

  // The first item that did not fit can still give up its opening lines. Without
  // this a 100-line item jumps whole and leaves most of the page blank, which is
  // the case that motivated line splitting in the first place.
  let tail: HTMLElement | null = null
  const firstMoving = kids[splitAt]
  if (firstMoving) {
    const itemTop = firstMoving.getBoundingClientRect().top - listTop
    const budget = availH - itemTop
    // A list whose only child is another list is what a document that starts at
    // a deeper level produces (renderBlocks attaches it to `parent.lastLi ??
    // parent.el`). Recursing is what makes THAT case splittable: the outer list
    // has no item of its own to cut.
    if (firstMoving.tagName === 'LI') tail = splitInlineAtHeight(firstMoving, budget)
    else if (firstMoving.tagName === 'OL' || firstMoving.tagName === 'UL') {
      tail = splitListItems(firstMoving, budget)
    }
  }
  if (splitAt === 0 && !tail) return null

  const rest = buildContinuationList(list, splitAt)
  if (tail) rest.appendChild(tail)
  // When the item was line-split its head stays behind, so the continuation
  // picks up at the item AFTER it.
  for (let i = tail ? splitAt + 1 : splitAt; i < kids.length; i++) rest.appendChild(kids[i])
  return rest
}

// Everything a continued list piece must inherit. cloneNode(false) carries the
// attributes (list-style-type, padding) but not the children, which is what we
// want. The counter is the per-piece state that is easy to lose: an <ol>
// continuation without `start` restarts at 1 on the next page, so a contract's
// clause "j." would come back as "a.".
function buildContinuationList(list: HTMLElement, splitAt: number): HTMLElement {
  const rest = list.cloneNode(false) as HTMLElement
  if (list.tagName === 'OL') {
    // Only <li> children advance an ordered list's counter, so a nested list
    // left behind must not be counted. `splitAt` is the index of the first item
    // that moves, which is also the count of numbers already used — and it holds
    // for a line-split item too: its tail leads the continuation with the marker
    // hidden, occupying the number the head already showed, so the next real
    // item still lands on the following number.
    const consumed = Array.from(list.children).slice(0, splitAt).filter(c => c.tagName === 'LI').length
    // `.start` already reports 1 when the attribute is absent, so no `|| 1`
    // fallback: numbering.ts passes w:start through verbatim, and a list that
    // legitimately starts at 0 would have its continuation shifted by one.
    ;(rest as HTMLOListElement).start = (list as HTMLOListElement).start + consumed
  }
  return rest
}

// Split a table's rows so the rows that fit within availH stay in the original
// table; the overflow rows move into a new table (same element/styles) that is
// returned. Returns null when nothing useful can be split off (no row fits the
// remaining space, or every row already fits). The table must be laid out (in
// the DOM) so row heights can be measured.
function splitTableRows(table: HTMLTableElement, availH: number): HTMLTableElement | null {
  const rows = Array.from(table.rows)
  let used = 0
  let splitAt = 0
  for (let i = 0; i < rows.length; i++) {
    const rh = rows[i].offsetHeight
    if (used > 0 && used + rh > availH) break
    used += rh
    splitAt = i + 1
  }
  if (splitAt === 0 || splitAt >= rows.length) return null
  const rest = buildContinuationTable(table, splitAt)
  for (let i = splitAt; i < rows.length; i++) rest.appendChild(rows[i])
  return rest
}

// The leading w:tblHeader rows of a rendered table, in order. Only a LEADING run
// of header rows repeats — a header flag on a row in the middle of the body is
// not a heading Word would carry to the next page.
function headerRowsOf(table: HTMLTableElement): HTMLTableRowElement[] {
  const heads: HTMLTableRowElement[] = []
  for (const row of Array.from(table.rows)) {
    if (row.dataset.ssdHeader !== '1') break
    heads.push(row)
  }
  return heads
}

// Everything a continued piece must inherit from the table it was split from.
// This is the ONE place to add the next such thing — the split path has lost
// per-piece state three times (footnote refs, column widths, heading rows), each
// time because a new ad hoc clone forgot something.
//
// `splitAt` is where the body was cut, and it decides whether heading rows are
// repeated. Repeating them is only correct — and only TERMINATES — when the cut
// lands past them: the caller loops until the continuation is strictly smaller,
// so prepending N header rows to a piece that already starts inside the header
// block would grow it back and spin forever.
function buildContinuationTable(table: HTMLTableElement, splitAt: number): HTMLTableElement {
  const rest = table.cloneNode(false) as HTMLTableElement

  // cloneNode(false) copies the table's ATTRIBUTES (so table-layout:fixed and the
  // absolute width survive) but NOT its children — including <colgroup>. A fixed
  // layout with no column definitions splits the width EQUALLY, so without this
  // the continued piece lost the document's column widths. Deep, and before the
  // rows: HTML requires colgroup first.
  const colgroup = table.querySelector('colgroup')
  if (colgroup) rest.appendChild(colgroup.cloneNode(true))

  const heads = headerRowsOf(table)
  if (splitAt > heads.length) {
    for (const head of heads) {
      const copy = head.cloneNode(true) as HTMLTableRowElement
      // A repeated heading must not re-register its footnotes: the id is what
      // footnoteNumbersIn counts and what the note's back-link targets, so a
      // clone would render the note again on every continuation page and
      // duplicate the anchor. The visible marker stays; only the id goes.
      const refs = copy.querySelectorAll('sup[id^="fnref-"], sup[id^="enref-"]')
      for (const sup of Array.from(refs)) sup.removeAttribute('id')
      rest.appendChild(copy)
    }
  }
  return rest
}

type PageMargins = { top: number; right: number; bottom: number; left: number }

// Page box styling shared by paginated renders (white sheet on the host bg).
// flex-shrink:0 keeps a flex host (e.g. a centered column viewport) from
// collapsing a page that has no in-flow content (such as a full-bleed image).
function pageBoxStyle(pw: number, ph: number, pm: PageMargins, extra: string[] = []): string {
  return [
    'position:relative',
    'box-sizing:border-box',
    'background-color:#fff',
    'margin:0 auto 16px',
    'box-shadow:0 2px 12px rgba(0,0,0,.25)',
    'flex-shrink:0',
    `width:${pw}px`,
    `min-height:${ph}px`,
    `padding:${pm.top}px ${pm.right}px ${pm.bottom}px ${pm.left}px`,
    ...extra,
  ].join(';')
}

// Paginate a plain document (no full-page background) by rendering the whole
// flow once and distributing the resulting top-level elements into page boxes by
// their real laid-out heights. This is exact — it accounts for margin collapsing
// and grouped lists, which a per-block measurement cannot.
function renderPlainPaginated(
  doc: DocxDocument, container: HTMLElement, pw: number, ph: number, pm: PageMargins,
  pageOffset = 0,
): void {
  const contentW = pw - pm.left - pm.right
  const contentH = ph - pm.top - pm.bottom

  // Stage the full flow hidden, in the page's font/width context.
  const stage = document.createElement('div')
  stage.className = 'ssd-page'
  stage.style.cssText =
    `position:absolute;left:-9999px;top:0;visibility:hidden;width:${contentW}px;` +
    `padding:0;margin:0;min-height:0;box-shadow:none;background:none`
  container.appendChild(stage)
  renderBlocks(doc.blocks, stage)

  // Effective height of each top-level element = distance to the next sibling's
  // top (captures collapsed margins); the last uses the stage's full height.
  const children = Array.from(stage.children) as HTMLElement[]
  const tops = children.map(c => c.offsetTop)
  const stageH = stage.scrollHeight
  const heights = children.map((_, i) => (i + 1 < children.length ? tops[i + 1] : stageH) - tops[i])

  const footnotes = doc.footnotes ?? []
  // Pre-measure each footnote at the page content width so we can reserve space.
  const footnoteH: Record<number, number> = {}
  for (const fn of footnotes) {
    const item = buildFootnoteItem(fn)
    item.style.position = 'absolute'
    item.style.visibility = 'hidden'
    item.style.width = `${contentW}px`
    stage.appendChild(item)
    footnoteH[fn.number] = item.offsetHeight
    stage.removeChild(item)
  }

  const pages: HTMLElement[] = []
  const pageFootnotes: number[][] = []
  let pageDiv: HTMLElement
  let used = 0       // content height consumed on the current page
  let reserve = 0    // height reserved at the bottom for this page's footnotes
  let pageHasFn = false

  const newPage = (): HTMLElement => {
    const div = document.createElement('div')
    div.className = 'ssd-page'
    div.style.cssText = pageBoxStyle(pw, ph, pm)
    container.appendChild(div)
    pages.push(div)
    pageFootnotes.push([])
    used = 0; reserve = 0; pageHasFn = false
    return div
  }
  // Space a set of footnotes adds to the page bottom (+ separator if first ones).
  const fnReserve = (nums: number[], hasFn: boolean): number =>
    nums.length === 0 ? 0 : (hasFn ? 0 : FOOTNOTE_SEPARATOR_H) + nums.reduce((a, n) => a + (footnoteH[n] ?? 0), 0)
  const recordFootnotes = (nums: number[], add: number): void => {
    if (nums.length === 0) return
    pageFootnotes[pages.length - 1].push(...nums)
    reserve += add
    pageHasFn = true
  }

  // Reused for every fit measurement: a zero-height, zero-margin marker whose
  // position reports where the next block would start on the current page.
  const fitProbe = document.createElement('div')
  fitProbe.style.cssText = 'height:0;margin:0;padding:0;border:0'

  pageDiv = newPage()
  for (let i = 0; i < children.length; i++) {
    const child = children[i]
    const forced = child.dataset.ssdBreak === '1'

    // Tables and lists split across pages: the rows/items that fit stay, the
    // rest continue on the next page (and split again if still too tall).
    const split = splitterFor(child)
    if (split) {
      if (forced && used > 0) pageDiv = newPage()
      let piece = child
      for (;;) {
        pageDiv.appendChild(piece)
        // Space actually consumed on this page, measured with a zero-height
        // sentinel placed after the piece: it lands exactly where the next block
        // would start, so margins that collapse OUT of the piece's own box are
        // counted. `piece.offsetHeight` is the border box and misses them — a
        // list whose last item carries a bottom margin fit "by the box" and then
        // overflowed the page by that margin, silently, since the page is sized
        // with min-height.
        pageDiv.appendChild(fitProbe)
        const usedAfter = fitProbe.offsetTop - pm.top
        pageDiv.removeChild(fitProbe)
        // Reserve space for footnotes referenced by the rows/items about to land
        // on this page: Word puts a footnote at the bottom of the page that
        // holds its reference, so the piece must break early enough to leave
        // room. Use the whole remaining piece's refs as an upper bound for the
        // split budget (safe: never overlaps), then record only the refs that
        // actually stayed on this page.
        const pieceRefs = footnoteNumbersIn(piece)
        const pieceReserve = fnReserve(pieceRefs, pageHasFn)
        if (usedAfter <= contentH - reserve - pieceReserve) {
          used = usedAfter
          recordFootnotes(pieceRefs, pieceReserve)
          break // whole remaining piece fits (with room for its footnotes)
        }
        const rest = split(piece, contentH - reserve - pieceReserve - used)
        if (!rest) {
          if (used === 0) {
            used = usedAfter
            recordFootnotes(pieceRefs, pieceReserve)
            break // taller than a full page; accept overflow
          }
          pageDiv.removeChild(piece)
          pageDiv = newPage()
          continue // retry on a fresh page
        }
        // `piece` now holds what fits on this page; record its footnotes here,
        // then continue with `rest` on a new page.
        const fitRefs = footnoteNumbersIn(piece)
        recordFootnotes(fitRefs, fnReserve(fitRefs, pageHasFn))
        pageDiv = newPage()
        piece = rest
      }
      continue
    }

    const h = heights[i]
    const refs = footnoteNumbersIn(child)
    let add = fnReserve(refs, pageHasFn)
    if ((forced && used > 0) || (used > 0 && used + h > contentH - reserve - add)) {
      pageDiv = newPage()
      add = fnReserve(refs, false) // fresh page: no footnotes yet, so include separator
    }
    pageDiv.appendChild(child) // moves the node out of the stage
    used += h
    recordFootnotes(refs, add)
  }
  stage.remove()

  // Drop blank pages (only empty paragraphs, no footnotes) — e.g. trailing empty
  // paragraphs pushed onto a new page by footnote reservation.
  for (let i = pages.length - 1; i >= 0; i--) {
    const hasFn = pageFootnotes[i].length > 0
    const hasContent = (pages[i].textContent ?? '').trim().length > 0 || !!pages[i].querySelector('img,table')
    if (!hasFn && !hasContent) {
      pages[i].remove()
      pages.splice(i, 1)
      pageFootnotes.splice(i, 1)
    }
  }

  // Footnotes at the bottom of the page that holds their reference (like Word).
  pages.forEach((page, i) => renderPageFootnotes(footnotes, page, pageFootnotes[i], pm))

  // Endnotes (if any) go to their own final page at the end of the document.
  if (doc.endnotes && doc.endnotes.length > 0) {
    const endHost = document.createElement('div')
    renderNotes(doc, endHost, { endnotes: true })
    if (endHost.childNodes.length > 0) {
      const div = newPage()
      while (endHost.firstChild) div.appendChild(endHost.firstChild)
    }
  }

  // Page footer / header (with page numbers) on every page. pageOffset
  // continues the page count across earlier sections.
  const footerPx = doc.pageSize?.footerPx ?? Math.round(pm.bottom / 2)
  pages.forEach((page, i) => renderFooter(doc, page, pageOffset + i + 1, pm, footerPx))
  const headerPx = doc.pageSize?.headerPx ?? Math.round(pm.top / 2)
  pages.forEach((page, i) => renderHeader(doc, page, pageOffset + i + 1, pm, headerPx))

  // NUMPAGES = total rendered pages so far across all sections, filled now that
  // this section's pagination is final (the last section sets the true total).
  fillTotalPages(container, container.querySelectorAll('.ssd-page').length)
}

export function render(doc: DocxDocument, container: HTMLElement, options: RenderOptions = {}): void {
  showRevisions = options.showRevisions ?? false

  const hasPageBg = doc.blocks.some(
    b => b.type === 'paragraph' && extractPageBackground(b as ParagraphBlock) !== null,
  )

  const ps = doc.pageSize
  // Without a page size we can't lay out pages — fall back to continuous flow.
  if (!ps || ps.widthPx === 0 || ps.heightPx === 0) {
    // Template covers rely on tall empty spacer lines; plain documents must not.
    emptyLineEm = hasPageBg ? EMPTY_LINE_EM : LINE_HEIGHT
    renderBlocks(doc.blocks, container)
    renderNotes(doc, container)
    fillTotalPages(container, 1)
    return
  }

  // Multi-section documents: render each section with its own page size, routing
  // each to the plain or page-background path by its own content (see
  // renderSections). This covers both plain multi-section docs and template
  // documents that mix page sizes/orientations across sections.
  if (doc.sections && doc.sections.length > 1) {
    renderSections(doc, container, hasPageBg)
    return
  }

  // Template covers rely on tall empty spacer lines; plain documents must not.
  emptyLineEm = hasPageBg ? EMPTY_LINE_EM : LINE_HEIGHT

  // Plain documents: exact flow-based pagination into white page boxes.
  if (!hasPageBg) {
    renderPlainPaginated(doc, container, ps.widthPx, ps.heightPx, ps.marginPx)
    return
  }

  // Single-section template document with a full-page background.
  renderPageBgPaginated(doc, container)
}

// Render a multi-section document. Each section paginates with its own page size
// and orientation; page boxes concatenate and the page count continues across
// sections.
function renderSections(doc: DocxDocument, container: HTMLElement, docHasPageBg: boolean): void {
  const sections = doc.sections!
  const lastIdx = sections.length - 1

  // No section uses a full-page background: the plain multi-section path. Each
  // section renders the footnotes it references; endnotes go on the last section.
  if (!docHasPageBg) {
    emptyLineEm = LINE_HEIGHT
    sections.forEach((section, i) => {
      const sp = section.pageSize
      const subDoc: DocxDocument = {
        blocks: section.blocks,
        pageSize: sp,
        footnotes: doc.footnotes,
        endnotes: i === lastIdx ? doc.endnotes : undefined,
        footer: section.footer ?? doc.footer,
        header: section.header ?? doc.header,
      }
      const pageOffset = container.querySelectorAll('.ssd-page').length
      renderPlainPaginated(subDoc, container, sp.widthPx, sp.heightPx, sp.marginPx, pageOffset)
    })
    return
  }

  // At least one section has a full-page background. Route each section by ITS
  // OWN content (a plain section between template sections still paginates as a
  // plain flow), then render all notes once on a final notes page — template
  // documents don't use per-page footnotes.
  for (const section of sections) {
    const sp = section.pageSize
    const sectionHasBg = section.blocks.some(
      b => b.type === 'paragraph' && extractPageBackground(b as ParagraphBlock) !== null,
    )
    emptyLineEm = sectionHasBg ? EMPTY_LINE_EM : LINE_HEIGHT
    const subDoc: DocxDocument = {
      blocks: section.blocks,
      pageSize: sp,
      footer: section.footer ?? doc.footer,
      header: section.header ?? doc.header,
    }
    if (sectionHasBg) {
      renderPageBgPaginated(subDoc, container)
    } else {
      const pageOffset = container.querySelectorAll('.ssd-page').length
      renderPlainPaginated(subDoc, container, sp.widthPx, sp.heightPx, sp.marginPx, pageOffset)
    }
  }

  appendNotesPage(doc, container)
  fillTotalPages(container, container.querySelectorAll('.ssd-page').length)
}

// Append a final page box holding all footnotes and endnotes. Used by the
// combined multi-section + template path, where per-page footnotes don't apply.
function appendNotesPage(doc: DocxDocument, container: HTMLElement): void {
  const host = document.createElement('div')
  renderNotes(doc, host)
  if (host.childNodes.length === 0) return
  const ps = doc.pageSize!
  const pm = ps.marginPx
  const div = document.createElement('div')
  div.className = 'ssd-page'
  div.style.cssText = [
    'position:relative', 'box-sizing:border-box', 'background-color:#fff',
    'margin:0 auto 16px', 'box-shadow:0 2px 12px rgba(0,0,0,.25)',
    `width:${ps.widthPx}px`, `min-height:${ps.heightPx}px`,
    `padding:${pm.top}px ${pm.right}px ${pm.bottom}px ${pm.left}px`,
  ].join(';')
  while (host.firstChild) div.appendChild(host.firstChild)
  container.appendChild(div)
}

// Render a single-section template document with full-page background images
// (covers, framed body pages, closing slides). A two-pass DOM measurement plus
// heuristics reconstruct the pages; see src/renderer/layout.ts.
function renderPageBgPaginated(doc: DocxDocument, container: HTMLElement): void {
  const hasPageBg = true
  const ps = doc.pageSize!
  const pw = ps.widthPx
  const ph = ps.heightPx
  const pm = ps.marginPx
  const contentW = pw - pm.left - pm.right
  const contentH = ph - pm.top - pm.bottom

  // A page-level image (background/watermark) only flows out of the text for
  // template documents (hasPageBg). Plain documents keep all their content.
  const toFlow = (block: Block): Block | null =>
    hasPageBg && block.type === 'paragraph' ? flowOnly(block as ParagraphBlock, pw, ph) : block

  // ── Pass 1: measure each block's flow height in a hidden container ──────────
  // Measure in the SAME visual context the pages render in: append to the host
  // container with the .ssd-page class so it inherits the host's font and
  // line-height. Measuring at document.body level (different font/line-height)
  // mis-estimates heights and overflows the page boxes.
  const measureDiv = document.createElement('div')
  measureDiv.className = 'ssd-page'
  measureDiv.style.cssText =
    `position:fixed;top:-9999px;left:-9999px;visibility:hidden;pointer-events:none;` +
    `width:${contentW}px;padding:0;margin:0;min-height:0;box-shadow:none;background:none`
  container.appendChild(measureDiv)

  // Measure every block in ONE pass, all of them left in the DOM together, and
  // take each block's footprint as the distance to the next measured block's
  // top. Measuring a block alone and reading offsetHeight drops its margins:
  // they collapse straight through a bare wrapper div, so the sum came out
  // short and the page overflowed its own box (12px on a real template, and
  // silently, because the box is sized with min-height). Top-to-top spacing is
  // what the layout engine actually produced, collapsing included, counted
  // once. Same rule renderPlainPaginated uses for its stage.
  const blockHeight: number[] = new Array(doc.blocks.length).fill(0)
  // Only blocks with renderable flow get an element, so keep the pairs rather
  // than a sparse array indexed by block: the footprint of a measured block runs
  // to the next MEASURED one, skipping whatever produced nothing.
  const measured: Array<{ block: number; el: HTMLElement }> = []
  for (let i = 0; i < doc.blocks.length; i++) {
    const toMeasure = toFlow(doc.blocks[i])
    if (!toMeasure) continue
    const el = document.createElement('div')
    renderBlocks([toMeasure], el)
    measureDiv.appendChild(el)
    measured.push({ block: i, el })
  }
  for (let k = 0; k < measured.length; k++) {
    const nextTop = k + 1 < measured.length ? measured[k + 1].el.offsetTop : measureDiv.scrollHeight
    blockHeight[measured[k].block] = nextTop - measured[k].el.offsetTop
  }

  container.removeChild(measureDiv)

  // Forced page breaks: w:pageBreakBefore always; section headings only for
  // template documents (the heading heuristic is template-shaped — a plain doc
  // shouldn't start a new page at every large-text heading).
  const n = doc.blocks.length
  const sectionStart: boolean[] = new Array(n).fill(false)
  if (hasPageBg) {
    for (let i = 0; i < n; i++) {
      let heading = isHeadingBlock(doc.blocks[i])
      if (heading && i > 0 && sectionStart[i - 1] && isIconOnly(doc.blocks[i - 1])) {
        heading = false
      }
      const iconLead = isIconOnly(doc.blocks[i]) && i + 1 < n && isHeadingBlock(doc.blocks[i + 1])
      sectionStart[i] = heading || iconLead
    }
  }
  const forcesBreak = (i: number): boolean =>
    sectionStart[i] || (doc.blocks[i].type === 'paragraph' && !!(doc.blocks[i] as ParagraphBlock).pageBreakBefore)

  // Pagination: a new page starts at a forced break, or when a block does not
  // fit in the remaining space on the current page. A block taller than a full
  // page gets its own page (it will overflow, unavoidable without splitting it).
  const blockPage: number[] = new Array(n).fill(0)
  let page = 0
  let usedOnPage = 0
  for (let i = 0; i < n; i++) {
    const h = blockHeight[i]
    if ((forcesBreak(i) && usedOnPage > 0) || (usedOnPage > 0 && usedOnPage + h > contentH)) {
      page++
      usedOnPage = 0
    }
    blockPage[i] = page
    usedOnPage += h
  }

  let totalPages = page + 1

  // ── Determine which page each background region covers ─────────────────────
  // Floating background anchors (behindDoc=1) use absolute page positioning, so
  // their position in the XML flow does NOT mark where the background visually
  // starts. What's reliable is the ORDER of distinct background images: the
  // first is the cover/title page, each subsequent distinct image takes over the
  // body from the next page on. So distinct background k starts at page k.
  const distinctBgs: ImageRun[] = []
  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    if (block.type !== 'paragraph') continue
    const bg = extractPageBackground(block as ParagraphBlock)
    if (!bg) continue
    const last = distinctBgs[distinctBgs.length - 1]
    if (!last || last.src !== bg.src) distinctBgs.push(bg)
  }

  // pageBg[p]: distinct bg k covers page k; the last one fills all trailing pages.
  // keepIfEmpty[p]: this page is kept even with no visible flow. ONLY the cover
  // (the first distinct background) qualifies: a cover whose only text is an
  // emptied template variable still has a real, intentional page. Later
  // background regions (body frames, closing frames) that START on an otherwise
  // empty page are runoff/overlay, not real pages — keeping them produced phantom
  // pages with just a frame and no text between the cover and the first content.
  const pageBg: Array<ImageRun | null> = new Array(totalPages).fill(null)
  const keepIfEmpty: boolean[] = new Array(totalPages).fill(false)
  for (let k = 0; k < distinctBgs.length; k++) {
    const startPage = Math.min(k, totalPages - 1)
    const end = k + 1 < distinctBgs.length ? Math.min(k + 1, totalPages) : totalPages
    if (k === 0) keepIfEmpty[startPage] = true
    for (let p = startPage; p < end; p++) pageBg[p] = distinctBgs[k]
  }

  // ── Group flow content (and collect watermark overlays) per page ───────────
  const pageBlocks: Block[][] = Array.from({ length: totalPages }, () => [])
  const pageWatermarks: ImageRun[][] = Array.from({ length: totalPages }, () => [])
  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    const renderable = toFlow(block)
    if (renderable) pageBlocks[blockPage[i]].push(renderable)
    if (hasPageBg) for (const wm of watermarksOf(block, pw, ph)) pageWatermarks[blockPage[i]].push(wm)
  }

  // A frame watermark belongs to its whole background region, not just the page
  // its anchor landed on. Map each background to its watermark and apply it to
  // every page with that background, so a foreground frame consistently covers
  // the background (e.g. a "clean" frame hiding a logo baked into the bg image).
  const wmForBg = new Map<string, ImageRun>()
  for (let p = 0; p < totalPages; p++) {
    const bgSrc = pageBg[p]?.src ?? ''
    for (const wm of pageWatermarks[p]) if (!wmForBg.has(bgSrc)) wmForBg.set(bgSrc, wm)
  }
  for (let p = 0; p < totalPages; p++) {
    const wm = wmForBg.get(pageBg[p]?.src ?? '')
    pageWatermarks[p] = wm ? [wm] : []
    // When a foreground frame (watermark) covers a page, it is the visible
    // frame; drop the behindDoc background (white shows through) so anything
    // baked into that bg — e.g. a logo — doesn't appear behind the clean frame.
    if (wm) pageBg[p] = null
  }

  // Footnotes/endnotes get their own final page so they appear at the end of the
  // document inside a page box like everything else.
  const notePage = document.createElement('div')
  renderNotes(doc, notePage)
  if (notePage.childNodes.length > 0) {
    pageBlocks.push([])
    pageWatermarks.push([])
    pageBg.push(null)
    totalPages++
  }

  // Shared page-box styling (white sheet on the host background).
  const pageBoxCss = (extra: string[] = []): string => [
    'position:relative',
    'box-sizing:border-box',
    'background-color:#fff',
    'margin:0 auto 16px',
    'box-shadow:0 2px 12px rgba(0,0,0,.25)',
    `width:${pw}px`,
    `min-height:${ph}px`,
    ...extra,
  ].join(';')

  // ── Pass 2: create page divs and render content ────────────────────────────
  const bgH = `${ph}px`
  for (let p = 0; p < totalPages; p++) {
    const blocks = pageBlocks[p]
    const isNotePage = p === totalPages - 1 && notePage.childNodes.length > 0 && blocks.length === 0

    // Skip blank pages (only empty paragraphs) — trailing/standalone whitespace
    // would otherwise produce an empty page. The cover is kept even with no
    // visible flow (its only text may be an emptied template variable) so its
    // background still renders; other empty pages — including a body/closing
    // frame region that starts on an empty page — are dropped.
    const keepEmptyCover = keepIfEmpty[p] && (pageBg[p] !== null || pageWatermarks[p].length > 0)
    if (!isNotePage && !blocks.some(isBlockVisible) && !keepEmptyCover) continue

    const div = document.createElement('div')
    div.className = 'ssd-page'

    // The notes page: render the prepared notes section into a white page box.
    if (isNotePage) {
      div.style.cssText = pageBoxCss([`padding:${pm.top}px ${pm.right}px ${pm.bottom}px ${pm.left}px`])
      while (notePage.firstChild) div.appendChild(notePage.firstChild)
      container.appendChild(div)
      continue
    }

    // A page that is just a full-bleed image (e.g. the closing slide) is drawn
    // edge-to-edge with no margins and no underlying frame (template docs only).
    const fullImg = hasPageBg ? fullPageImage(blocks, pw) : null
    if (fullImg) {
      div.style.cssText = [
        'position:relative', 'box-sizing:border-box',
        `width:${pw}px`, `height:${ph}px`, `min-height:${ph}px`,
        'flex-shrink:0',
        'margin:0 auto 16px', 'overflow:hidden',
        'box-shadow:0 2px 12px rgba(0,0,0,.25)',
      ].join(';')
      const img = document.createElement('img')
      setImageSrc(img, fullImg.src)
      // Absolute inset:0 fills the page box regardless of host img rules (a
      // host `img { height: auto }` would otherwise collapse a height:100% img
      // to 0 and the page would vanish).
      img.style.cssText = 'position:absolute;top:0;left:0;width:100%;height:100%;object-fit:cover'
      div.appendChild(img)
      container.appendChild(div)
      continue
    }

    const bgImg = pageBg[p]
    const extra = [
      'background-repeat:no-repeat',
      'background-position:top center',
      `background-size:100% ${bgH}`,
      `padding:${pm.top}px ${pm.right}px ${pm.bottom}px ${pm.left}px`,
    ]
    // Sanitize the background image src before interpolating it into a CSS url()
    // so a malicious src can't break out of the url('…') and inject CSS.
    const bgSafe = bgImg ? sanitizeImageSrc(bgImg.src) : ''
    if (bgSafe) extra.push(`background-image:url('${bgSafe}')`)
    div.style.cssText = pageBoxCss(extra)

    // Watermark overlays: absolutely positioned behind the text layer. object-fit
    // fill makes the frame cover the page exactly (and hide anything in the bg).
    for (const wm of pageWatermarks[p]) {
      const img = document.createElement('img')
      setImageSrc(img, wm.src)
      img.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;object-fit:fill;z-index:0;pointer-events:none'
      div.appendChild(img)
    }

    // Text content sits above any watermark.
    const content = document.createElement('div')
    content.style.cssText = 'position:relative;z-index:1'
    renderBlocks(blocks, content)
    div.appendChild(content)
    container.appendChild(div)
  }

  // NUMPAGES = total rendered pages, now that all page boxes exist.
  fillTotalPages(container, container.querySelectorAll('.ssd-page').length)
}
