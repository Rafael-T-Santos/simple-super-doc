import type { DocxDocument, Block, ParagraphBlock, TableBlock, TextRun, ImageRun, Run, ComputedStyle, NoteEntry, TabStop } from '../types.js'
import {
  EMPTY_LINE_EM, LINE_HEIGHT,
  extractPageBackground, isBlockVisible, isHeadingBlock, isIconOnly,
  fullPageImage, flowOnly, watermarksOf,
} from './layout.js'

function styleToCss(s: ComputedStyle): string {
  const parts: string[] = []
  if (s.bold) parts.push('font-weight:bold')
  if (s.italic) parts.push('font-style:italic')
  // underline and strikethrough combine into one text-decoration.
  const deco = [s.underline ? 'underline' : '', s.strike ? 'line-through' : ''].filter(Boolean)
  if (deco.length) parts.push(`text-decoration:${deco.join(' ')}`)
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
    const img = document.createElement('img')
    img.src = (run as ImageRun).src
    img.width = (run as ImageRun).widthPx
    img.height = (run as ImageRun).heightPx
    img.style.display = 'inline-block'
    img.style.maxWidth = '100%'
    target.appendChild(img)
    return
  }

  const textRun = run as TextRun

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
    if (css) {
      const span = document.createElement('span')
      span.style.cssText = css
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

// Empty-paragraph line height, set per render(): template covers (hasPageBg)
// need a taller empty line to match Word's layout (EMPTY_LINE_EM); plain
// documents use a normal single line (LINE_HEIGHT) so blank lines don't bloat
// the page. LINE_HEIGHT and EMPTY_LINE_EM live in ./layout.
let emptyLineEm = LINE_HEIGHT

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
function renderFooter(doc: DocxDocument, pageDiv: HTMLElement, pageNum: number, pm: PageMargins, footerPx: number): void {
  if (!doc.footer || doc.footer.length === 0) return
  const prev = currentPageNumber
  currentPageNumber = pageNum
  const el = document.createElement('div')
  el.className = 'ssd-footer'
  el.style.cssText = `position:absolute;left:${pm.left}px;right:${pm.right}px;bottom:${footerPx}px`
  renderBlocks(doc.footer, el)
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
  if (!doc.header || doc.header.length === 0) return
  const prev = currentPageNumber
  currentPageNumber = pageNum
  const el = document.createElement('div')
  el.className = 'ssd-header'
  el.style.cssText = `position:absolute;left:${pm.left}px;right:${pm.right}px;top:${headerPx}px`
  renderBlocks(doc.header, el)
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
  const lh = style.lineHeightPx != null ? `${style.lineHeightPx}px` : `${style.lineHeight ?? LINE_HEIGHT}`
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

  // A right/center/decimal tab stop with leading tabs (a table-of-contents row:
  // "Title.....12") is laid out with a flex leader instead of a blank spacer.
  const stops = block.style.tabStops
  const hasTabRun = block.runs.some(r => r.type === 'run' && ((r as TextRun).tabs ?? 0) > 0)
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

  // Apply the document's column widths (w:tblGrid) with a fixed layout so column
  // proportions — and therefore text wrapping and row heights — match the
  // document. Without this the browser sizes columns by content, narrowing some
  // columns, wrapping text to more lines, and fitting fewer rows per page.
  // Falls back to content-based sizing when the document gives no grid.
  if (block.columnWidths && block.columnWidths.length > 0) {
    const total = block.columnWidths.reduce((a, b) => a + b, 0)
    if (total > 0) {
      table.style.tableLayout = 'fixed'
      table.style.width = '100%'
      const colgroup = document.createElement('colgroup')
      for (const w of block.columnWidths) {
        const col = document.createElement('col')
        col.style.width = `${((w / total) * 100).toFixed(3)}%`
        colgroup.appendChild(col)
      }
      table.appendChild(colgroup)
    }
  }

  const pad = block.cellPadding
  for (const row of block.rows) {
    const tr = document.createElement('tr')
    for (const cell of row.cells) {
      const td = document.createElement('td')
      if (cell.colSpan > 1) td.colSpan = cell.colSpan
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan
      if (cell.backgroundColor) td.style.backgroundColor = `#${cell.backgroundColor}`
      // Cell padding from the document's w:tcMar (keeps rows as compact as Word).
      if (pad) td.style.padding = `${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px`
      td.style.verticalAlign = 'top'
      // Let long template tokens wrap instead of forcing overflow.
      td.style.overflowWrap = 'break-word'
      td.style.wordBreak = 'break-word'
      // Cell padding already provides the spacing, so don't add the paragraph's
      // before/after margins inside cells (Word keeps cell content tight).
      const prev = inTableCell
      inTableCell = true
      renderBlocks(cell.blocks, td)
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
      const { numId, ordered, start, ilvl, format } = block.list

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
        const styleType = LIST_STYLE[format]
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
  const rest = table.cloneNode(false) as HTMLTableElement
  for (let i = splitAt; i < rows.length; i++) rest.appendChild(rows[i])
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

  pageDiv = newPage()
  for (let i = 0; i < children.length; i++) {
    const child = children[i]
    const forced = child.dataset.ssdBreak === '1'

    // Tables can split across pages: rows that fit stay, the rest continue on
    // the next page (and split again if still too tall).
    if (child.tagName === 'TABLE') {
      if (forced && used > 0) pageDiv = newPage()
      let table = child as HTMLTableElement
      for (;;) {
        pageDiv.appendChild(table)
        const th = table.offsetHeight
        if (used + th <= contentH - reserve) { used += th; break } // fits
        const rest = splitTableRows(table, contentH - reserve - used)
        if (!rest) {
          if (used === 0) { used += th; break } // taller than a full page; accept overflow
          pageDiv.removeChild(table)
          pageDiv = newPage()
          continue // retry on a fresh page
        }
        // `table` now holds the rows that fit; `rest` continues on a new page.
        pageDiv = newPage()
        table = rest
      }
      const refs = footnoteNumbersIn(table)
      recordFootnotes(refs, fnReserve(refs, pageHasFn))
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

  // Page footer / header (with page numbers) on every page.
  const footerPx = doc.pageSize?.footerPx ?? Math.round(pm.bottom / 2)
  pages.forEach((page, i) => renderFooter(doc, page, i + 1, pm, footerPx))
  const headerPx = doc.pageSize?.headerPx ?? Math.round(pm.top / 2)
  pages.forEach((page, i) => renderHeader(doc, page, i + 1, pm, headerPx))

  // NUMPAGES = total rendered pages (body + any endnote page), filled now that
  // pagination is final.
  fillTotalPages(container, container.querySelectorAll('.ssd-page').length)
}

export function render(doc: DocxDocument, container: HTMLElement): void {
  const hasPageBg = doc.blocks.some(
    b => b.type === 'paragraph' && extractPageBackground(b as ParagraphBlock) !== null,
  )

  // Template covers rely on tall empty spacer lines; plain documents must not.
  emptyLineEm = hasPageBg ? EMPTY_LINE_EM : LINE_HEIGHT

  const ps = doc.pageSize
  // Without a page size we can't lay out pages — fall back to continuous flow.
  if (!ps || ps.widthPx === 0 || ps.heightPx === 0) {
    renderBlocks(doc.blocks, container)
    renderNotes(doc, container)
    fillTotalPages(container, 1)
    return
  }

  // Plain documents: exact flow-based pagination into white page boxes.
  if (!hasPageBg) {
    renderPlainPaginated(doc, container, ps.widthPx, ps.heightPx, ps.marginPx)
    return
  }

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

  const blockHeight: number[] = new Array(doc.blocks.length).fill(0)
  for (let i = 0; i < doc.blocks.length; i++) {
    const toMeasure = toFlow(doc.blocks[i])
    if (!toMeasure) continue
    const el = document.createElement('div')
    renderBlocks([toMeasure], el)
    measureDiv.appendChild(el)
    blockHeight[i] = el.offsetHeight
    measureDiv.removeChild(el)
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
  const pageBg: Array<ImageRun | null> = new Array(totalPages).fill(null)
  for (let k = 0; k < distinctBgs.length; k++) {
    const startPage = Math.min(k, totalPages - 1)
    const end = k + 1 < distinctBgs.length ? Math.min(k + 1, totalPages) : totalPages
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
    // would otherwise produce an empty page.
    if (!isNotePage && !blocks.some(isBlockVisible)) continue

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
      img.src = fullImg.src
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
    if (bgImg) extra.push(`background-image:url('${bgImg.src}')`)
    div.style.cssText = pageBoxCss(extra)

    // Watermark overlays: absolutely positioned behind the text layer. object-fit
    // fill makes the frame cover the page exactly (and hide anything in the bg).
    for (const wm of pageWatermarks[p]) {
      const img = document.createElement('img')
      img.src = wm.src
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
