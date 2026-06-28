import type { DocxDocument, Block, ParagraphBlock, TableBlock, TextRun, ImageRun, Run, ComputedStyle } from '../types.js'
import {
  EMPTY_LINE_EM, LINE_HEIGHT,
  extractPageBackground, isBlockVisible, isHeadingBlock, isIconOnly,
  fullPageImage, flowOnly, watermarksOf,
} from './layout.js'

function styleToCss(s: ComputedStyle): string {
  const parts: string[] = []
  if (s.bold) parts.push('font-weight:bold')
  if (s.italic) parts.push('font-style:italic')
  if (s.underline) parts.push('text-decoration:underline')
  if (s.fontSize != null) parts.push(`font-size:${s.fontSize}pt`)
  if (s.fontFamily) parts.push(`font-family:${s.fontFamily},sans-serif`)
  if (s.color) parts.push(`color:#${s.color}`)
  if (s.alignment) parts.push(`text-align:${s.alignment}`)
  if (s.backgroundColor) parts.push(`background-color:#${s.backgroundColor}`)
  return parts.join(';')
}

function renderRun(run: Run, parent: HTMLElement): void {
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

  const css = styleToCss(textRun.style)

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

// Render the footnote and endnote sections at the end of the document. Each note
// is an <li> (numbered to match its in-text marker) with a back-reference link.
function renderNotes(doc: DocxDocument, container: HTMLElement): void {
  const sections: Array<{ kind: 'footnote' | 'endnote'; prefix: string; label: string; notes: DocxDocument['footnotes'] }> = [
    { kind: 'footnote', prefix: 'fn', label: 'Footnotes', notes: doc.footnotes },
    { kind: 'endnote', prefix: 'en', label: 'Endnotes', notes: doc.endnotes },
  ]
  for (const { prefix, label, notes } of sections) {
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

// CSS for a paragraph / list item: the document's w:spacing before/after as
// margins, its line spacing, and the run-level style. Replaces the browser's
// default 1em margins so the vertical rhythm matches the document.
function paragraphCss(style: ComputedStyle): string {
  const mt = style.spaceBefore ?? 0
  const mb = style.spaceAfter ?? 0
  const lh = style.lineHeightPx != null ? `${style.lineHeightPx}px` : `${style.lineHeight ?? LINE_HEIGHT}`
  let css = `margin:${mt}px 0 ${mb}px;line-height:${lh}`
  const inline = styleToCss(style)
  if (inline) css += ';' + inline
  return css
}

function renderParagraph(block: ParagraphBlock): HTMLElement {
  const p = document.createElement('p')
  p.style.cssText = paragraphCss(block.style)
  for (const run of block.runs) {
    renderRun(run, p)
  }
  ensureLineBox(p)
  if (block.pageBreakBefore) p.dataset.ssdBreak = '1'
  return p
}

function renderTable(block: TableBlock, container: HTMLElement): void {
  const table = document.createElement('table')
  table.style.borderCollapse = 'collapse'
  // Cap tables at the container width so a wide table can't spill past the page
  // frame, but DON'T force width:100% — that would stretch a small table to full
  // width. Combined with per-cell word-break (below), an over-wide table shrinks
  // by wrapping its content instead of overflowing.
  table.style.maxWidth = '100%'
  for (const row of block.rows) {
    const tr = document.createElement('tr')
    for (const cell of row.cells) {
      const td = document.createElement('td')
      if (cell.colSpan > 1) td.colSpan = cell.colSpan
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan
      if (cell.backgroundColor) td.style.backgroundColor = `#${cell.backgroundColor}`
      // Let long template tokens wrap instead of forcing overflow.
      td.style.overflowWrap = 'break-word'
      td.style.wordBreak = 'break-word'
      renderBlocks(cell.blocks, td)
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
        const parent = stack[stack.length - 1]
        if (parent) (parent.lastLi ?? parent.el).appendChild(listEl)
        stack.push({ el: listEl, numId, lastLi: null })
      }

      const frame = stack[stack.length - 1]
      const li = document.createElement('li')
      li.style.cssText = paragraphCss(block.style)
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

type PageMargins = { top: number; right: number; bottom: number; left: number }

// Page box styling shared by paginated renders (white sheet on the host bg).
function pageBoxStyle(pw: number, ph: number, pm: PageMargins, extra: string[] = []): string {
  return [
    'position:relative',
    'box-sizing:border-box',
    'background-color:#fff',
    'margin:0 auto 16px',
    'box-shadow:0 2px 12px rgba(0,0,0,.25)',
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

  const newPage = (): HTMLElement => {
    const div = document.createElement('div')
    div.className = 'ssd-page'
    div.style.cssText = pageBoxStyle(pw, ph, pm)
    container.appendChild(div)
    return div
  }

  let pageDiv = newPage()
  let used = 0
  for (let i = 0; i < children.length; i++) {
    const h = heights[i]
    const forced = children[i].dataset.ssdBreak === '1'
    if ((forced && used > 0) || (used > 0 && used + h > contentH)) {
      pageDiv = newPage()
      used = 0
    }
    pageDiv.appendChild(children[i]) // moves the node out of the stage
    used += h
  }
  stage.remove()

  // Footnotes/endnotes on their own final page.
  const noteHost = document.createElement('div')
  renderNotes(doc, noteHost)
  if (noteHost.childNodes.length > 0) {
    const div = newPage()
    while (noteHost.firstChild) div.appendChild(noteHost.firstChild)
  }
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
        `width:${pw}px`, `height:${ph}px`,
        'margin:0 auto 16px', 'overflow:hidden',
        'box-shadow:0 2px 12px rgba(0,0,0,.25)',
      ].join(';')
      const img = document.createElement('img')
      img.src = fullImg.src
      img.style.cssText = 'display:block;width:100%;height:100%;object-fit:cover'
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

    // Watermark overlays: absolutely positioned behind the text layer.
    for (const wm of pageWatermarks[p]) {
      const img = document.createElement('img')
      img.src = wm.src
      img.style.cssText =
        'position:absolute;top:0;left:0;width:100%;height:100%;object-fit:contain;z-index:0;pointer-events:none'
      div.appendChild(img)
    }

    // Text content sits above any watermark.
    const content = document.createElement('div')
    content.style.cssText = 'position:relative;z-index:1'
    renderBlocks(blocks, content)
    div.appendChild(content)
    container.appendChild(div)
  }
}
