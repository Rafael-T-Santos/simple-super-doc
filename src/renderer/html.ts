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
  if (run.type === 'image') {
    const img = document.createElement('img')
    img.src = (run as ImageRun).src
    img.width = (run as ImageRun).widthPx
    img.height = (run as ImageRun).heightPx
    img.style.display = 'inline-block'
    img.style.maxWidth = '100%'
    parent.appendChild(img)
    return
  }

  const textRun = run as TextRun
  const css = styleToCss(textRun.style)

  if (css) {
    const span = document.createElement('span')
    span.style.cssText = css
    // SECURITY: use textContent, never innerHTML
    span.textContent = textRun.text
    parent.appendChild(span)
  } else {
    parent.appendChild(document.createTextNode(textRun.text))
  }
}

// LINE_HEIGHT and EMPTY_LINE_EM live in ./layout (see their docs there).

function ensureLineBox(el: HTMLElement): void {
  // Empty runs still create (empty) text nodes, so check for visible content
  // rather than child count. Images carry their own height.
  if (!el.textContent && !el.querySelector('img')) {
    el.style.minHeight = `${EMPTY_LINE_EM}em`
  }
}

function renderParagraph(block: ParagraphBlock): HTMLElement {
  const p = document.createElement('p')
  // Zero the browser's default 1em paragraph margins and pin line-height to
  // Word's 1.15 so measured heights match Word's layout (this document's
  // vertical rhythm comes from empty paragraphs, not from default margins).
  p.style.cssText = `margin:0;line-height:${LINE_HEIGHT}`
  const css = styleToCss(block.style)
  if (css) p.style.cssText += ';' + css
  for (const run of block.runs) {
    renderRun(run, p)
  }
  ensureLineBox(p)
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

function renderBlocks(blocks: Block[], container: HTMLElement): void {
  let currentList: HTMLElement | null = null
  let currentNumId: string | null = null
  let currentOrdered: boolean | null = null

  const closeList = (): void => {
    if (currentList) {
      container.appendChild(currentList)
      currentList = null
      currentNumId = null
      currentOrdered = null
    }
  }

  for (const block of blocks) {
    if (block.type === 'paragraph' && block.list) {
      const { numId, ordered, start } = block.list

      if (numId !== currentNumId || ordered !== currentOrdered) {
        closeList()
        currentList = document.createElement(ordered ? 'ol' : 'ul')
        if (ordered) (currentList as HTMLOListElement).start = start
        currentNumId = numId
        currentOrdered = ordered
      }

      const li = document.createElement('li')
      const css = styleToCss(block.style)
      if (css) li.style.cssText = css
      for (const run of block.runs) {
        renderRun(run, li)
      }
      currentList!.appendChild(li)
    } else {
      closeList()

      if (block.type === 'paragraph') {
        container.appendChild(renderParagraph(block))
      } else if (block.type === 'table') {
        renderTable(block, container)
      }
    }
  }

  closeList()
}

export function render(doc: DocxDocument, container: HTMLElement): void {
  const hasPageBg = doc.blocks.some(
    b => b.type === 'paragraph' && extractPageBackground(b as ParagraphBlock) !== null,
  )

  if (!hasPageBg) {
    renderBlocks(doc.blocks, container)
    return
  }

  const ps = doc.pageSize
  if (!ps || ps.widthPx === 0 || ps.heightPx === 0) {
    renderBlocks(doc.blocks, container)
    return
  }

  const pw = ps.widthPx
  const ph = ps.heightPx
  const pm = ps.marginPx
  const contentW = pw - pm.left - pm.right
  const contentH = ph - pm.top - pm.bottom

  // ── Pass 1: measure each block's rendered height in a hidden container ──────
  // Background anchor blocks (isPageBackground) contribute no visible height but
  // may contain text — that text is measured separately via withoutBgRuns().
  const measureDiv = document.createElement('div')
  measureDiv.style.cssText =
    `position:fixed;top:-9999px;left:-9999px;visibility:hidden;pointer-events:none;width:${contentW}px;`
  document.body.appendChild(measureDiv)

  // Measure each block's flow height. Page-level images (backgrounds and
  // watermarks) are absolutely positioned and contribute no flow height.
  const blockHeight: number[] = new Array(doc.blocks.length).fill(0)
  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    const toMeasure: Block | null =
      block.type === 'paragraph' ? flowOnly(block as ParagraphBlock, pw, ph) : block
    if (!toMeasure) continue
    const el = document.createElement('div')
    renderBlocks([toMeasure], el)
    measureDiv.appendChild(el)
    blockHeight[i] = el.offsetHeight
    measureDiv.removeChild(el)
  }

  document.body.removeChild(measureDiv)

  // Section starts: each section heading begins a new page. When a small icon
  // immediately precedes a text heading, the break goes before the icon so the
  // two stay together.
  const n = doc.blocks.length
  const sectionStart: boolean[] = new Array(n).fill(false)
  for (let i = 0; i < n; i++) {
    let heading = isHeadingBlock(doc.blocks[i])
    // Avoid a double break when an icon already opened this section.
    if (heading && i > 0 && sectionStart[i - 1] && isIconOnly(doc.blocks[i - 1])) {
      heading = false
    }
    const iconLead = isIconOnly(doc.blocks[i]) && i + 1 < n && isHeadingBlock(doc.blocks[i + 1])
    sectionStart[i] = heading || iconLead
  }

  // Pagination: a new page starts at a section heading, or when a block does not
  // fit in the remaining space on the current page. A block taller than a full
  // page gets its own page (it will overflow, unavoidable without splitting it).
  const blockPage: number[] = new Array(n).fill(0)
  let page = 0
  let usedOnPage = 0
  for (let i = 0; i < n; i++) {
    const h = blockHeight[i]
    const forceBreak = sectionStart[i] && usedOnPage > 0
    if (forceBreak || (usedOnPage > 0 && usedOnPage + h > contentH)) {
      page++
      usedOnPage = 0
    }
    blockPage[i] = page
    usedOnPage += h
  }

  const totalPages = page + 1

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
    const renderable: Block | null =
      block.type === 'paragraph' ? flowOnly(block as ParagraphBlock, pw, ph) : block
    if (renderable) pageBlocks[blockPage[i]].push(renderable)
    for (const wm of watermarksOf(block, pw, ph)) pageWatermarks[blockPage[i]].push(wm)
  }

  // ── Pass 2: create page divs and render content ────────────────────────────
  const bgH = `${ph}px`
  for (let p = 0; p < totalPages; p++) {
    const blocks = pageBlocks[p]

    // Skip blank pages (only empty paragraphs) — trailing/standalone whitespace
    // would otherwise produce an empty framed page.
    if (!blocks.some(isBlockVisible)) continue

    const div = document.createElement('div')
    div.className = 'ssd-page'

    // A page that is just a full-bleed image (e.g. the closing slide) is drawn
    // edge-to-edge with no margins and no underlying frame.
    const fullImg = fullPageImage(blocks, pw)
    if (fullImg) {
      div.style.cssText = [
        'position:relative',
        'box-sizing:border-box',
        `width:${pw}px`,
        `height:${ph}px`,
        'margin:0 auto 16px',
        'overflow:hidden',
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
    const styles = [
      'position:relative',
      'box-sizing:border-box',
      'background-repeat:no-repeat',
      'background-position:top center',
      `background-size:100% ${bgH}`,
      `padding:${pm.top}px ${pm.right}px ${pm.bottom}px ${pm.left}px`,
      'margin:0 auto 16px',
      'box-shadow:0 2px 12px rgba(0,0,0,.25)',
      `width:${pw}px`,
      `min-height:${ph}px`,
    ]
    if (bgImg) styles.push(`background-image:url('${bgImg.src}')`)
    div.style.cssText = styles.join(';')

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
