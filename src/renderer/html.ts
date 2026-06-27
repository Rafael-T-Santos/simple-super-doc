import type { DocxDocument, Block, ParagraphBlock, TableBlock, TextRun, ImageRun, Run, ComputedStyle } from '../types.js'

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

function renderParagraph(block: ParagraphBlock): HTMLElement {
  const p = document.createElement('p')
  const css = styleToCss(block.style)
  if (css) p.style.cssText = css
  for (const run of block.runs) {
    renderRun(run, p)
  }
  return p
}

function renderTable(block: TableBlock, container: HTMLElement): void {
  const table = document.createElement('table')
  table.style.borderCollapse = 'collapse'
  for (const row of block.rows) {
    const tr = document.createElement('tr')
    for (const cell of row.cells) {
      const td = document.createElement('td')
      if (cell.colSpan > 1) td.colSpan = cell.colSpan
      if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan
      if (cell.backgroundColor) td.style.backgroundColor = `#${cell.backgroundColor}`
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

// Return the isPageBackground ImageRun from a paragraph, or null.
function extractPageBackground(block: ParagraphBlock): ImageRun | null {
  for (const run of block.runs) {
    if (run.type === 'image' && (run as ImageRun).isPageBackground) return run as ImageRun
  }
  return null
}

// Return a version of the block with background image runs stripped.
// Returns null when the block has no remaining renderable content.
function withoutBgRuns(block: ParagraphBlock): ParagraphBlock | null {
  const runs = block.runs.filter(r => !(r as ImageRun).isPageBackground)
  const hasContent = runs.some(r => r.type !== 'run' || (r as TextRun).text.length > 0)
  return hasContent ? { ...block, runs } : null
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

  // blockPage[i] = 0-indexed page the block starts on (based on measured heights)
  const blockPage: number[] = new Array(doc.blocks.length).fill(0)
  let cumH = 0

  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    blockPage[i] = Math.floor(cumH / contentH)

    const bg = block.type === 'paragraph' ? extractPageBackground(block as ParagraphBlock) : null
    if (bg) {
      // Measure any text in this bg-anchor paragraph (e.g. "Fidelização" text)
      const textOnly = withoutBgRuns(block as ParagraphBlock)
      if (textOnly) {
        const el = document.createElement('div')
        renderBlocks([textOnly], el)
        measureDiv.appendChild(el)
        cumH += el.offsetHeight
        measureDiv.removeChild(el)
      }
      continue
    }

    const el = document.createElement('div')
    renderBlocks([block], el)
    measureDiv.appendChild(el)
    cumH += el.offsetHeight
    measureDiv.removeChild(el)
  }

  document.body.removeChild(measureDiv)

  const totalPages = Math.max(1, Math.ceil(cumH / contentH))

  // ── Determine which page each background region covers ─────────────────────
  // Each behindDoc=1 anchor targets the page it falls on in the measured layout.
  // Consecutive anchors with the same image on the same page are collapsed.
  // A background region covers from its target page until the next region starts.
  type BgRegion = { startPage: number; bg: ImageRun }
  const bgRegions: BgRegion[] = []

  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    if (block.type !== 'paragraph') continue
    const bg = extractPageBackground(block as ParagraphBlock)
    if (!bg) continue

    const targetPage = Math.min(blockPage[i], totalPages - 1)
    const last = bgRegions[bgRegions.length - 1]
    if (!last || last.bg.src !== bg.src || last.startPage !== targetPage) {
      bgRegions.push({ startPage: targetPage, bg })
    }
  }

  // Build pageBg[p] = background image for page p (fill-forward from each region)
  const pageBg: Array<ImageRun | null> = new Array(totalPages).fill(null)
  for (let r = 0; r < bgRegions.length; r++) {
    const { startPage, bg } = bgRegions[r]
    const end = r + 1 < bgRegions.length ? bgRegions[r + 1].startPage : totalPages
    for (let p = startPage; p < end; p++) pageBg[p] = bg
  }
  // Fill any trailing pages with the last known background
  if (bgRegions.length > 0) {
    const lastBg = bgRegions[bgRegions.length - 1].bg
    for (let p = bgRegions[bgRegions.length - 1].startPage; p < totalPages; p++) {
      if (!pageBg[p]) pageBg[p] = lastBg
    }
  }

  // ── Group renderable content into pages ────────────────────────────────────
  // For bg-anchor blocks with text, strip the bg run and include the text.
  const pageBlocks: Block[][] = Array.from({ length: totalPages }, () => [])
  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    const bg = block.type === 'paragraph' ? extractPageBackground(block as ParagraphBlock) : null
    const renderable: Block | null = bg ? withoutBgRuns(block as ParagraphBlock) : block
    if (renderable) pageBlocks[blockPage[i]].push(renderable)
  }

  // ── Pass 2: create page divs and render content ────────────────────────────
  const bgH = `${ph}px`
  for (let p = 0; p < totalPages; p++) {
    const div = document.createElement('div')
    div.className = 'ssd-page'

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

    renderBlocks(pageBlocks[p], div)
    container.appendChild(div)
  }
}
