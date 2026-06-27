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

// Detect whether a paragraph is purely a page background carrier
function extractPageBackground(block: ParagraphBlock): ImageRun | null {
  for (const run of block.runs) {
    if (run.type === 'image' && (run as ImageRun).isPageBackground) {
      return run as ImageRun
    }
  }
  return null
}

export function render(doc: DocxDocument, container: HTMLElement): void {
  const hasPageBg = doc.blocks.some(
    b => b.type === 'paragraph' && extractPageBackground(b as ParagraphBlock) !== null,
  )

  if (!hasPageBg) {
    renderBlocks(doc.blocks, container)
    return
  }

  // Page-aware rendering: each behindDoc=1 anchor starts a new section div.
  // Without explicit page breaks we can only approximate pagination — content
  // may overflow the page background vertically, which is intentional.
  const ps = doc.pageSize
  const pw = ps?.widthPx ?? 0
  const ph = ps?.heightPx ?? 0
  const pm = ps?.marginPx ?? { top: 96, right: 96, bottom: 64, left: 96 }

  let pageDiv: HTMLElement | null = null
  let currentBgSrc = ''
  let pendingBlocks: Block[] = []

  const flushPage = () => {
    if (pageDiv && pendingBlocks.length > 0) renderBlocks(pendingBlocks, pageDiv)
    pendingBlocks = []
  }

  const createPageDiv = (bg: ImageRun) => {
    flushPage()
    const div = document.createElement('div')
    div.className = 'ssd-page'
    // background-size uses fixed pixel height so the image renders at exactly
    // one page height even when content overflows below it.
    const bgH = ph > 0 ? `${ph}px` : 'auto'
    const styles = [
      'position:relative',
      'box-sizing:border-box',
      'background-repeat:no-repeat',
      'background-position:top center',
      `background-size:100% ${bgH}`,
      `background-image:url('${bg.src}')`,
      `padding:${pm.top}px ${pm.right}px ${pm.bottom}px ${pm.left}px`,
      'margin:0 auto 16px',
      'box-shadow:0 2px 12px rgba(0,0,0,.25)',
    ]
    if (pw > 0) styles.push(`width:${pw}px`, `min-height:${ph}px`)
    div.style.cssText = styles.join(';')
    container.appendChild(div)
    pageDiv = div
    currentBgSrc = bg.src
  }

  for (const block of doc.blocks) {
    const bg = block.type === 'paragraph' ? extractPageBackground(block as ParagraphBlock) : null

    if (bg) {
      // Collapse consecutive anchors with the same image when nothing has been
      // queued between them — avoids empty phantom page divs.
      const collapse = pageDiv !== null && bg.src === currentBgSrc && pendingBlocks.length === 0

      if (!collapse) createPageDiv(bg)

      // Render any non-background runs in this same paragraph
      const otherRuns = (block as ParagraphBlock).runs.filter(r => !(r as ImageRun).isPageBackground)
      if (otherRuns.some(r => r.type !== 'run' || (r as TextRun).text)) {
        if (pageDiv) {
          const p = document.createElement('p')
          const css = styleToCss((block as ParagraphBlock).style)
          if (css) p.style.cssText = css
          for (const run of otherRuns) renderRun(run, p)
          if (p.textContent || p.children.length) pageDiv.appendChild(p)
        }
      }
    } else {
      if (pageDiv) {
        pendingBlocks.push(block)
      } else {
        renderBlocks([block], container)
      }
    }
  }

  flushPage()
}
