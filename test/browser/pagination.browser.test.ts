import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import JSZip from 'jszip'
import { build } from 'esbuild'
import { chromium, type Browser } from 'playwright-core'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

// Browser-only regression tests. The two-pass, page-aware pagination in
// src/renderer needs a real layout engine (offsetHeight, margin collapsing),
// so it can't run in the node/jsdom test environment the unit suite uses. We
// bundle the library to an IIFE with esbuild, drive it in headless Chromium via
// Playwright, and assert on the rendered DOM. Run with `npm run test:browser`
// (needs `npx playwright install chromium` once).

const __dirname = dirname(fileURLToPath(import.meta.url))
const ENTRY = join(__dirname, '..', '..', 'src', 'index.ts')

// A minimal .docx whose only table is tall enough to span two pages and whose
// FIRST row references a footnote. This is the exact shape that regressed:
// footnotes referenced by rows that stay on an earlier page were dropped, and
// their reserved height was ignored when choosing the table's break point.
async function buildSplitTableWithFootnoteDocx(): Promise<string> {
  const rows: string[] = []
  for (let i = 0; i < 12; i++) {
    const noteRef =
      i === 0
        ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>`
        : ''
    rows.push(
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>Row ${i + 1} of a table that spans a page break</w:t></w:r>${noteRef}</w:p>` +
        `</w:tc></w:tr>`,
    )
  }

  const body =
    `<w:tbl>` +
    `<w:tblPr><w:tblW w:w="9000" w:type="dxa"/>` +
    `<w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:left w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:bottom w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:right w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideH w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideV w:val="single" w:sz="4" w:color="000000"/>` +
    `</w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
    rows.join('') +
    `</w:tbl>` +
    // Body-level section: a short page so the 12-row table must split.
    `<w:sectPr><w:pgSz w:w="12240" w:h="2600"/>` +
    `<w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/footnotes.xml',
    `<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
      `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
      `<w:footnote w:id="2"><w:p><w:r><w:t>The footnote that must appear on page 1.</w:t></w:r></w:p></w:footnote>` +
      `</w:footnotes>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

// A .docx with a bullet list whose numbering defines a hyphen glyph
// (w:lvlText "-"). The renderer must honor that literal marker instead of
// falling back to a generic disc "•".
async function buildHyphenBulletDocx(): Promise<string> {
  const body =
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
    `<w:r><w:t>First bullet</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
    `<w:r><w:t>Second bullet</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
  )
  zip.file(
    'word/numbering.xml',
    `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
      `<w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

let browser: Browser
let bundleJs: string

beforeAll(async () => {
  const out = await build({
    entryPoints: [ENTRY],
    bundle: true,
    format: 'iife',
    globalName: 'SimpleDoc',
    platform: 'browser',
    write: false,
  })
  bundleJs = out.outputFiles[0].text
  browser = await chromium.launch()
}, 60_000)

afterAll(async () => {
  await browser?.close()
})

async function renderInBrowser(b64: string) {
  const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
  await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
  await page.addScriptTag({ content: bundleJs })
  const result = await page.evaluate(async (b64: string) => {
    const bin = atob(b64)
    const arr = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const SD = (window as any).SimpleDoc
    const doc = await SD.parse(arr.buffer)
    SD.render(doc, document.getElementById('view'))
    const pages = Array.from(document.querySelectorAll('.ssd-page')) as HTMLElement[]
    const refPageIndex = pages.findIndex(p => p.querySelector('sup[id^="fnref-"]'))
    return {
      nPages: pages.length,
      nTables: document.querySelectorAll('table').length,
      refPageIndex,
      footnoteBoxOnRefPage:
        refPageIndex >= 0 ? pages[refPageIndex].querySelectorAll('.ssd-footnotes').length : 0,
      totalFootnoteBoxes: document.querySelectorAll('.ssd-footnotes').length,
      footnoteText: (document.querySelector('.ssd-footnotes')?.textContent ?? '').trim(),
    }
  }, b64)
  await page.close()
  return result
}

describe('page-aware pagination: footnotes on a split table', () => {
  it('renders a row-referenced footnote on the page that holds the row', async () => {
    const b64 = await buildSplitTableWithFootnoteDocx()
    const r = await renderInBrowser(b64)

    // The table must actually split, otherwise the regression can't occur.
    expect(r.nPages).toBeGreaterThanOrEqual(2)
    expect(r.nTables).toBeGreaterThanOrEqual(2)

    // The footnote reference is in row 1 (an early page). Its footnote must be
    // rendered exactly once, on the same page as the reference. Before the fix
    // it was dropped entirely (0 boxes) because only the final split piece was
    // scanned for footnote refs.
    expect(r.refPageIndex).toBeGreaterThanOrEqual(0)
    expect(r.footnoteBoxOnRefPage).toBe(1)
    expect(r.totalFootnoteBoxes).toBe(1)
    expect(r.footnoteText).toContain('must appear on page 1')
  }, 30_000)
})

describe('bullet lists: literal lvlText glyph', () => {
  it('renders a hyphen bullet as "-" (list-style-type), not a generic disc', async () => {
    const b64 = await buildHyphenBulletDocx()
    const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
    await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
    await page.addScriptTag({ content: bundleJs })
    const styleType = await page.evaluate(async (b64: string) => {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const SD = (window as any).SimpleDoc
      const doc = await SD.parse(arr.buffer)
      SD.render(doc, document.getElementById('view'))
      const ul = document.querySelector('ul') as HTMLElement | null
      return ul?.style.listStyleType ?? ''
    }, b64)
    await page.close()
    // CSS string marker for a hyphen; must not be the generic "disc".
    expect(styleType).toBe('"-"')
  }, 30_000)
})

// A .docx with a letterhead-style header (an anchored logo positioned toward the
// right) and a "left <tab> right" footer, to prove header/footer positioning:
// the footer's trailing segment right-aligns and the anchored logo lands at its
// offset instead of flowing inline at the left.
async function buildHeaderFooterLayoutDocx(): Promise<string> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`)
  zip.file('_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
    `<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>` +
    `</Relationships>`)
  zip.file('word/_rels/header1.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>` +
    `</Relationships>`)
  zip.file('word/media/logo.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47]))
  zip.file('word/header1.xml',
    `<?xml version="1.0"?><w:hdr ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<w:p><w:r><w:drawing><wp:anchor behindDoc="1">` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>5000000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="900000" cy="250000"/>` +
    `<a:graphic><a:graphicData><pic:pic><pic:blipFill>` +
    `<a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r></w:p></w:hdr>`)
  zip.file('word/footer1.xml',
    `<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:p><w:r><w:t xml:space="preserve">LEFTSIDE </w:t><w:tab/><w:tab/>` +
    `<w:t xml:space="preserve">RIGHTSIDE</w:t></w:r></w:p></w:ftr>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>` +
    `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>` +
    `<w:footerReference w:type="default" r:id="rId11"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>` +
    `</w:sectPr></w:body></w:document>`)
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

describe('header/footer positioning', () => {
  it('right-aligns a footer\'s trailing tab segment and positions a header logo', async () => {
    const b64 = await buildHeaderFooterLayoutDocx()
    const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
    await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
    await page.addScriptTag({ content: bundleJs })
    const r = await page.evaluate(async (b64: string) => {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const SD = (window as any).SimpleDoc
      const doc = await SD.parse(arr.buffer)
      SD.render(doc, document.getElementById('view'))
      const pageEl = document.querySelector('.ssd-page') as HTMLElement
      const footer = pageEl.querySelector('.ssd-footer') as HTMLElement
      const fr = footer.getBoundingClientRect()
      const findSpan = (t: string) =>
        [...footer.querySelectorAll('span')].find(s => (s.textContent || '').trim() === t) as HTMLElement
      const left = findSpan('LEFTSIDE').getBoundingClientRect()
      const right = findSpan('RIGHTSIDE').getBoundingClientRect()
      const header = pageEl.querySelector('.ssd-header') as HTMLElement
      const hr = header.getBoundingClientRect()
      const img = header.querySelector('img') as HTMLImageElement
      const ir = img.getBoundingClientRect()
      return {
        footerW: fr.width,
        leftStart: left.left - fr.left,
        rightEnd: fr.right - right.right,
        imgPos: getComputedStyle(img).position,
        imgLeftFrac: (ir.left - hr.left) / hr.width,
      }
    }, b64)
    await page.close()
    // "LEFTSIDE" hugs the left edge; "RIGHTSIDE" hugs the right edge.
    expect(r.leftStart).toBeLessThan(8)
    expect(r.rightEnd).toBeLessThan(8)
    // The logo is absolutely positioned in the right half, not inline at left.
    expect(r.imgPos).toBe('absolute')
    expect(r.imgLeftFrac).toBeGreaterThan(0.5)
  }, 30_000)
})

async function buildHiddenTextDocx(): Promise<string> {
  const body =
    `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>SECRET</w:t></w:r>` +
    `<w:r><w:t>VISIBLE</w:t></w:r></w:p>`
  const zip = new JSZip()
  zip.file('[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`)
  zip.file('_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults></w:styles>`)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`)
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

describe('hidden text (w:vanish)', () => {
  it('does not render hidden text but keeps the rest', async () => {
    const b64 = await buildHiddenTextDocx()
    const page = await browser.newPage()
    await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
    await page.addScriptTag({ content: bundleJs })
    const text = await page.evaluate(async (b64: string) => {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const SD = (window as any).SimpleDoc
      const doc = await SD.parse(arr.buffer)
      SD.render(doc, document.getElementById('view'))
      return document.getElementById('view')!.textContent ?? ''
    }, b64)
    await page.close()
    expect(text).not.toContain('SECRET')
    expect(text).toContain('VISIBLE')
  }, 30_000)
})

// A table whose w:tblGrid gives two very UNEQUAL columns, with enough rows to
// span three pages. Three, not two, because each continuation is split again
// from the previous continuation — so the column widths have to survive being
// carried forward more than once.
async function buildUnequalColumnSplitTableDocx(): Promise<string> {
  const rows: string[] = []
  for (let i = 0; i < 16; i++) {
    rows.push(
      `<w:tr>` +
        `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>${i + 1}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="7500" w:type="dxa"/></w:tcPr>` +
        `<w:p><w:r><w:t>Wide description cell for row ${i + 1}</w:t></w:r></w:p></w:tc>` +
        `</w:tr>`,
    )
  }

  const body =
    `<w:tbl>` +
    `<w:tblPr><w:tblW w:w="9000" w:type="dxa"/>` +
    `<w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:left w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:bottom w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:right w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideH w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideV w:val="single" w:sz="4" w:color="000000"/>` +
    `</w:tblBorders></w:tblPr>` +
    // 1500 / 7500 twips — a 1:5 ratio, impossible to confuse with equal columns.
    `<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="7500"/></w:tblGrid>` +
    rows.join('') +
    `</w:tbl>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="2600"/>` +
    `<w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

describe('page-aware pagination: column widths on a split table', () => {
  it('keeps the document column widths on every continued piece', async () => {
    const b64 = await buildUnequalColumnSplitTableDocx()
    const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
    await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
    await page.addScriptTag({ content: bundleJs })
    const pieces = await page.evaluate(async (b64: string) => {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const SD = (window as any).SimpleDoc
      const doc = await SD.parse(arr.buffer)
      SD.render(doc, document.getElementById('view'))
      // Only the real page boxes — the hidden staging div also carries .ssd-page.
      const pages = Array.from(document.querySelectorAll('.ssd-page')) as HTMLElement[]
      const out: { hasColgroup: boolean; widths: number[] }[] = []
      for (const pg of pages) {
        if (pg.style.visibility === 'hidden') continue
        for (const t of Array.from(pg.querySelectorAll('table'))) {
          const first = (t as HTMLTableElement).rows[0]
          if (!first) continue
          out.push({
            hasColgroup: !!t.querySelector('colgroup'),
            widths: Array.from(first.cells).map(c => Math.round(c.getBoundingClientRect().width)),
          })
        }
      }
      return out
    }, b64)
    await page.close()

    // The table must actually split more than once, or the regression can't show.
    expect(pieces.length).toBeGreaterThanOrEqual(3)

    const [firstPiece, ...continued] = pieces
    expect(firstPiece.widths).toHaveLength(2)
    // 1500:7500 is 1:5 — the narrow column must stay clearly narrower.
    expect(firstPiece.widths[0]).toBeLessThan(firstPiece.widths[1] / 2)

    // Every continuation must match the first piece. Before the fix, the shallow
    // cloneNode(false) dropped <colgroup> while keeping table-layout:fixed, so a
    // continued piece split its width equally (e.g. [339,339] instead of
    // [124,556]) and the columns visibly changed shape mid-table.
    for (const piece of continued) {
      expect(piece.hasColgroup).toBe(true)
      expect(piece.widths).toEqual(firstPiece.widths)
    }
  }, 30_000)
})

// w:tblHeader marks a heading row that Word repeats at the top of every page the
// table continues onto. Building the continuation is the third thing this path
// has had to carry over (footnote refs, column widths, now heading rows), so it
// lives in one place: buildContinuationTable.
async function buildRepeatingHeaderDocx(
  bodyRows: number,
  headerText: string,
  pageHeight = 2600,
  headerFootnote = false,
): Promise<string> {
  const noteRef = headerFootnote
    ? `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>`
    : ''
  const rows: string[] = [
    `<w:tr><w:trPr><w:tblHeader/></w:trPr>` +
      `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>` +
      `<w:p><w:r><w:t>${headerText}</w:t></w:r>${noteRef}</w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>` +
      `<w:p><w:r><w:t>HEAD-B</w:t></w:r></w:p></w:tc></w:tr>`,
  ]
  for (let i = 0; i < bodyRows; i++) {
    rows.push(
      `<w:tr><w:tc><w:p><w:r><w:t>R${i + 1}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:p><w:r><w:t>body ${i + 1}</w:t></w:r></w:p></w:tc></w:tr>`,
    )
  }
  const body =
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/>` +
    `</w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>` +
    rows.join('') + `</w:tbl>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="${pageHeight}"/>` +
    `<w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>`

  const zip = new JSZip()
  const notesOverride = headerFootnote
    ? `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`
    : ''
  zip.file('[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      notesOverride + `</Types>`)
  zip.file('_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
  if (headerFootnote) {
    zip.file('word/footnotes.xml',
      `<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="2"><w:p><w:r><w:t>the heading note</w:t></w:r></w:p></w:footnote></w:footnotes>`)
  }
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      (headerFootnote
        ? `<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>`
        : '') +
      `</Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`)
  return (await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })).toString('base64')
}

async function headerTablePieces(b64: string) {
  const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
  await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
  await page.addScriptTag({ content: bundleJs })
  const result = await page.evaluate(async (b64: string) => {
    const bin = atob(b64)
    const arr = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const SD = (window as any).SimpleDoc
    SD.render(await SD.parse(arr.buffer), document.getElementById('view'))
    const pieces: { rows: number; firstCell: string; headers: number; widths: number[] }[] = []
    for (const pg of Array.from(document.querySelectorAll('.ssd-page')) as HTMLElement[]) {
      if (pg.style.visibility === 'hidden') continue
      for (const t of Array.from(pg.querySelectorAll('table')) as HTMLTableElement[]) {
        pieces.push({
          rows: t.rows.length,
          firstCell: t.rows[0]?.cells[0]?.textContent ?? '',
          headers: t.querySelectorAll('tr[data-ssd-header]').length,
          widths: Array.from(t.rows[0]?.cells ?? []).map(c => Math.round(c.getBoundingClientRect().width)),
        })
      }
    }
    return {
      pieces,
      fnrefIds: document.querySelectorAll('sup[id^="fnref-"]').length,
      fnBoxes: document.querySelectorAll('.ssd-footnotes').length,
    }
  }, b64)
  await page.close()
  return result
}

describe('page-aware pagination: repeating heading rows (w:tblHeader)', () => {
  it('repeats the heading row on every continued piece', async () => {
    const { pieces } = await headerTablePieces(await buildRepeatingHeaderDocx(18, 'HEAD-A'))
    expect(pieces.length).toBeGreaterThanOrEqual(3)
    for (const piece of pieces) {
      expect(piece.firstCell).toBe('HEAD-A')
      expect(piece.headers).toBe(1)
      // The repeated heading and the carried column widths coexist.
      expect(piece.widths[0]).toBeLessThan(piece.widths[1])
    }
  }, 30_000)

  it('does not register the heading footnote again on each continuation', async () => {
    const { pieces, fnrefIds, fnBoxes } = await headerTablePieces(
      await buildRepeatingHeaderDocx(18, 'HEAD-A', 2600, true))
    expect(pieces.length).toBeGreaterThanOrEqual(2)
    // The reference id lives on the ORIGINAL heading only. A repeated heading
    // keeps the visible marker but drops the id, so the note is not counted (and
    // rendered) once per page, and the back-link anchor stays unique.
    expect(fnrefIds).toBe(1)
    expect(fnBoxes).toBe(1)
  }, 30_000)

  it('terminates when the heading alone overflows the page', async () => {
    // The paginator loops until the continuation is strictly smaller. Repeating a
    // heading taller than the page would grow it back — so the repeat is skipped
    // when the cut lands inside the heading block.
    const { pieces } = await headerTablePieces(
      await buildRepeatingHeaderDocx(6, 'HEAD '.repeat(300), 900))
    expect(pieces.length).toBeGreaterThan(0)
  }, 30_000)
})

// A one-row table whose three cells exercise the cell-orientation properties:
// a plain cell, a bottom-aligned cell, and a rotated cell. The row is given a
// large w:trHeight so vertical alignment has room to be visible.
async function buildCellOrientationDocx(textDirection: 'tbRl' | 'btLr'): Promise<string> {
  const cell = (tcPr: string, text: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

  const body =
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:bottom w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2250"/><w:gridCol w:w="2250"/>` +
    `<w:gridCol w:w="2250"/><w:gridCol w:w="2250"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:trHeight w:val="2000"/></w:trPr>` +
    cell(``, 'plain') +
    cell(`<w:vAlign w:val="bottom"/>`, 'bottom') +
    cell(`<w:vAlign w:val="center"/>`, 'middle') +
    cell(`<w:textDirection w:val="${textDirection}"/><w:vAlign w:val="center"/>`, 'rotated') +
    `</w:tr></w:tbl>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

async function cellOrientation(b64: string) {
  const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
  await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
  await page.addScriptTag({ content: bundleJs })
  const result = await page.evaluate(async (b64: string) => {
    const bin = atob(b64)
    const arr = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const SD = (window as any).SimpleDoc
    const doc = await SD.parse(arr.buffer)
    SD.render(doc, document.getElementById('view'))
    const cells = Array.from(document.querySelectorAll('td')) as HTMLElement[]
    const read = (td: HTMLElement) => {
      const cs = getComputedStyle(td)
      // Where the text actually sits inside the cell, as a fraction of the cell
      // height: computed vertical-align only proves the property was set, this
      // proves the layout engine moved the content.
      const cellBox = td.getBoundingClientRect()
      const inner = (td.querySelector('p, div') as HTMLElement | null) ?? td
      const innerBox = inner.getBoundingClientRect()
      return {
        verticalAlign: cs.verticalAlign,
        writingMode: cs.writingMode,
        text: (td.textContent ?? '').trim(),
        cellHeight: Math.round(cellBox.height),
        contentCenterFraction:
          cellBox.height > 0
            ? (innerBox.top + innerBox.height / 2 - cellBox.top) / cellBox.height
            : -1,
        innerTransform: (td.firstElementChild as HTMLElement | null)
          ? getComputedStyle(td.firstElementChild as HTMLElement).transform
          : 'none',
      }
    }
    return cells.map(read)
  }, b64)
  await page.close()
  return result
}

describe('table cell orientation (w:vAlign, w:textDirection)', () => {
  it('puts each w:vAlign on its own third of the cell', async () => {
    const [plain, bottom, middle] = await cellOrientation(await buildCellOrientationDocx('tbRl'))
    // 'center' is the OOXML name; CSS spells the same thing 'middle'.
    expect(plain.verticalAlign).toBe('top')
    expect(bottom.verticalAlign).toBe('bottom')
    expect(middle.verticalAlign).toBe('middle')
    // The row is ~133px tall for one line of text, so the three cells' content
    // must actually sit in different bands of the cell, not just carry the CSS.
    expect(plain.contentCenterFraction).toBeLessThan(0.4)
    expect(middle.contentCenterFraction).toBeGreaterThan(0.4)
    expect(middle.contentCenterFraction).toBeLessThan(0.6)
    expect(bottom.contentCenterFraction).toBeGreaterThan(0.6)
  }, 30_000)

  it('rotates a w:textDirection="tbRl" cell with writing-mode on the cell itself', async () => {
    const [, , , rotated] = await cellOrientation(await buildCellOrientationDocx('tbRl'))
    expect(rotated.writingMode).toBe('vertical-rl')
    expect(rotated.text).toBe('rotated')
    // tbRl reads top-to-bottom, so no counter-rotation wrapper is added.
    expect(rotated.innerTransform).toBe('none')
  }, 30_000)

  it('turns a w:textDirection="btLr" cell 180° on an inner element', async () => {
    const [, , , rotated] = await cellOrientation(await buildCellOrientationDocx('btLr'))
    expect(rotated.writingMode).toBe('vertical-rl')
    expect(rotated.text).toBe('rotated')
    // matrix(-1, 0, 0, -1, 0, 0) is rotate(180deg). It must be on a child, not
    // on the cell: rotating the cell would flip its background and borders too.
    expect(rotated.innerTransform).toBe('matrix(-1, 0, 0, -1, 0, 0)')
  }, 30_000)
})

// A rotated cell with NO explicit row height, in a fixed-layout table (w:tblGrid
// present). This is the clipping risk: if the table sizing algorithm does not
// measure the rotated text, the row stays one line tall and the label spills out
// of the cell instead of the row growing to hold it.
async function buildTallRotatedCellDocx(): Promise<string> {
  const long = 'Rotated label that is far longer than one line of a cell'
  const body =
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:bottom w:val="single" w:sz="4" w:color="000000"/>` +
    `<w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="8000"/></w:tblGrid>` +
    `<w:tr>` +
    `<w:tc><w:tcPr><w:textDirection w:val="btLr"/></w:tcPr>` +
    `<w:p><w:r><w:t>${long}</w:t></w:r></w:p></w:tc>` +
    `<w:tc><w:tcPr/><w:p><w:r><w:t>short</w:t></w:r></w:p></w:tc>` +
    `</w:tr></w:tbl>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

describe('rotated cell sizing', () => {
  it('grows the row to fit the rotated text instead of clipping it', async () => {
    const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
    await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
    await page.addScriptTag({ content: bundleJs })
    const r = await page.evaluate(async (b64: string) => {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const SD = (window as any).SimpleDoc
      const doc = await SD.parse(arr.buffer)
      SD.render(doc, document.getElementById('view'))
      const td = document.querySelectorAll('td')[0] as HTMLElement
      const cellBox = td.getBoundingClientRect()
      const range = document.createRange()
      range.selectNodeContents(td)
      const textBox = range.getBoundingClientRect()
      return {
        cellHeight: Math.round(cellBox.height),
        textHeight: Math.round(textBox.height),
        // Positive means the text paints below the cell's bottom edge.
        overflowPx: Math.round(textBox.bottom - cellBox.bottom),
      }
    }, await buildTallRotatedCellDocx())
    await page.close()
    // The label is ~380px of rotated text; the row must be at least that tall.
    expect(r.textHeight).toBeGreaterThan(200)
    expect(r.cellHeight).toBeGreaterThanOrEqual(r.textHeight)
    expect(r.overflowPx).toBeLessThanOrEqual(2)
  }, 30_000)
})

// A numbered list long enough to cross a page break, preceded by filler that
// pushes its start down the page. `nested` puts a deeper-level item in the
// middle, which renderBlocks attaches as a bare <ol> inside the list.
async function buildLongListDocx(
  items: number,
  itemText = 'Item text long enough to occupy a real line of the page',
  fillerParas = 6,
  nested = false,
  pageTwips = 5000,
  ordered = true,
): Promise<string> {
  const filler = Array.from({ length: fillerParas }, (_, i) =>
    `<w:p><w:r><w:t>Filler paragraph ${i + 1} standing between the top of the page and the list.</w:t></w:r></w:p>`,
  ).join('')
  const li = (i: number, ilvl = 0) =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
    `<w:r><w:t>${i}. ${itemText}</w:t></w:r></w:p>`
  const list = Array.from({ length: items }, (_, i) =>
    nested && i === 2 ? li(i + 1) + li(99, 1) : li(i + 1),
  ).join('')

  // A short page (default ~3.5in of height) so a modest list is guaranteed to
  // cross a break; a letter-height page swallows 40 items whole and the test
  // would pass without ever exercising a split.
  const body =
    filler + list +
    `<w:sectPr><w:pgSz w:w="12240" w:h="${pageTwips}"/>` +
    `<w:pgMar w:top="400" w:bottom="400" w:left="600" w:right="600"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
  )
  zip.file(
    'word/numbering.xml',
    `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:abstractNum w:abstractNumId="0">` +
      `<w:lvl w:ilvl="0"><w:numFmt w:val="${ordered ? 'decimal' : 'bullet'}"/>` +
      `<w:lvlText w:val="${ordered ? '%1.' : '-'}"/></w:lvl>` +
      `<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/></w:lvl>` +
      `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

async function listPages(b64: string) {
  const page = await browser.newPage({ viewport: { width: 1000, height: 1400 } })
  await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
  await page.addScriptTag({ content: bundleJs })
  const r = await page.evaluate(async (b64: string) => {
    const bin = atob(b64)
    const arr = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const SD = (window as any).SimpleDoc
    const doc = await SD.parse(arr.buffer)
    SD.render(doc, document.getElementById('view'))
    const pages = Array.from(document.querySelectorAll('.ssd-page')) as HTMLElement[]
    return {
      nPages: pages.length,
      // Content that spills past a page's declared box stretches it, because the
      // box is sized with min-height. Real height above min-height is the signal.
      overflowing: pages.filter(p => {
        const min = parseFloat(getComputedStyle(p).minHeight)
        return p.getBoundingClientRect().height > min + 1
      }).length,
      perPage: pages.map(p => ({
        lists: Array.from(p.querySelectorAll('ol')).map(ol => ({
          start: (ol as HTMLOListElement).start,
          items: ol.querySelectorAll(':scope > li').length,
          nested: ol.querySelectorAll(':scope > ol').length,
        })),
        items: p.querySelectorAll('li').length,
        text: (p.textContent ?? '').replace(/\s+/g, ' ').trim(),
      })),
      totalItems: document.querySelectorAll('li').length,
      // A <ul> carries no counter: buildContinuationList must not stamp `start`
      // on a bullet continuation.
      bulletContinuations: Math.max(0, document.querySelectorAll('ul').length - 1),
      bulletsWithStart: Array.from(document.querySelectorAll('ul'))
        .filter(ul => ul.hasAttribute('start')).length,
    }
  }, b64)
  await page.close()
  return r
}

describe('page-aware pagination: lists split across pages', () => {
  it('continues a long list on the next page instead of moving it whole', async () => {
    const r = await listPages(await buildLongListDocx(40))
    expect(r.nPages).toBeGreaterThan(1)
    // The list starts partway down page 1, so page 1 must carry SOME items. The
    // bug this covers moved the whole list to page 2 and left page 1 half empty.
    expect(r.perPage[0].items).toBeGreaterThan(0)
    expect(r.perPage[1].items).toBeGreaterThan(0)
    expect(r.totalItems).toBe(40)
  }, 30_000)

  it('does not overflow any page box', async () => {
    const r = await listPages(await buildLongListDocx(40))
    expect(r.overflowing).toBe(0)
  }, 30_000)

  it('resumes the numbering on the continuation instead of restarting at 1', async () => {
    const r = await listPages(await buildLongListDocx(40))
    const firstPageItems = r.perPage[0].lists.reduce((a, l) => a + l.items, 0)
    const continuation = r.perPage[1].lists[0]
    expect(continuation).toBeDefined()
    // Item n+1 must be numbered n+1 on the next page, not 1.
    expect(continuation.start).toBe(firstPageItems + 1)
  }, 30_000)

  it('keeps every item exactly once across the split', async () => {
    const r = await listPages(await buildLongListDocx(40))
    // No \b before the digit: textContent runs adjacent items together
    // ("...of the page2. Item text..."), and there is no word boundary between
    // "e" and "2", so the anchored pattern found only the first item per page.
    const seen = r.perPage.flatMap(p => p.text.match(/\d+\. Item text/g) ?? [])
    expect(seen.length).toBe(40)
    expect(new Set(seen).size).toBe(40)
  }, 30_000)

  it('never separates a nested list from the item above it', async () => {
    // renderBlocks attaches a deeper level to `parent.lastLi ?? parent.el`, so a
    // nested <ol> can sit as a direct child of the list. Cutting immediately
    // before it would strand it at the top of the next page under no item.
    const r = await listPages(await buildLongListDocx(40, undefined, 6, true))
    for (const page of r.perPage) {
      const first = page.lists[0]
      if (!first) continue
      expect(first.items).toBeGreaterThan(0)
    }
    expect(r.overflowing).toBe(0)
  }, 30_000)

  it('splits a bullet list too, without inventing a counter for it', async () => {
    // splitterFor accepts UL as well as OL. A <ul> has no counter, so
    // buildContinuationList must leave `start` alone rather than stamping one on.
    const r = await listPages(await buildLongListDocx(40, undefined, 6, false, 5000, false))
    expect(r.nPages).toBeGreaterThan(1)
    expect(r.totalItems).toBe(40)
    expect(r.overflowing).toBe(0)
    expect(r.perPage.every(p => p.lists.length === 0)).toBe(true) // no <ol> anywhere
    expect(r.bulletContinuations).toBeGreaterThan(0)
    expect(r.bulletsWithStart).toBe(0)
  }, 30_000)

  it('terminates when a single item is taller than a whole page', async () => {
    // One item that cannot fit any page: the paginator must accept the overflow
    // rather than loop forever trying to split it onto a fresh page.
    const r = await listPages(await buildLongListDocx(3, 'WORD '.repeat(2000), 2, false, 4000))
    expect(r.nPages).toBeGreaterThan(0)
    expect(r.totalItems).toBe(3)
  }, 30_000)
})

// A template-shaped .docx: a full-page behindDoc background image plus enough
// text to fill several pages. This is the only shape that reaches
// renderPageBgPaginated, the second paginator, which assigns whole blocks to
// pages from pre-measured heights instead of moving DOM nodes.
async function buildPageBackgroundDocx(paras = 40): Promise<string> {
  // 1x1 transparent PNG — the smallest thing that parses as an image.
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  )
  // A behindDoc anchor sized to the whole page marks the background.
  const bgAnchor =
    `<w:p><w:r><w:drawing>` +
    `<wp:anchor behindDoc="1" distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="7772400" cy="10058400"/>` +
    `<wp:docPr id="1" name="bg"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:blipFill><a:blip r:embed="rId3"/></pic:blipFill>` +
    `</pic:pic></a:graphicData></a:graphic></wp:anchor>` +
    `</w:drawing></w:r></w:p>`

  // Spaced paragraphs: the margins between them are exactly what a per-block
  // offsetHeight measurement drops.
  const text = Array.from({ length: paras }, (_, i) =>
    `<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>` +
    `<w:r><w:t>Paragraph ${i + 1} carrying enough words to take a full line of the page box.</w:t></w:r></w:p>`,
  ).join('')

  const body =
    bgAnchor + text +
    `<w:sectPr><w:pgSz w:w="12240" w:h="6000"/>` +
    `<w:pgMar w:top="500" w:bottom="500" w:left="600" w:right="600"/></w:sectPr>`

  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bg.png"/></Relationships>`,
  )
  zip.file('word/media/bg.png', png)
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ` +
      `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>${body}</w:body></w:document>`,
  )
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  return buf.toString('base64')
}

describe('page-background pagination: page boxes hold their content', () => {
  it('never lets a page grow past its own height', async () => {
    // The page box is sized with min-height, so content that does not fit
    // stretches the sheet instead of visibly breaking. Measuring each block on
    // its own dropped the margins between them and overfilled the page by 12px
    // on a real template, with nothing on screen to show for it.
    const b64 = await buildPageBackgroundDocx(40)
    const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } })
    await page.setContent('<!doctype html><meta charset="utf-8"><div id="view"></div>')
    await page.addScriptTag({ content: bundleJs })
    const r = await page.evaluate(async (b64: string) => {
      const bin = atob(b64)
      const arr = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const SD = (window as any).SimpleDoc
      const doc = await SD.parse(arr.buffer)
      SD.render(doc, document.getElementById('view'))
      const pages = Array.from(document.querySelectorAll('.ssd-page')) as HTMLElement[]
      return {
        nPages: pages.length,
        // Took the bg path (a background image is drawn) rather than the plain one.
        withBg: pages.filter(p => p.style.backgroundImage && p.style.backgroundImage !== 'none').length,
        overflow: pages.map(p =>
          Math.round(p.getBoundingClientRect().height - parseFloat(getComputedStyle(p).minHeight)),
        ),
        paragraphs: (document.body.textContent ?? '').match(/Paragraph \d+ carrying/g)?.length ?? 0,
      }
    }, b64)
    await page.close()

    expect(r.withBg).toBeGreaterThan(0) // guard: this fixture must reach the bg paginator
    expect(r.nPages).toBeGreaterThan(1)
    expect(Math.max(...r.overflow)).toBeLessThanOrEqual(1)
    expect(r.paragraphs).toBe(40) // and nothing was dropped to make it fit
  }, 30_000)
})
