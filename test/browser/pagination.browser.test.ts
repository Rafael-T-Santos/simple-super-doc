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
