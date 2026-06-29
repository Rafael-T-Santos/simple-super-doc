import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock } from '../src/types.js'

// portrait Letter
const PORTRAIT = `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>`
// landscape Letter
const LANDSCAPE = `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>`

async function buildDocx(documentBody: string): Promise<ArrayBuffer> {
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
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('multiple sections (w:sectPr)', () => {
  it('splits a portrait section then a landscape section', async () => {
    const body =
      `<w:p><w:r><w:t>portrait body</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr>${PORTRAIT}</w:sectPr></w:pPr></w:p>` +
      `<w:p><w:r><w:t>landscape body</w:t></w:r></w:p>` +
      `<w:sectPr>${LANDSCAPE}</w:sectPr>`
    const doc = await parse(await buildDocx(body))
    expect(doc.sections).toBeDefined()
    expect(doc.sections!.length).toBe(2)

    const [s1, s2] = doc.sections!
    // section 1 = portrait (width < height)
    expect(s1.pageSize.widthPx).toBeLessThan(s1.pageSize.heightPx)
    // section 2 = landscape (width > height)
    expect(s2.pageSize.widthPx).toBeGreaterThan(s2.pageSize.heightPx)

    // doc.pageSize is the body (last) section = landscape
    expect(doc.pageSize!.widthPx).toBeGreaterThan(doc.pageSize!.heightPx)
  })

  it('strips the transient sectionPageSize tag from the IR', async () => {
    const body =
      `<w:p><w:r><w:t>a</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr>${PORTRAIT}</w:sectPr></w:pPr></w:p>` +
      `<w:p><w:r><w:t>b</w:t></w:r></w:p>` +
      `<w:sectPr>${LANDSCAPE}</w:sectPr>`
    const doc = await parse(await buildDocx(body))
    for (const b of doc.blocks) {
      expect((b as ParagraphBlock).sectionPageSize).toBeUndefined()
    }
  })

  it('does not add a sections array for a single-section document', async () => {
    const body = `<w:p><w:r><w:t>only</w:t></w:r></w:p><w:sectPr>${PORTRAIT}</w:sectPr>`
    const doc = await parse(await buildDocx(body))
    expect(doc.sections).toBeUndefined()
  })

  it('concatenated section blocks equal the flat block list', async () => {
    const body =
      `<w:p><w:r><w:t>one</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr>${PORTRAIT}</w:sectPr></w:pPr></w:p>` +
      `<w:p><w:r><w:t>two</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>three</w:t></w:r></w:p>` +
      `<w:sectPr>${LANDSCAPE}</w:sectPr>`
    const doc = await parse(await buildDocx(body))
    const flat = doc.sections!.flatMap(s => s.blocks)
    expect(flat.length).toBe(doc.blocks.length)
    expect(flat).toEqual(doc.blocks)
  })
})
