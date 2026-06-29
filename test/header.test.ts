import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

// Build a minimal .docx with a default header part (headerN.xml) referenced from
// the body sectPr, plus a footer for contrast, to prove the header is resolved
// and parsed the same way the footer is.
async function buildDocxWithHeader(
  headerXml: string | null,
  footerXml: string | null,
  pgMar = `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>`,
): Promise<ArrayBuffer> {
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

  const rels: string[] = [
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
  ]
  const refs: string[] = []
  if (headerXml != null) {
    zip.file('word/header1.xml',
      `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${headerXml}</w:hdr>`)
    rels.push(`<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`)
    refs.push(`<w:headerReference w:type="default" r:id="rId10"/>`)
  }
  if (footerXml != null) {
    zip.file('word/footer1.xml',
      `<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${footerXml}</w:ftr>`)
    rels.push(`<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`)
    refs.push(`<w:footerReference w:type="default" r:id="rId11"/>`)
  }

  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    rels.join('') + `</Relationships>`)

  const sectPr =
    `<w:sectPr>${refs.join('')}` +
    `<w:pgSz w:w="12240" w:h="15840"/>${pgMar}</w:sectPr>`
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body><w:p><w:r><w:t>Body text</w:t></w:r></w:p>${sectPr}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const texts = (block: ParagraphBlock): string[] =>
  block.runs.map(r => (r as TextRun).text)

describe('default header (w:headerReference)', () => {
  it('resolves and parses the default header part', async () => {
    const doc = await parse(await buildDocxWithHeader(
      `<w:p><w:r><w:t>My Header</w:t></w:r></w:p>`,
      null,
    ))
    expect(doc.header).toBeDefined()
    expect(doc.header!.length).toBe(1)
    expect(texts(doc.header![0] as ParagraphBlock)).toEqual(['My Header'])
  })

  it('parses the header offset (w:header) from pgMar into headerPx', async () => {
    const doc = await parse(await buildDocxWithHeader(
      `<w:p><w:r><w:t>H</w:t></w:r></w:p>`,
      null,
    ))
    // 720 twips * 96 / 1440 = 48px
    expect(doc.pageSize?.headerPx).toBe(48)
  })

  it('renders a PAGE field in the header as the page number', async () => {
    const doc = await parse(await buildDocxWithHeader(
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
      null,
    ))
    expect(doc.header).toBeDefined()
    const run = (doc.header![0] as ParagraphBlock).runs.find(r => (r as TextRun).pageNumber)
    expect(run).toBeDefined()
  })

  it('resolves header and footer independently', async () => {
    const doc = await parse(await buildDocxWithHeader(
      `<w:p><w:r><w:t>Top</w:t></w:r></w:p>`,
      `<w:p><w:r><w:t>Bottom</w:t></w:r></w:p>`,
    ))
    expect(texts(doc.header![0] as ParagraphBlock)).toEqual(['Top'])
    expect(texts(doc.footer![0] as ParagraphBlock)).toEqual(['Bottom'])
  })

  it('omits header when there is no headerReference', async () => {
    const doc = await parse(await buildDocxWithHeader(null, null))
    expect(doc.header).toBeUndefined()
  })
})
