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

  it('splits a run\'s <w:t> segments around tabs, preserving order (Google Docs packing)', async () => {
    // A run holding "text <w:tab/><w:tab/> text" — fast-xml-parser groups r.t
    // into an array and loses the tab position. The text must survive AND stay
    // split around the tabs so the trailing segment can right-align in a footer.
    const doc = await parse(await buildDocxWithHeader(
      null,
      `<w:p><w:r>` +
      `<w:t xml:space="preserve">V.4 - 04.03.2026 </w:t><w:tab/><w:tab/>` +
      `<w:t xml:space="preserve">Classificação: CONFIDENCIAL</w:t>` +
      `</w:r></w:p>`,
    ))
    const runs = (doc.footer![0] as ParagraphBlock).runs as TextRun[]
    expect(runs.map(r => r.text)).toEqual([
      'V.4 - 04.03.2026 ',
      'Classificação: CONFIDENCIAL',
    ])
    // The tabs attach to the trailing segment (it right-aligns), not the first.
    expect(runs[0].tabs ?? 0).toBe(0)
    expect(runs[1].tabs).toBe(2)
  })
})

// A tiny .docx whose default header holds an anchored image referenced through
// the header's OWN relationships file (word/_rels/header1.xml.rels), not the
// document's. Proves header/footer parts resolve images against their own rId
// namespace.
async function buildDocxWithHeaderImage(): Promise<ArrayBuffer> {
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
    `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults></w:styles>`)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
    `</Relationships>`)
  // The header's rId1 must resolve here, NOT in document.xml.rels (where rId1
  // is the styles part). This is the exact collision the fix handles.
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
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="953" cy="953"/>` +
    `<a:graphic><a:graphicData><pic:pic><pic:blipFill>` +
    `<a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r></w:p></w:hdr>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>` +
    `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>` +
    `<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('header image via the part\'s own relationships', () => {
  it('resolves an anchored header image through header1.xml.rels', async () => {
    const doc = await parse(await buildDocxWithHeaderImage())
    expect(doc.header).toBeDefined()
    const runs = (doc.header![0] as ParagraphBlock).runs
    const image = runs.find(r => r.type === 'image')
    expect(image).toBeDefined()
    expect((image as { src: string }).src).toMatch(/^data:image\/png;base64,/)
    // The anchor offset is captured (EMU → px) so a letterhead logo can be
    // positioned (5000000 EMU / 9525 ≈ 525px, 95250 / 9525 = 10px).
    expect((image as { anchorXPx?: number }).anchorXPx).toBe(525)
    expect((image as { anchorYPx?: number }).anchorYPx).toBe(10)
  })
})
