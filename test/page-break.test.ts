import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

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
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const texts = (p: ParagraphBlock): string[] =>
  p.runs.map(r => (r as TextRun).text)

describe('explicit page break (w:br w:type="page")', () => {
  it('splits a paragraph at a mid-paragraph page break and forces a break', async () => {
    const body =
      `<w:p>` +
      `<w:r><w:t>before</w:t></w:r>` +
      `<w:r><w:br w:type="page"/></w:r>` +
      `<w:r><w:t>after</w:t></w:r>` +
      `</w:p>`
    const doc = await parse(await buildDocx(body))
    expect(doc.blocks.length).toBe(2)
    const a = doc.blocks[0] as ParagraphBlock
    const b = doc.blocks[1] as ParagraphBlock
    expect(texts(a)).toEqual(['before'])
    expect(a.pageBreakBefore).toBeFalsy()
    expect(texts(b)).toEqual(['after'])
    expect(b.pageBreakBefore).toBe(true)
    // The page-break marker itself must not survive as a run.
    expect([...a.runs, ...b.runs].some(r => (r as TextRun).pageBreak)).toBe(false)
  })

  it('handles a page break as its own paragraph', async () => {
    const body =
      `<w:p><w:r><w:t>one</w:t></w:r></w:p>` +
      `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
      `<w:p><w:r><w:t>two</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    // p1 -> 1 block; p2 -> 2 segments (empty before/after break); p3 -> 1 block
    expect(doc.blocks.length).toBe(4)
    expect(texts(doc.blocks[0] as ParagraphBlock)).toEqual(['one'])
    // The segment after the break forces a new page.
    expect((doc.blocks[2] as ParagraphBlock).pageBreakBefore).toBe(true)
    expect(texts(doc.blocks[3] as ParagraphBlock)).toEqual(['two'])
  })

  it('splits on multiple page breaks in one paragraph', async () => {
    const body =
      `<w:p>` +
      `<w:r><w:t>a</w:t></w:r><w:r><w:br w:type="page"/></w:r>` +
      `<w:r><w:t>b</w:t></w:r><w:r><w:br w:type="page"/></w:r>` +
      `<w:r><w:t>c</w:t></w:r>` +
      `</w:p>`
    const doc = await parse(await buildDocx(body))
    expect(doc.blocks.length).toBe(3)
    expect(texts(doc.blocks[0] as ParagraphBlock)).toEqual(['a'])
    expect(texts(doc.blocks[1] as ParagraphBlock)).toEqual(['b'])
    expect(texts(doc.blocks[2] as ParagraphBlock)).toEqual(['c'])
    expect((doc.blocks[1] as ParagraphBlock).pageBreakBefore).toBe(true)
    expect((doc.blocks[2] as ParagraphBlock).pageBreakBefore).toBe(true)
  })

  it('leaves a paragraph without a page break unsplit', async () => {
    const body = `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    expect(doc.blocks.length).toBe(1)
    expect(texts(doc.blocks[0] as ParagraphBlock)).toEqual(['plain'])
  })

  it('still treats a column break as dropped (no split)', async () => {
    const body =
      `<w:p><w:r><w:t>x</w:t></w:r><w:r><w:br w:type="column"/></w:r><w:r><w:t>y</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    expect(doc.blocks.length).toBe(1)
    expect(texts(doc.blocks[0] as ParagraphBlock)).toEqual(['x', 'y'])
  })
})
