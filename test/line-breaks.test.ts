import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

// A <w:br/>, <w:cr/> or page break can sit BETWEEN two <w:t> elements of the same
// run — the shape Word produces for Shift+Enter / Ctrl+Enter mid-paragraph.
// fast-xml-parser groups a run's children by tag, so the parser has to recover
// their order from raw XML; these tests pin that down for the body AND for the
// header/footer/note parts, which are parsed through a different entry point.

async function buildDocx(body: string, headerXml?: string): Promise<ArrayBuffer> {
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

  const rels = [`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`]
  let sectPr = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`
  if (headerXml) {
    zip.file('word/header1.xml',
      `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${headerXml}</w:hdr>`)
    zip.file('word/_rels/header1.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>`)
    rels.push(`<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`)
    sectPr = `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>` +
      `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`
  }
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    rels.join('') + `</Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sectPr}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const p0 = (doc: { blocks: unknown[] }): ParagraphBlock => doc.blocks[0] as ParagraphBlock
const texts = (p: ParagraphBlock): string[] => p.runs.map(r => (r as TextRun).text)
const breaks = (p: ParagraphBlock): boolean[] => p.runs.map(r => !!(r as TextRun).lineBreak)

describe('breaks inside a run, in document order', () => {
  it('puts a <w:br/> BETWEEN the two texts, not after both', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['A', 'B'])
    // The break belongs to 'A' — it renders after that run.
    expect(breaks(p0(doc))).toEqual([true, false])
  })

  it('keeps N consecutive breaks as N breaks', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>A</w:t><w:br/><w:br/><w:br/><w:t>B</w:t></w:r></w:p>`))
    // Three breaks: one closing 'A', then two empty runs carrying the rest.
    expect(breaks(p0(doc)).filter(Boolean).length).toBe(3)
    expect(texts(p0(doc)).join('')).toBe('AB')
  })

  it('treats <w:cr/> the same as a soft break', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>A</w:t><w:cr/><w:t>B</w:t></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['A', 'B'])
    expect(breaks(p0(doc))).toEqual([true, false])
  })

  // Ctrl+Enter mid-paragraph: Word packs the page break INTO the run that holds
  // the surrounding text. Treating the break as the whole run dropped both texts.
  it('does not lose text sharing a run with a page break', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>ANTES</w:t><w:br w:type="page"/><w:t>DEPOIS</w:t></w:r></w:p>`))
    expect(doc.blocks.length).toBe(2)
    expect(texts(p0(doc))).toEqual(['ANTES'])
    expect(texts(doc.blocks[1] as ParagraphBlock)).toEqual(['DEPOIS'])
    expect((doc.blocks[1] as ParagraphBlock).pageBreakBefore).toBe(true)
  })

  it('splits at a page break that follows the text in the same run', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>ONLY</w:t><w:br w:type="page"/></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['ONLY'])
    expect(doc.blocks.length).toBe(2)
  })

  // GUARD-RAIL: the shapes that already rendered correctly must not change.
  it('GUARD keeps a page break in its own run working', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>ANTES</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>DEPOIS</w:t></w:r></w:p>`))
    expect(doc.blocks.length).toBe(2)
    expect(texts(p0(doc))).toEqual(['ANTES'])
    expect(texts(doc.blocks[1] as ParagraphBlock)).toEqual(['DEPOIS'])
  })

  it('GUARD keeps a trailing <w:br/> on its own run', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>A</w:t><w:br/></w:r><w:r><w:t>B</w:t></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['A', 'B'])
    expect(breaks(p0(doc))).toEqual([true, false])
  })

  it('GUARD keeps tabs leading the segment that follows them', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>V.4</w:t><w:tab/><w:tab/><w:tab/><w:t>CONFIDENCIAL</w:t></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['V.4', 'CONFIDENCIAL'])
    expect((p0(doc).runs[0] as TextRun).tabs).toBeUndefined()
    expect((p0(doc).runs[1] as TextRun).tabs).toBe(3)
  })

  // An empty run parses to the string "" rather than an object; the `in` checks
  // used to inspect a run threw a TypeError on it, failing the WHOLE document.
  it.each(['<w:r/>', '<w:r></w:r>', '<w:r> </w:r>'])(
    'does not throw on the empty run %s', async (emptyRun) => {
      const doc = await parse(await buildDocx(`<w:p>${emptyRun}<w:r><w:t>after</w:t></w:r></w:p>`))
      expect(texts(p0(doc))).toContain('after')
    })

  // An empty self-closed run has no XML slice of its own. It must not shift the
  // slices of the runs after it, or they parse against the wrong run's children.
  it('keeps run/xml alignment across a self-closed empty run', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r/><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p>`))
    expect(texts(p0(doc)).join('')).toBe('AB')
    const withText = p0(doc).runs.filter(r => (r as TextRun).text !== '')
    expect((withText[0] as TextRun).text).toBe('A')
    expect((withText[0] as TextRun).lineBreak).toBe(true)
    expect((withText[1] as TextRun).lineBreak).toBeUndefined()
  })

  it('orders a tab and a break packed into one run', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['A', 'B', 'C'])
    expect((p0(doc).runs[1] as TextRun).tabs).toBe(1)
    expect(breaks(p0(doc))).toEqual([false, true, false])
  })
})

// Headers, footers and notes are parsed through their own entry points, which
// previously passed no raw XML at all — so neither run order NOR a run's child
// order was recovered there.
describe('document order in a header part', () => {
  const header = (doc: { header?: unknown[] }): ParagraphBlock =>
    (doc.header as ParagraphBlock[])[0]

  it('puts a break between the two texts of a header run', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>Body</w:t></w:r></w:p>`,
      `<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p>`))
    expect(texts(header(doc))).toEqual(['A', 'B'])
    expect(breaks(header(doc))).toEqual([true, false])
  })

  it('keeps a hyperlink in the MIDDLE of a header paragraph in place', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>Body</w:t></w:r></w:p>`,
      `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>LINK</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`))
    expect(texts(header(doc))).toEqual(['before ', 'LINK', ' after'])
    expect((header(doc).runs[1] as TextRun).href).toBe('https://example.com')
  })
})
