import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import { extractPPr } from '../src/parser/styles.js'
import { XMLParser } from 'fast-xml-parser'
import type { ParagraphBlock } from '../src/types.js'

const xmlParser = new XMLParser({
  removeNSPrefix: true, attributeNamePrefix: '', ignoreAttributes: false, parseAttributeValue: false,
})
const pPrOf = (inner: string) =>
  xmlParser.parse(`<w:pPr xmlns:w="x">${inner}</w:pPr>`).pPr as Record<string, unknown>

// docDefaultsSpacing goes into <w:docDefaults><w:pPrDefault>, so a paragraph's
// own spacing has something to override — which is where this went wrong.
async function buildDocx(paraSpacing: string, docDefaultsSpacing = ''): Promise<ArrayBuffer> {
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
    `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault>` +
    (docDefaultsSpacing ? `<w:pPrDefault><w:pPr>${docDefaultsSpacing}</w:pPr></w:pPrDefault>` : '') +
    `</w:docDefaults></w:styles>`)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
    `<w:p><w:pPr>${paraSpacing}</w:pPr><w:r><w:t>Text of the paragraph</w:t></w:r></w:p>` +
    `</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const styleOf = async (paraSpacing: string, docDefaultsSpacing = '') => {
  const doc = await parse(await buildDocx(paraSpacing, docDefaultsSpacing))
  return (doc.blocks[0] as ParagraphBlock).style
}

const SINGLE = `<w:spacing w:line="240" w:lineRule="auto"/>`
const ONE_AND_HALF = `<w:spacing w:line="360" w:lineRule="auto"/>`
const DEFAULTS_115 = `<w:spacing w:line="276" w:lineRule="auto"/>`

describe('line spacing (w:spacing w:line)', () => {
  it('reads single spacing as the font\'s own line, not the number 1', () => {
    // 240/240 = one LINE, and a line is ascent + descent + gap. Emitting the
    // number 1 set every line to exactly the font size, ~25% tighter than Word.
    const s = extractPPr(pPrOf(SINGLE))
    expect(s.lineHeightSingle).toBe(true)
    expect(s.lineHeight).toBeUndefined()
  })

  it('reads a non-single multiplier as a number', () => {
    const s = extractPPr(pPrOf(ONE_AND_HALF))
    expect(s.lineHeight).toBeCloseTo(1.5)
    expect(s.lineHeightSingle).toBe(false)
  })

  it('reads an exact rule as a pixel height', () => {
    const s = extractPPr(pPrOf(`<w:spacing w:line="480" w:lineRule="exact"/>`))
    expect(s.lineHeightPx).toBe(32) // 480 twips = 24pt = 32px
    expect(s.lineHeightSingle).toBeUndefined()
  })

  it('lets an explicit single override an inherited multiplier', async () => {
    // The bug this pins: styles merge with Object.assign, so leaving single
    // spacing ABSENT let a docDefaults of 1.15 lines survive and the paragraph
    // silently kept 1.15. Both branches must write for either to win.
    const s = await styleOf(SINGLE, DEFAULTS_115)
    expect(s.lineHeightSingle).toBe(true)
    // And the IR must not still report a multiplier nothing uses.
    expect(s.lineHeight).toBeUndefined()
  })

  it('lets an explicit multiplier override an inherited single', async () => {
    const s = await styleOf(ONE_AND_HALF, SINGLE)
    expect(s.lineHeightSingle).toBe(false)
    expect(s.lineHeight).toBeCloseTo(1.5)
  })

  it('inherits the document default when the paragraph says nothing', async () => {
    const s = await styleOf('', DEFAULTS_115)
    expect(s.lineHeight).toBeCloseTo(1.15)
    expect(s.lineHeightSingle).toBe(false)
  })

  it('leaves a document that specifies no spacing at all alone', async () => {
    const s = await styleOf('')
    expect(s.lineHeight).toBeUndefined()
    expect(s.lineHeightSingle).toBeUndefined()
  })
})
