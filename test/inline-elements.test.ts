import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

// Minimal .docx from a body fragment (mirrors the other parse tests).
async function buildDocx(body: string): Promise<ArrayBuffer> {
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
  return zip.generateAsync({ type: 'arraybuffer' })
}

const p0 = (doc: Awaited<ReturnType<typeof parse>>) => doc.blocks[0] as ParagraphBlock
const runs0 = (doc: Awaited<ReturnType<typeof parse>>) => p0(doc).runs as TextRun[]

describe('inline run elements (content that used to be dropped)', () => {
  it('w:sym → the symbol codepoint as text, carrying its symbol font', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:sym w:font="Wingdings" w:char="F0E0"/></w:r></w:p>`))
    const run = runs0(doc)[0]
    expect(run.text).toBe(String.fromCodePoint(0xf0e0))
    expect(run.style.fontFamily).toBe('Wingdings')
  })

  it('w:sym with an out-of-range char does not throw (robustness)', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:sym w:font="X" w:char="110000"/><w:t>ok</w:t></w:r></w:p>`))
    // The bad symbol is dropped; the rest of the run survives.
    expect(runs0(doc).map(r => r.text).join('')).toBe('ok')
  })

  it('w:cr → a soft line break', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>A</w:t><w:cr/><w:t>B</w:t></w:r></w:p>`))
    const rs = runs0(doc)
    expect(rs.map(r => r.text)).toEqual(['A', 'B'])
    expect(rs[rs.length - 1].lineBreak).toBe(true)
  })

  it('w:noBreakHyphen → U+2011 between the words', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>2024</w:t><w:noBreakHyphen/><w:t>2025</w:t></w:r></w:p>`))
    expect(runs0(doc).map(r => r.text).join('')).toBe('2024‑2025')
  })

  it('w:softHyphen → U+00AD (invisible, allows a break)', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:t>super</w:t><w:softHyphen/><w:t>cali</w:t></w:r></w:p>`))
    expect(runs0(doc).map(r => r.text).join('')).toBe('super­cali')
  })
})

describe('run/paragraph formatting (previously unmapped)', () => {
  it('w:vanish → style.hidden on the run', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r></w:p>`))
    expect(runs0(doc)[0].style.hidden).toBe(true)
  })

  it('w:caps / w:smallCaps / w:dstrike map to style flags', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>a</w:t></w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:smallCaps/></w:rPr><w:t>b</w:t></w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:dstrike/></w:rPr><w:t>c</w:t></w:r></w:p>`))
    expect((doc.blocks[0] as ParagraphBlock).runs[0].style.caps).toBe(true)
    expect((doc.blocks[1] as ParagraphBlock).runs[0].style.smallCaps).toBe(true)
    expect((doc.blocks[2] as ParagraphBlock).runs[0].style.doubleStrike).toBe(true)
  })

  it('paragraph-level w:shd → block background color', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:pPr><w:shd w:val="clear" w:fill="FFFF00"/></w:pPr><w:r><w:t>hi</w:t></w:r></w:p>`))
    expect(p0(doc).style.backgroundColor).toBe('FFFF00')
  })

  it('a "false" toggle disables the flag (w:vanish w:val="0")', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>shown</w:t></w:r></w:p>`))
    expect(runs0(doc)[0].style.hidden).toBe(false)
  })
})
