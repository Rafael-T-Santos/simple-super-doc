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
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

describe('bidi / RTL and complex scripts', () => {
  it('flags a w:bidi paragraph as right-to-left', async () => {
    const body = `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>مرحبا بالعالم</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.style.rtl).toBe(true)
    expect((p.runs[0] as TextRun).text).toBe('مرحبا بالعالم')
  })

  it('flags a w:rtl run as right-to-left', async () => {
    const body = `<w:p><w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום עולם</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const run = (doc.blocks[0] as ParagraphBlock).runs[0] as TextRun
    expect(run.style.rtl).toBe(true)
    expect(run.text).toBe('שלום עולם')
  })

  it('does not flag rtl when w:bidi is explicitly disabled (val=0)', async () => {
    const body = `<w:p><w:pPr><w:bidi w:val="0"/></w:pPr><w:r><w:t>hello</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    expect((doc.blocks[0] as ParagraphBlock).style.rtl).toBeFalsy()
  })

  it('preserves CJK text (Chinese, Japanese) unchanged', async () => {
    const body =
      `<w:p><w:r><w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>你好世界</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>こんにちは世界</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    expect((doc.blocks[0] as ParagraphBlock).runs[0] as TextRun).toMatchObject({ text: '你好世界' })
    expect((doc.blocks[1] as ParagraphBlock).runs[0] as TextRun).toMatchObject({ text: 'こんにちは世界' })
  })
})
