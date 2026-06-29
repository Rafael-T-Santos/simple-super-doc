import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { Block, ParagraphBlock, TextRun } from '../src/types.js'

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

const allText = (blocks: Block[]): string =>
  blocks.map(b => b.type === 'paragraph'
    ? (b as ParagraphBlock).runs.map(r => (r.type === 'run' ? (r as TextRun).text : '')).join('')
    : '').join('\n')

describe('content controls (w:sdt)', () => {
  it('recovers block-level sdt content into the flow', async () => {
    const body =
      `<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtEndPr/><w:sdtContent>` +
      `<w:p><w:r><w:t>Inside the content control</w:t></w:r></w:p>` +
      `</w:sdtContent></w:sdt>`
    const doc = await parse(await buildDocx(body))
    expect(allText(doc.blocks)).toContain('Inside the content control')
  })

  it('recovers inline sdt content in document order', async () => {
    const body =
      `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>` +
      `<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent><w:r><w:t>inside</w:t></w:r></w:sdtContent></w:sdt>` +
      `<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.runs.map(r => (r as TextRun).text).join('')).toBe('Before inside after')
  })

  it('recovers nested sdt content (control inside a control)', async () => {
    const body =
      `<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>` +
      `<w:p><w:r><w:t>Outer one</w:t></w:r></w:p>` +
      `<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent>` +
      `<w:p><w:r><w:t>Inner two</w:t></w:r></w:p>` +
      `</w:sdtContent></w:sdt>` +
      `<w:p><w:r><w:t>Outer three</w:t></w:r></w:p>` +
      `</w:sdtContent></w:sdt>`
    const doc = await parse(await buildDocx(body))
    const text = allText(doc.blocks)
    expect(text).toContain('Outer one')
    expect(text).toContain('Inner two')
    expect(text).toContain('Outer three')
    // order preserved
    expect(text.indexOf('Outer one')).toBeLessThan(text.indexOf('Inner two'))
    expect(text.indexOf('Inner two')).toBeLessThan(text.indexOf('Outer three'))
  })

  it('does not leak sdt control properties (sdtPr) into the text', async () => {
    const body =
      `<w:sdt><w:sdtPr><w:alias w:val="CompanyName"/><w:tag w:val="company"/><w:id w:val="9"/>` +
      `<w:dropDownList><w:listItem w:displayText="Acme" w:value="acme"/></w:dropDownList></w:sdtPr>` +
      `<w:sdtContent><w:p><w:r><w:t>Acme Corp</w:t></w:r></w:p></w:sdtContent></w:sdt>`
    const doc = await parse(await buildDocx(body))
    const text = allText(doc.blocks)
    expect(text).toContain('Acme Corp')
    expect(text).not.toContain('CompanyName')
    expect(text).not.toContain('company')
  })

  it('recovers an sdt that wraps a table', async () => {
    const body =
      `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>` +
      `<w:tbl><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      `</w:sdtContent></w:sdt>`
    const doc = await parse(await buildDocx(body))
    expect(doc.blocks.some(b => b.type === 'table')).toBe(true)
    expect(allText((doc.blocks.find(b => b.type === 'table') as { rows: { cells: { blocks: Block[] }[] }[] }).rows[0].cells[0].blocks)).toContain('Cell text')
  })
})
