import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { TableBlock, ParagraphBlock, TextRun } from '../src/types.js'

async function buildDocx(documentBody: string, rels = ''): Promise<ArrayBuffer> {
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
    rels + `</Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const cellTexts = (table: TableBlock, r: number, c: number): string[] => {
  const p = table.rows[r].cells[c].blocks[0] as ParagraphBlock
  return p.runs.map(run => (run as TextRun).text)
}

const linkRel = `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>`

describe('run order inside table cells', () => {
  it('keeps a mid-cell hyperlink in document order', async () => {
    const body =
      `<w:tbl><w:tr><w:tc><w:p>` +
      `<w:r><w:t xml:space="preserve">before </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>LINK</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> after</w:t></w:r>` +
      `</w:p></w:tc></w:tr></w:tbl>`
    const doc = await parse(await buildDocx(body, linkRel))
    const table = doc.blocks[0] as TableBlock
    expect(cellTexts(table, 0, 0)).toEqual(['before ', 'LINK', ' after'])
    const linkRun = (table.rows[0].cells[0].blocks[0] as ParagraphBlock).runs[1] as TextRun
    expect(linkRun.href).toBe('https://example.com')
  })

  it('keeps order across a two-cell row independently', async () => {
    const body =
      `<w:tbl><w:tr>` +
      `<w:tc><w:p><w:r><w:t xml:space="preserve">a </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>L1</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:hyperlink r:id="rId9"><w:r><w:t>L2</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`
    const doc = await parse(await buildDocx(body, linkRel))
    const table = doc.blocks[0] as TableBlock
    expect(cellTexts(table, 0, 0)).toEqual(['a ', 'L1', ' b'])
    expect(cellTexts(table, 0, 1)).toEqual(['L2', ' tail'])
  })

  it('still recovers order in a body paragraph after a table', async () => {
    const body =
      `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      `<w:p><w:r><w:t xml:space="preserve">x </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>Y</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> z</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body, linkRel))
    const para = doc.blocks[1] as ParagraphBlock
    expect(para.runs.map(r => (r as TextRun).text)).toEqual(['x ', 'Y', ' z'])
  })

  it('recovers order in a nested-table cell one level deep', async () => {
    const body =
      `<w:tbl><w:tr><w:tc>` +
      `<w:tbl><w:tr><w:tc><w:p>` +
      `<w:r><w:t xml:space="preserve">n </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>IN</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> end</w:t></w:r>` +
      `</w:p></w:tc></w:tr></w:tbl>` +
      `</w:tc></w:tr></w:tbl>`
    const doc = await parse(await buildDocx(body, linkRel))
    const outer = doc.blocks[0] as TableBlock
    const inner = outer.rows[0].cells[0].blocks[0] as TableBlock
    const p = inner.rows[0].cells[0].blocks[0] as ParagraphBlock
    expect(p.runs.map(r => (r as TextRun).text)).toEqual(['n ', 'IN', ' end'])
  })
})
