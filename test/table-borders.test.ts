import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { TableBlock, TableCell } from '../src/types.js'

// stylesBody is injected inside <w:styles> (after docDefaults) so tests can
// declare table styles (w:type="table") with their own w:tblBorders.
async function buildDocx(documentBody: string, stylesBody = ''): Promise<ArrayBuffer> {
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
    `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>${stylesBody}</w:styles>`)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

// A 2x2 table; tblPr is injected (style ref and/or direct borders), tcPr0 sets
// the first cell's own properties.
const table = (tblPr: string, tc0Pr = '') =>
  `<w:tbl><w:tblPr>${tblPr}</w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>` +
  `<w:tr>` +
  `<w:tc><w:tcPr>${tc0Pr}</w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:tcPr/><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>` +
  `</w:tr>` +
  `<w:tr>` +
  `<w:tc><w:tcPr/><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:tcPr/><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>` +
  `</w:tr></w:tbl>`

// A TableGrid table style (basedOn TableNormal) with single 0.5pt borders on all
// sides + inside — the classic implicit-border source.
const tableGridStyle =
  `<w:style w:type="table" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>` +
  `<w:style w:type="table" w:styleId="TableGrid"><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>` +
  `<w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>` +
  `<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>` +
  `<w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/>` +
  `</w:tblBorders></w:tblPr></w:style>`

const firstCell = (doc: { blocks: unknown[] }): TableCell =>
  (doc.blocks[0] as TableBlock).rows[0].cells[0]

describe('table cell borders', () => {
  it('recovers borders that come only from the table style (TableGrid)', async () => {
    const doc = await parse(await buildDocx(table(`<w:tblStyle w:val="TableGrid"/>`), tableGridStyle))
    const c = firstCell(doc)
    expect(c.border?.top).toBe('1px solid #000')
    expect(c.border?.left).toBe('1px solid #000')
    expect(c.border?.right).toBe('1px solid #000')
    expect(c.border?.bottom).toBe('1px solid #000')
  })

  it('recovers borders from direct w:tblBorders', async () => {
    const tblPr = `<w:tblBorders><w:insideH w:val="single" w:sz="8" w:color="FF0000"/>` +
      `<w:insideV w:val="single" w:sz="8" w:color="FF0000"/></w:tblBorders>`
    const doc = await parse(await buildDocx(table(tblPr)))
    const c = firstCell(doc)
    // sz=8 eighths = 1pt → round(1*96/72)=1px
    expect(c.border?.top).toBe('1px solid #FF0000')
  })

  it('lets a cell w:tcBorders override the table border', async () => {
    const tc0 = `<w:tcBorders><w:top w:val="double" w:sz="24" w:color="00AA00"/></w:tcBorders>`
    const doc = await parse(await buildDocx(table(`<w:tblStyle w:val="TableGrid"/>`, tc0), tableGridStyle))
    const c = firstCell(doc)
    // sz=24 eighths = 3pt → round(3*96/72)=4px, double
    expect(c.border?.top).toBe('4px double #00AA00')
    // other sides still inherit the uniform style border
    expect(c.border?.left).toBe('1px solid #000')
  })

  it('honors an explicit "none" cell border (suppresses the uniform border)', async () => {
    const tc0 = `<w:tcBorders><w:top w:val="nil"/></w:tcBorders>`
    const doc = await parse(await buildDocx(table(`<w:tblStyle w:val="TableGrid"/>`, tc0), tableGridStyle))
    const c = firstCell(doc)
    expect(c.border?.top).toBeUndefined()
    expect(c.border?.bottom).toBe('1px solid #000')
  })

  it('emits no border when the table defines none', async () => {
    const doc = await parse(await buildDocx(table(``)))
    expect(firstCell(doc).border).toBeUndefined()
  })
})
