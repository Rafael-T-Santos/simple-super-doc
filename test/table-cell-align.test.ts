import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { TableBlock, TableCell } from '../src/types.js'

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

// A 1x2 table whose FIRST cell carries the properties under test.
const table = (tc0Pr: string) =>
  `<w:tbl><w:tblPr/>` +
  `<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>` +
  `<w:tr>` +
  `<w:tc><w:tcPr>${tc0Pr}</w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>` +
  `<w:tc><w:tcPr/><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>` +
  `</w:tr></w:tbl>`

const firstCell = async (tc0Pr: string): Promise<TableCell> => {
  const doc = await parse(await buildDocx(table(tc0Pr)))
  return (doc.blocks[0] as TableBlock).rows[0].cells[0]
}

describe('cell vertical alignment (w:vAlign)', () => {
  it('reads center', async () => {
    expect((await firstCell(`<w:vAlign w:val="center"/>`)).verticalAlign).toBe('center')
  })

  it('reads bottom', async () => {
    expect((await firstCell(`<w:vAlign w:val="bottom"/>`)).verticalAlign).toBe('bottom')
  })

  // "both" is vertical justification. CSS has no equivalent and Word draws a
  // single line of it centered, so it lands on center rather than being dropped.
  it('maps "both" to center', async () => {
    expect((await firstCell(`<w:vAlign w:val="both"/>`)).verticalAlign).toBe('center')
  })

  // Top is the default; leaving it unset keeps the IR free of no-op properties.
  it('leaves an explicit top unset', async () => {
    expect((await firstCell(`<w:vAlign w:val="top"/>`)).verticalAlign).toBeUndefined()
  })

  it('leaves a cell with no w:vAlign unset', async () => {
    expect((await firstCell(``)).verticalAlign).toBeUndefined()
  })

  it('does not leak onto sibling cells', async () => {
    const doc = await parse(await buildDocx(table(`<w:vAlign w:val="bottom"/>`)))
    expect((doc.blocks[0] as TableBlock).rows[0].cells[1].verticalAlign).toBeUndefined()
  })

  it('coexists with shading and spans', async () => {
    const c = await firstCell(
      `<w:gridSpan w:val="2"/><w:shd w:fill="FF6109"/><w:vAlign w:val="center"/>`,
    )
    expect(c.verticalAlign).toBe('center')
    expect(c.colSpan).toBe(2)
    expect(c.backgroundColor).toBe('FF6109')
  })
})

describe('cell text direction (w:textDirection)', () => {
  it('reads tbRl (rotate 90°, reads top-to-bottom)', async () => {
    expect((await firstCell(`<w:textDirection w:val="tbRl"/>`)).textDirection).toBe('tbRl')
  })

  it('reads btLr (rotate 270°, reads bottom-to-top)', async () => {
    expect((await firstCell(`<w:textDirection w:val="btLr"/>`)).textDirection).toBe('btLr')
  })

  it('leaves the horizontal default (lrTb) unset', async () => {
    expect((await firstCell(`<w:textDirection w:val="lrTb"/>`)).textDirection).toBeUndefined()
  })

  // The *V variants are vertical CJK layout, not the cell rotation Word's UI
  // produces. Rendering them rotated would be wrong, so they stay horizontal.
  it('leaves the vertical CJK variants unset', async () => {
    expect((await firstCell(`<w:textDirection w:val="tbRlV"/>`)).textDirection).toBeUndefined()
    expect((await firstCell(`<w:textDirection w:val="lrTbV"/>`)).textDirection).toBeUndefined()
  })

  it('leaves a cell with no w:textDirection unset', async () => {
    expect((await firstCell(``)).textDirection).toBeUndefined()
  })

  it('keeps the cell content when the cell is rotated', async () => {
    const c = await firstCell(`<w:textDirection w:val="btLr"/><w:vAlign w:val="center"/>`)
    expect(c.textDirection).toBe('btLr')
    expect(c.verticalAlign).toBe('center')
    expect(c.blocks).toHaveLength(1)
  })
})
