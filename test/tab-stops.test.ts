import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock } from '../src/types.js'

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

describe('tab stops (w:tabs)', () => {
  it('parses a right tab stop with a dot leader (TOC row)', async () => {
    const body =
      `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr>` +
      `<w:r><w:t>Chapter One</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.style.tabStops).toBeDefined()
    expect(p.style.tabStops!.length).toBe(1)
    const stop = p.style.tabStops![0]
    expect(stop.val).toBe('right')
    expect(stop.leader).toBe('dot')
    // 9350 twips * 96 / 1440 = 623px (rounded)
    expect(stop.posPx).toBe(623)
    // the tab run carries the tab count consumed by the layout
    const tabRun = p.runs.find(r => (r as any).tabs)
    expect(tabRun).toBeDefined()
  })

  it('sorts multiple tab stops by position and maps alignments', async () => {
    const body =
      `<w:p><w:pPr><w:tabs>` +
      `<w:tab w:val="center" w:pos="4000"/>` +
      `<w:tab w:val="left" w:pos="1000"/>` +
      `<w:tab w:val="decimal" w:leader="hyphen" w:pos="8000"/>` +
      `</w:tabs></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const stops = (doc.blocks[0] as ParagraphBlock).style.tabStops!
    expect(stops.map(s => s.val)).toEqual(['left', 'center', 'decimal'])
    expect(stops[2].leader).toBe('hyphen')
  })

  it('ignores a "clear" tab stop', async () => {
    const body =
      `<w:p><w:pPr><w:tabs>` +
      `<w:tab w:val="clear" w:pos="500"/>` +
      `<w:tab w:val="right" w:pos="6000"/>` +
      `</w:tabs></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const stops = (doc.blocks[0] as ParagraphBlock).style.tabStops!
    expect(stops.length).toBe(1)
    expect(stops[0].val).toBe('right')
  })

  it('leaves a paragraph without tabs free of tabStops', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`))
    expect((doc.blocks[0] as ParagraphBlock).style.tabStops).toBeUndefined()
  })
})
