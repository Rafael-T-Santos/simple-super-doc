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
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="x" xmlns:a="x" xmlns:wps="x" xmlns:v="x" xmlns:mc="x" xmlns:o="x">` +
    `<w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const allText = (blocks: Block[]): string =>
  blocks.map(b => b.type === 'paragraph'
    ? (b as ParagraphBlock).runs.map(r => (r.type === 'run' ? (r as TextRun).text : '')).join('')
    : '').join('\n')

// A DrawingML text box: <w:drawing> ... <wps:txbx><w:txbxContent>…</…>.
const drawingTxbx = (text: string) =>
  `<w:r><w:drawing><wp:anchor behindDoc="0"><a:graphic><a:graphicData>` +
  `<wps:wsp><wps:txbx><w:txbxContent>` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` +
  `</w:txbxContent></wps:txbx></wps:wsp>` +
  `</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`

// A VML text box: <w:pict><v:shape><v:textbox><w:txbxContent>…
const vmlTxbx = (text: string) =>
  `<w:r><w:pict><v:shape><v:textbox><w:txbxContent>` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` +
  `</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`

describe('text boxes (w:txbxContent)', () => {
  it('recovers text from a DrawingML text box into the flow', async () => {
    const doc = await parse(await buildDocx(`<w:p>${drawingTxbx('Inside the drawing text box')}</w:p>`))
    expect(allText(doc.blocks)).toContain('Inside the drawing text box')
  })

  it('recovers text from a VML text box into the flow', async () => {
    const doc = await parse(await buildDocx(`<w:p>${vmlTxbx('Inside the VML text box')}</w:p>`))
    expect(allText(doc.blocks)).toContain('Inside the VML text box')
  })

  it('does not duplicate a text box wrapped in mc:AlternateContent', async () => {
    // Word emits the same text box as a DrawingML Choice and a VML Fallback.
    const body =
      `<w:p><w:r><mc:AlternateContent>` +
      `<mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData>` +
      `<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>BOX TEXT</w:t></w:r></w:p>` +
      `</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
      `<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>BOX TEXT</w:t></w:r></w:p>` +
      `</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
      `</mc:AlternateContent></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const occurrences = allText(doc.blocks).split('BOX TEXT').length - 1
    expect(occurrences).toBe(1)
  })

  it('emits text box content after its carrier paragraph', async () => {
    const body = `<w:p><w:r><w:t>Carrier paragraph</w:t></w:r>${drawingTxbx('Box content')}</w:p>`
    const doc = await parse(await buildDocx(body))
    const text = allText(doc.blocks)
    expect(text.indexOf('Carrier paragraph')).toBeLessThan(text.indexOf('Box content'))
  })

  it('handles multiple paragraphs inside one text box', async () => {
    const body =
      `<w:p><w:r><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent>` +
      `<w:p><w:r><w:t>Line one</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Line two</w:t></w:r></w:p>` +
      `</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const text = allText(doc.blocks)
    expect(text).toContain('Line one')
    expect(text).toContain('Line two')
  })

  it('leaves a plain paragraph (no text box) untouched', async () => {
    const doc = await parse(await buildDocx(`<w:p><w:r><w:t>Just text</w:t></w:r></w:p>`))
    expect(doc.blocks.length).toBe(1)
    expect(allText(doc.blocks)).toBe('Just text')
  })
})
