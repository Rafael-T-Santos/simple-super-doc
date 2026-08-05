import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, ImageRun } from '../src/types.js'

const PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

// rels are injected after the styles relationship; media files added to word/media.
async function buildDocx(documentBody: string, rels: string, withMedia = true): Promise<ArrayBuffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
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
    `${rels}</Relationships>`)
  if (withMedia) zip.file('word/media/image1.png', PNG_B64, { base64: true })
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:o" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const firstImage = (doc: { blocks: unknown[] }): ImageRun | undefined =>
  (doc.blocks[0] as ParagraphBlock).runs.find(r => r.type === 'image') as ImageRun | undefined

describe('images', () => {
  it('renders a VML image (w:pict > v:imagedata) with size from the shape style', async () => {
    const body =
      `<w:p><w:r><w:pict><v:shape id="Picture 1" type="#_x0000_t75" ` +
      `style="width:150pt;height:75pt"><v:imagedata r:id="rId9" o:title=""/></v:shape></w:pict></w:r></w:p>`
    const doc = await parse(await buildDocx(body, `<Relationship Id="rId9" Type="${IMAGE_REL}" Target="media/image1.png"/>`))
    const img = firstImage(doc)
    expect(img?.src.startsWith('data:image/png;base64,')).toBe(true)
    // 150pt -> 200px, 75pt -> 100px (96/72)
    expect(img?.widthPx).toBe(200)
    expect(img?.heightPx).toBe(100)
  })

  it('renders an external (linked) image as its URL', async () => {
    const body =
      `<w:p><w:r><w:drawing><wp:inline xmlns:wp="x" xmlns:a="x"><wp:extent cx="952500" cy="952500"/>` +
      `<a:graphic><a:graphicData><pic:pic xmlns:pic="x"><pic:blipFill>` +
      `<a:blip r:link="rId8"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    const doc = await parse(await buildDocx(
      body,
      `<Relationship Id="rId8" Type="${IMAGE_REL}" Target="https://example.com/photo.png" TargetMode="External"/>`,
      false,
    ))
    const img = firstImage(doc)
    expect(img?.src).toBe('https://example.com/photo.png')
  })

  it('a VML text box (pict with no imagedata) yields no image', async () => {
    const body =
      `<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>` +
      `<w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`
    const doc = await parse(await buildDocx(body, ''))
    // no image run; the text is recovered into the flow instead
    const text = doc.blocks.map(b => b.type === 'paragraph'
      ? (b as ParagraphBlock).runs.map(r => (r.type === 'run' ? r.text : '')).join('') : '').join('')
    expect(firstImage(doc)).toBeUndefined()
    expect(text).toContain('boxed')
  })
})

// A run can carry an image alongside its text. Word and Google Docs both emit
// "<w:t>…</w:t><w:drawing/><w:t>…</w:t>" inside a single <w:r>. The image used to
// stand for the WHOLE run, so any text sharing it was silently dropped.
describe('image sharing a run with text', () => {
  const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  const REL = `<Relationship Id="rId9" Type="${IMAGE_REL}" Target="media/image1.png"/>`
  const DRAWING =
    `<w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData>` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>` +
    `<a:blip r:embed="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`

  const shape = (doc: { blocks: unknown[] }): string[] =>
    (doc.blocks[0] as ParagraphBlock).runs.map(r => (r.type === 'image' ? '[img]' : r.text))

  it('keeps the text on both sides of an inline image', async () => {
    const body = `<w:p><w:r><w:t>ANTES</w:t>${DRAWING}<w:t>DEPOIS</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body, REL))
    expect(shape(doc)).toEqual(['ANTES', '[img]', 'DEPOIS'])
  })

  it('keeps text that precedes an image in the same run', async () => {
    const body = `<w:p><w:r><w:t>only text</w:t>${DRAWING}</w:r></w:p>`
    const doc = await parse(await buildDocx(body, REL))
    expect(shape(doc)).toEqual(['only text', '[img]'])
  })

  it('orders an image against a line break in the same run', async () => {
    const body = `<w:p><w:r><w:t>A</w:t><w:br/>${DRAWING}<w:t>B</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body, REL))
    const runs = (doc.blocks[0] as ParagraphBlock).runs
    expect(shape(doc)).toEqual(['A', '[img]', 'B'])
    // The break belongs to 'A' — it renders after that run, before the image.
    expect(runs[0].type === 'run' && runs[0].lineBreak).toBe(true)
  })

  it('GUARD an image-only run still yields just the image', async () => {
    const body = `<w:p><w:r>${DRAWING}</w:r></w:p>`
    const doc = await parse(await buildDocx(body, REL))
    expect(shape(doc)).toEqual(['[img]'])
    expect((firstImage(doc) as ImageRun).widthPx).toBe(96)
  })

  it('advances through several images in one run', async () => {
    const body =
      `<w:p><w:r><w:t>a</w:t>${DRAWING}<w:t>b</w:t>${DRAWING}<w:t>c</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body, REL))
    expect(shape(doc)).toEqual(['a', '[img]', 'b', '[img]', 'c'])
  })

  it('GUARD an unresolvable image drops the run instead of emitting an empty one', async () => {
    // rId9 is not declared, so the drawing resolves to nothing.
    const body = `<w:p><w:r>${DRAWING}</w:r><w:r><w:t>after</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body, ''))
    expect(shape(doc)).toEqual(['after'])
  })
})
