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
