import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

// A run's child order is recovered from its own raw <w:r> slice. If those slices
// ever misalign with the parsed runs, a run would be ordered by ANOTHER run's
// children — silent corruption rather than a crash. These tests pin the
// alignment at the places it could realistically drift.

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
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>` +
    `</Relationships>`)
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const p0 = (doc: { blocks: unknown[] }): ParagraphBlock => doc.blocks[0] as ParagraphBlock
const texts = (p: ParagraphBlock): string[] => p.runs.map(r => (r as TextRun).text)
const breaks = (p: ParagraphBlock): boolean[] => p.runs.map(r => !!(r as TextRun).lineBreak)

describe('run-to-XML alignment', () => {
  // A wrapper contributes SEVERAL runs from one order entry. Only the second run
  // here carries a break — if the slices shifted, it would land on the first.
  it('gives each run of a multi-run hyperlink its own child order', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:hyperlink r:id="rId9">` +
      `<w:r><w:t>one</w:t></w:r>` +
      `<w:r><w:t>two</w:t><w:br/><w:t>three</w:t></w:r>` +
      `</w:hyperlink></w:p>`))
    expect(texts(p0(doc))).toEqual(['one', 'two', 'three'])
    expect(breaks(p0(doc))).toEqual([false, true, false])
    expect(p0(doc).runs.every(r => (r as TextRun).href === 'https://example.com')).toBe(true)
  })

  // Run properties are stripped before scanning: <w:sz>/<w:rFonts> and friends
  // must not be mistaken for the run's own content children.
  it('ignores rPr children when recovering a run\'s child order', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/></w:rPr>` +
      `<w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p>`))
    expect(texts(p0(doc))).toEqual(['A', 'B'])
    expect(breaks(p0(doc))).toEqual([true, false])
  })

  // A text box nests whole paragraphs inside a run. Their <w:t> must not be
  // counted as the outer run's text — the scan declines and the run falls back.
  it('falls back safely for a run wrapping a text box', async () => {
    const doc = await parse(await buildDocx(
      `<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>` +
      `<w:p><w:r><w:t>INNER</w:t></w:r></w:p>` +
      `</w:txbxContent></v:textbox></v:shape></w:pict></w:r>` +
      `<w:r><w:t>OUTER</w:t></w:r></w:p>`))
    const all = JSON.stringify(doc.blocks)
    expect(all).toContain('OUTER')
    expect(all).toContain('INNER')
  })
})
