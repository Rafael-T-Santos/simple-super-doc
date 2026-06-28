import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

// Build a minimal .docx in memory with a paragraph that mixes a direct run, a
// hyperlink, and a tracked insertion in NON-grouped document order, to prove the
// parser recovers run order instead of appending each group.
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

const texts = (p: ParagraphBlock): string[] =>
  p.runs.map(r => (r as TextRun).text)

describe('intra-paragraph run order', () => {
  it('keeps a hyperlink in the middle of a paragraph in document order', async () => {
    const body =
      `<w:p>` +
      `<w:r><w:t xml:space="preserve">before </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>LINK</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> after</w:t></w:r>` +
      `</w:p>`
    const rel = `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>`
    const doc = await parse(await buildDocx(body, rel))
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['before ', 'LINK', ' after'])
    expect((p.runs[1] as TextRun).href).toBe('https://example.com')
  })

  it('keeps a tracked insertion in the middle in document order', async () => {
    const body =
      `<w:p>` +
      `<w:r><w:t xml:space="preserve">a </w:t></w:r>` +
      `<w:ins w:id="1" w:author="t" w:date="2026-01-01T00:00:00Z"><w:r><w:t>INS</w:t></w:r></w:ins>` +
      `<w:r><w:t xml:space="preserve"> b</w:t></w:r>` +
      `</w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['a ', 'INS', ' b'])
  })

  it('handles multiple interleaved hyperlinks', async () => {
    const body =
      `<w:p>` +
      `<w:hyperlink r:id="rId9"><w:r><w:t>one</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> mid </w:t></w:r>` +
      `<w:hyperlink r:id="rId10"><w:r><w:t>two</w:t></w:r></w:hyperlink>` +
      `</w:p>`
    const rels =
      `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://a.test" TargetMode="External"/>` +
      `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://b.test" TargetMode="External"/>`
    const doc = await parse(await buildDocx(body, rels))
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['one', ' mid ', 'two'])
    expect((p.runs[0] as TextRun).href).toBe('https://a.test')
    expect((p.runs[2] as TextRun).href).toBe('https://b.test')
  })
})
