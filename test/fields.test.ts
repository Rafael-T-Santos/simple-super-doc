import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

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

// A complex field: begin, code, separate, cached result, end.
const field = (code: string, cached = '') =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  (cached ? `<w:r><w:t>${cached}</w:t></w:r>` : '') +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`

describe('field codes', () => {
  it('PAGE with a \\* MERGEFORMAT switch still resolves to a live page number', async () => {
    const doc = await parse(await buildDocx(`<w:p>${field('PAGE \\* MERGEFORMAT', '3')}</w:p>`))
    const p = doc.blocks[0] as ParagraphBlock
    // exactly one run: the live page-number marker — NOT the cached "3"
    expect(p.runs.length).toBe(1)
    expect((p.runs[0] as TextRun).pageNumber).toBe(true)
    expect(p.runs.some(r => (r as TextRun).text === '3')).toBe(false)
  })

  it('NUMPAGES resolves to a total-pages marker and suppresses the cached count', async () => {
    const doc = await parse(await buildDocx(`<w:p>Page ${field('NUMPAGES', '5')}</w:p>`))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.runs.some(r => (r as TextRun).totalPages)).toBe(true)
    expect(p.runs.some(r => (r as TextRun).text === '5')).toBe(false)
  })

  it('a "Page X of Y" footer pattern yields page + total markers', async () => {
    const body =
      `<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>` +
      field('PAGE', '1') +
      `<w:r><w:t xml:space="preserve"> of </w:t></w:r>` +
      field('NUMPAGES', '2') +
      `</w:p>`
    const doc = await parse(await buildDocx(body))
    const runs = (doc.blocks[0] as ParagraphBlock).runs as TextRun[]
    expect(runs.some(r => r.pageNumber)).toBe(true)
    expect(runs.some(r => r.totalPages)).toBe(true)
    expect(runs.map(r => r.text).join('')).toContain('Page ')
    expect(runs.map(r => r.text).join('')).toContain(' of ')
  })

  it('a non-live field (DATE) renders its cached result as text', async () => {
    const doc = await parse(await buildDocx(`<w:p>${field('DATE \\@ "yyyy-MM-dd"', '2026-06-29')}</w:p>`))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.runs.some(r => (r as TextRun).text === '2026-06-29')).toBe(true)
    // the field code itself never renders
    expect(p.runs.some(r => (r as TextRun).text.includes('DATE'))).toBe(false)
  })

  it('a REF field renders its cached cross-reference text', async () => {
    const doc = await parse(await buildDocx(`<w:p>${field('REF _Ref123 \\h', 'Section 2')}</w:p>`))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.runs.some(r => (r as TextRun).text === 'Section 2')).toBe(true)
  })
})
