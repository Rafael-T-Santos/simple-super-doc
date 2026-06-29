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

const ins = (text: string) =>
  `<w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`
const del = (text: string) =>
  `<w:del w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`

const texts = (p: ParagraphBlock): string[] => p.runs.map(r => (r as TextRun).text)

describe('tracked changes (parse)', () => {
  it('captures a deleted run from w:delText and flags it', async () => {
    const body = `<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r>${del('gone')}<w:r><w:t xml:space="preserve"> stay</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['keep ', 'gone', ' stay'])
    const deleted = p.runs.find(r => (r as TextRun).text === 'gone') as TextRun
    expect(deleted.deleted).toBe(true)
  })

  it('flags an inserted run', async () => {
    const body = `<w:p><w:r><w:t xml:space="preserve">a </w:t></w:r>${ins('NEW')}<w:r><w:t xml:space="preserve"> b</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['a ', 'NEW', ' b'])
    const inserted = p.runs.find(r => (r as TextRun).text === 'NEW') as TextRun
    expect(inserted.inserted).toBe(true)
  })

  it('keeps deletion and insertion interleaved in document order', async () => {
    const body =
      `<w:p><w:r><w:t>1</w:t></w:r>${del('2')}${ins('3')}<w:r><w:t>4</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['1', '2', '3', '4'])
    expect((p.runs[1] as TextRun).deleted).toBe(true)
    expect((p.runs[2] as TextRun).inserted).toBe(true)
  })
})
