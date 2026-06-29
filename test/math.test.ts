import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { Block, ParagraphBlock, TableBlock, TextRun } from '../src/types.js'

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
    `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${documentBody}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const allText = (blocks: Block[]): string =>
  blocks.map(b => b.type === 'paragraph'
    ? (b as ParagraphBlock).runs.map(r => (r.type === 'run' ? (r as TextRun).text : '')).join('')
    : '').join('\n')

// An OMML equation whose math runs (m:r/m:t) spell out `text`.
const oMath = (...parts: string[]) =>
  `<m:oMath>${parts.map(p => `<m:r><m:t>${p}</m:t></m:r>`).join('')}</m:oMath>`

describe('OMML math (m:oMath)', () => {
  it('recovers an inline equation as linear text in document order', async () => {
    const body =
      `<w:p><w:r><w:t xml:space="preserve">Area is </w:t></w:r>` +
      oMath('A=', 'π', 'r', '2') +
      `<w:r><w:t xml:space="preserve"> exactly</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    const p = doc.blocks[0] as ParagraphBlock
    expect(p.runs.map(r => (r as TextRun).text).join('')).toBe('Area is A=πr2 exactly')
  })

  it('recovers a display equation wrapped in m:oMathPara', async () => {
    const body =
      `<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>` +
      oMath('E=mc', '2') +
      `</m:oMathPara></w:p>`
    const doc = await parse(await buildDocx(body))
    expect(allText(doc.blocks)).toContain('E=mc2')
  })

  it('recovers an equation inside a table cell', async () => {
    const body =
      `<w:tbl><w:tr><w:tc><w:tcPr/><w:p><w:r><w:t xml:space="preserve">x = </w:t></w:r>` +
      oMath('x=', '-b', '±', 'b2-4ac', '2a') + `</w:p></w:tc></w:tr></w:tbl>`
    const doc = await parse(await buildDocx(body))
    const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock
    expect(allText(tbl.rows[0].cells[0].blocks)).toContain('x=-b±b2-4ac2a')
  })

  it('recovers multiple equations in one paragraph', async () => {
    const body = `<w:p>${oMath('a=b')}<w:r><w:t xml:space="preserve"> and </w:t></w:r>${oMath('c=d')}</w:p>`
    const doc = await parse(await buildDocx(body))
    expect(allText(doc.blocks)).toBe('a=b and c=d')
  })

  it('does not crash on an empty equation', async () => {
    const body = `<w:p><w:r><w:t>before</w:t></w:r><m:oMath></m:oMath><w:r><w:t>after</w:t></w:r></w:p>`
    const doc = await parse(await buildDocx(body))
    expect(allText(doc.blocks)).toBe('beforeafter')
  })
})
