import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { Block, ParagraphBlock, TextRun } from '../src/types.js'

const hdr = (text: string) =>
  `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:hdr>`

const textOf = (blocks: Block[] | undefined): string =>
  (blocks ?? []).map(b => b.type === 'paragraph'
    ? (b as ParagraphBlock).runs.map(r => (r.type === 'run' ? (r as TextRun).text : '')).join('') : '').join('')

// Builder with header parts, their relationships, a body sectPr, and optional settings.xml.
async function buildDocx(opts: {
  sectPr: string
  rels: string
  parts: Record<string, string>
  settings?: string
}): Promise<ArrayBuffer> {
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
    `${opts.rels}</Relationships>`)
  for (const [name, content] of Object.entries(opts.parts)) zip.file(`word/${name}`, content)
  if (opts.settings) zip.file('word/settings.xml', opts.settings)
  const body =
    `<w:p><w:r><w:t>body</w:t></w:r></w:p>` +
    `<w:sectPr>${opts.sectPr}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const HREF = (type: string, rId: string) => `<w:headerReference w:type="${type}" r:id="${rId}"/>`
const HREL = (rId: string, target: string) =>
  `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${target}"/>`

describe('per-page headers/footers', () => {
  it('resolves a distinct first-page header and sets titlePg', async () => {
    const doc = await parse(await buildDocx({
      sectPr: `<w:titlePg/>${HREF('default', 'rId10')}${HREF('first', 'rId11')}`,
      rels: HREL('rId10', 'header1.xml') + HREL('rId11', 'header2.xml'),
      parts: { 'header1.xml': hdr('DEFAULT HEADER'), 'header2.xml': hdr('FIRST PAGE HEADER') },
    }))
    expect(doc.titlePg).toBe(true)
    expect(textOf(doc.header)).toBe('DEFAULT HEADER')
    expect(textOf(doc.headerFirst)).toBe('FIRST PAGE HEADER')
  })

  it('resolves a distinct even-page header when settings enable evenAndOddHeaders', async () => {
    const doc = await parse(await buildDocx({
      sectPr: `${HREF('default', 'rId10')}${HREF('even', 'rId12')}`,
      rels: HREL('rId10', 'header1.xml') + HREL('rId12', 'header3.xml'),
      parts: { 'header1.xml': hdr('ODD/DEFAULT'), 'header3.xml': hdr('EVEN HEADER') },
      settings: `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/></w:settings>`,
    }))
    expect(doc.evenAndOddHeaders).toBe(true)
    expect(textOf(doc.headerEven)).toBe('EVEN HEADER')
  })

  it('does not set titlePg or evenAndOddHeaders when absent', async () => {
    const doc = await parse(await buildDocx({
      sectPr: HREF('default', 'rId10'),
      rels: HREL('rId10', 'header1.xml'),
      parts: { 'header1.xml': hdr('ONLY DEFAULT') },
    }))
    expect(doc.titlePg).toBeUndefined()
    expect(doc.evenAndOddHeaders).toBeUndefined()
    expect(doc.headerFirst).toBeUndefined()
    expect(doc.headerEven).toBeUndefined()
  })
})
