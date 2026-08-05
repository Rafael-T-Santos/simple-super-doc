import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'
import type { ParagraphBlock, TextRun } from '../src/types.js'

// Document order has to be recovered from raw XML (fast-xml-parser groups
// children by tag). parseDocument always passed that raw XML down; the footer
// and notes entry points did not, so run order silently degraded there. These
// tests cover the parts OTHER than the body: footer and footnotes/endnotes.

async function buildDocx(opts: {
  footerXml?: string
  footnotesXml?: string
  bodyXml?: string
}): Promise<ArrayBuffer> {
  const zip = new JSZip()
  const overrides: string[] = []
  const rels = [`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`]
  const refs: string[] = []

  zip.file('_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/styles.xml',
    `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults></w:styles>`)

  if (opts.footerXml) {
    zip.file('word/footer1.xml',
      `<?xml version="1.0"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${opts.footerXml}</w:ftr>`)
    zip.file('word/_rels/footer1.xml.rels',
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>`)
    rels.push(`<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`)
    refs.push(`<w:footerReference w:type="default" r:id="rId11"/>`)
    overrides.push(`<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
  }
  if (opts.footnotesXml) {
    zip.file('word/footnotes.xml',
      `<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${opts.footnotesXml}</w:footnotes>`)
    rels.push(`<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>`)
    overrides.push(`<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`)
  }

  zip.file('[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    overrides.join('') + `</Types>`)
  zip.file('word/_rels/document.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    rels.join('') + `</Relationships>`)

  const sectPr = `<w:sectPr>${refs.join('')}<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`
  const body = opts.bodyXml ?? `<w:p><w:r><w:t>Body</w:t></w:r></w:p>`
  zip.file('word/document.xml',
    `<?xml version="1.0"?><w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}${sectPr}</w:body></w:document>`)
  return zip.generateAsync({ type: 'arraybuffer' })
}

const texts = (p: ParagraphBlock): string[] => p.runs.map(r => (r as TextRun).text)
const breaks = (p: ParagraphBlock): boolean[] => p.runs.map(r => !!(r as TextRun).lineBreak)

describe('document order in a footer part', () => {
  it('puts a break between the two texts of a footer run', async () => {
    const doc = await parse(await buildDocx({
      footerXml: `<w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p>`,
    }))
    const p = doc.footer![0] as ParagraphBlock
    expect(texts(p)).toEqual(['A', 'B'])
    expect(breaks(p)).toEqual([true, false])
  })

  it('keeps a hyperlink in the MIDDLE of a footer paragraph in place', async () => {
    const doc = await parse(await buildDocx({
      footerXml:
        `<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r>` +
        `<w:hyperlink r:id="rId9"><w:r><w:t>LINK</w:t></w:r></w:hyperlink>` +
        `<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`,
    }))
    const p = doc.footer![0] as ParagraphBlock
    expect(texts(p)).toEqual(['before ', 'LINK', ' after'])
    expect((p.runs[1] as TextRun).href).toBe('https://example.com')
  })

  it('keeps the "V.4 <tabs> CONFIDENCIAL" footer split intact', async () => {
    const doc = await parse(await buildDocx({
      footerXml: `<w:p><w:r><w:t>V.4</w:t><w:tab/><w:tab/><w:t>CONFIDENCIAL</w:t></w:r></w:p>`,
    }))
    const p = doc.footer![0] as ParagraphBlock
    expect(texts(p)).toEqual(['V.4', 'CONFIDENCIAL'])
    expect((p.runs[1] as TextRun).tabs).toBe(2)
  })
})

describe('document order inside footnotes', () => {
  // The separator/continuation pseudo-notes come first in a real footnotes.xml.
  // They are skipped when building the map, so the raw-XML slices must stay
  // aligned with the PARSED notes by position or a note gets another note's XML.
  const withSeparators = (notes: string): string =>
    `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
    notes

  const noteRef = (id: string): string =>
    `<w:p><w:r><w:t>text</w:t></w:r><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`

  it('puts a break between the two texts of a footnote run', async () => {
    const doc = await parse(await buildDocx({
      bodyXml: noteRef('2'),
      footnotesXml: withSeparators(
        `<w:footnote w:id="2"><w:p><w:r><w:t>A</w:t><w:br/><w:t>B</w:t></w:r></w:p></w:footnote>`),
    }))
    const p = doc.footnotes![0].blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['A', 'B'])
    expect(breaks(p)).toEqual([true, false])
  })

  it('aligns each footnote with its OWN xml when several are present', async () => {
    const doc = await parse(await buildDocx({
      bodyXml:
        `<w:p><w:r><w:t>x</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r>` +
        `<w:r><w:t>y</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>`,
      footnotesXml: withSeparators(
        `<w:footnote w:id="2"><w:p><w:r><w:t>ONE</w:t><w:br/><w:t>1</w:t></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="3"><w:p><w:r><w:t>TWO</w:t><w:tab/><w:t>2</w:t></w:r></w:p></w:footnote>`),
    }))
    const first = doc.footnotes![0].blocks[0] as ParagraphBlock
    const second = doc.footnotes![1].blocks[0] as ParagraphBlock
    expect(texts(first)).toEqual(['ONE', '1'])
    expect(breaks(first)).toEqual([true, false])
    expect(texts(second)).toEqual(['TWO', '2'])
    // The tab belongs to the SECOND note, proving the slices didn't shift.
    expect((second.runs[1] as TextRun).tabs).toBe(1)
    expect(breaks(second)).toEqual([false, false])
  })

  it('does not lose footnote text sharing a run with a page break', async () => {
    const doc = await parse(await buildDocx({
      bodyXml: noteRef('2'),
      footnotesXml: withSeparators(
        `<w:footnote w:id="2"><w:p><w:r><w:t>ANTES</w:t><w:br w:type="page"/><w:t>DEPOIS</w:t></w:r></w:p></w:footnote>`),
    }))
    const blocks = doc.footnotes![0].blocks
    expect(texts(blocks[0] as ParagraphBlock)).toEqual(['ANTES'])
    expect(texts(blocks[1] as ParagraphBlock)).toEqual(['DEPOIS'])
  })
})

describe('tracked deletions keep break order', () => {
  it('places a break between two w:delText of one run', async () => {
    const doc = await parse(await buildDocx({
      bodyXml:
        `<w:p><w:del w:id="1" w:author="A"><w:r><w:delText>A</w:delText><w:br/>` +
        `<w:delText>B</w:delText></w:r></w:del></w:p>`,
    }))
    // Deleted runs are hidden in the default (accepted) view; parse them via the
    // IR, which keeps them flagged rather than dropped.
    const p = doc.blocks[0] as ParagraphBlock
    expect(texts(p)).toEqual(['A', 'B'])
    expect(breaks(p)).toEqual([true, false])
    expect(p.runs.every(r => (r as TextRun).deleted)).toBe(true)
  })
})
