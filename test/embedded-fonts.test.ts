import { describe, it, expect } from 'vitest'
import JSZip from 'jszip'
import { parse } from '../src/index.js'

// A minimal but real TrueType header: the sfnt version 1.0 that fontMime sniffs.
const ttf = (): Uint8Array => {
  const b = new Uint8Array(64)
  b.set([0x00, 0x01, 0x00, 0x00, 0x00, 0x10, 0x01, 0x00], 0)
  for (let i = 8; i < b.length; i++) b[i] = i
  return b
}

// Word obfuscates an embedded font by XORing its first 32 bytes with the
// fontKey GUID's bytes, reversed, applied twice. Producing one here lets the
// de-obfuscation be tested as a round trip rather than against a golden blob.
function obfuscate(bytes: Uint8Array, guid: string): Uint8Array {
  const hex = guid.replace(/[^0-9a-fA-F]/g, '')
  const key = new Uint8Array(16)
  for (let i = 0; i < 16; i++) key[15 - i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16)
  const out = bytes.slice()
  for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= key[i % 16]
  return out
}

type FontPart = { file: string; bytes: Uint8Array; key: string; slot: string; rid: string }

async function buildDocx(family: string, parts: FontPart[]): Promise<ArrayBuffer> {
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
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>`)

  const embeds = parts
    .map(p => `<w:${p.slot} w:fontKey="${p.key}" r:id="${p.rid}" w:subsetted="0"/>`)
    .join('')
  zip.file('word/fontTable.xml',
    `<?xml version="1.0"?><w:fonts ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:font w:name="Arial"/>` +
    `<w:font w:name="${family}">${embeds}</w:font></w:fonts>`)
  zip.file('word/_rels/fontTable.xml.rels',
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    parts.map(p =>
      `<Relationship Id="${p.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/${p.file}"/>`,
    ).join('') +
    `</Relationships>`)
  for (const p of parts) zip.file(`word/fonts/${p.file}`, p.bytes)

  return zip.generateAsync({ type: 'arraybuffer' })
}

const NO_KEY = '{00000000-0000-0000-0000-000000000000}'
const REAL_KEY = '{1A2B3C4D-5E6F-7081-92A3-B4C5D6E7F809}'

const decodeSrc = (src: string): Uint8Array => {
  const b64 = src.slice(src.indexOf(',') + 1)
  const bin = Buffer.from(b64, 'base64')
  return new Uint8Array(bin)
}

describe('embedded fonts', () => {
  it('exposes each embed slot with its CSS weight and style', async () => {
    const doc = await parse(await buildDocx('DM Sans', [
      { file: 'r.ttf', bytes: ttf(), key: NO_KEY, slot: 'embedRegular', rid: 'rId1' },
      { file: 'b.ttf', bytes: ttf(), key: NO_KEY, slot: 'embedBold', rid: 'rId2' },
      { file: 'i.ttf', bytes: ttf(), key: NO_KEY, slot: 'embedItalic', rid: 'rId3' },
      { file: 'bi.ttf', bytes: ttf(), key: NO_KEY, slot: 'embedBoldItalic', rid: 'rId4' },
    ]))
    expect(doc.fonts).toHaveLength(4)
    expect(doc.fonts!.map(f => `${f.weight}/${f.style}`)).toEqual([
      'normal/normal', 'bold/normal', 'normal/italic', 'bold/italic',
    ])
    // The family is the name runs refer to, so it must match w:font w:name.
    expect(new Set(doc.fonts!.map(f => f.family))).toEqual(new Set(['DM Sans']))
    expect(doc.fonts![0].src.startsWith('data:font/ttf;base64,')).toBe(true)
  })

  it('de-obfuscates an .odttf back to the original bytes', async () => {
    const original = ttf()
    const doc = await parse(await buildDocx('Poppins', [
      { file: 'r.odttf', bytes: obfuscate(original, REAL_KEY), key: REAL_KEY, slot: 'embedRegular', rid: 'rId1' },
    ]))
    expect(doc.fonts).toHaveLength(1)
    // Byte-for-byte: an off-by-one in the key order would still decode the
    // sniffable header while corrupting the rest of the file.
    expect(Array.from(decodeSrc(doc.fonts![0].src))).toEqual(Array.from(original))
  })

  it('leaves a font with an all-zero key untouched', async () => {
    const original = ttf()
    const doc = await parse(await buildDocx('DM Sans', [
      { file: 'r.ttf', bytes: original, key: NO_KEY, slot: 'embedRegular', rid: 'rId1' },
    ]))
    expect(Array.from(decodeSrc(doc.fonts![0].src))).toEqual(Array.from(original))
  })

  it('skips a part that is not a font a browser can load', async () => {
    const junk = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]) // a zip, not a font
    const doc = await parse(await buildDocx('DM Sans', [
      { file: 'r.ttf', bytes: junk, key: NO_KEY, slot: 'embedRegular', rid: 'rId1' },
      { file: 'b.ttf', bytes: ttf(), key: NO_KEY, slot: 'embedBold', rid: 'rId2' },
    ]))
    // The unreadable one is dropped; the good one still comes through.
    expect(doc.fonts).toHaveLength(1)
    expect(doc.fonts![0].weight).toBe('bold')
  })

  it('omits fonts entirely when the document embeds none', async () => {
    const doc = await parse(await buildDocx('DM Sans', []))
    expect(doc.fonts).toBeUndefined()
  })
})
