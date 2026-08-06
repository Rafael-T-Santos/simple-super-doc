import type JSZip from 'jszip'
import type { RelationshipMap } from './relationships.js'
import { XMLParser } from 'fast-xml-parser'

// A font the document carries inside itself, ready to be registered with the
// browser. Without these a document whose font is not installed renders in a
// fallback, which breaks line and word count — and with them, pagination.
export type EmbeddedFont = {
  family: string // the name runs refer to, e.g. "DM Sans"
  weight: 'normal' | 'bold'
  style: 'normal' | 'italic'
  src: string // data: URL, usable directly as a CSS @font-face src
}

const parser = new XMLParser({
  removeNSPrefix: true,
  attributeNamePrefix: '',
  ignoreAttributes: false,
  parseAttributeValue: false,
  isArray: name => name === 'font',
})

// The four embed slots a <w:font> can declare, and the CSS they map to.
const SLOTS: Array<{ tag: string; weight: 'normal' | 'bold'; style: 'normal' | 'italic' }> = [
  { tag: 'embedRegular', weight: 'normal', style: 'normal' },
  { tag: 'embedBold', weight: 'bold', style: 'normal' },
  { tag: 'embedItalic', weight: 'normal', style: 'italic' },
  { tag: 'embedBoldItalic', weight: 'bold', style: 'italic' },
]

function uint8ArrayToBase64(bytes: Uint8Array): string {
  let b64 = ''
  for (let i = 0; i < bytes.length; i += 8192) {
    b64 += String.fromCharCode(...bytes.subarray(i, i + 8192))
  }
  return btoa(b64)
}

// Word ships embedded fonts obfuscated (.odttf): the first 32 bytes are XORed
// with the 16-byte key from w:fontKey, applied twice. The key is the GUID's hex
// digits read as bytes and REVERSED (ECMA-376 §17.8.1). Producers that do not
// obfuscate write an all-zero key, which makes this a no-op anyway.
function deobfuscate(bytes: Uint8Array, fontKey: string | undefined): Uint8Array {
  const hex = (fontKey ?? '').replace(/[^0-9a-fA-F]/g, '')
  if (hex.length !== 32) return bytes
  const key = new Uint8Array(16)
  for (let i = 0; i < 16; i++) key[15 - i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16)
  const out = bytes.slice()
  for (let i = 0; i < 32 && i < out.length; i++) out[i] ^= key[i % 16]
  return out
}

// Read the format from the file itself rather than the extension: a .odttf can
// hold either flavour, and the extension is gone once it is de-obfuscated.
function fontMime(bytes: Uint8Array): string | null {
  if (bytes.length < 4) return null
  const tag = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3])
  if (tag === 'OTTO') return 'font/otf'
  if (tag === 'ttcf') return 'font/collection'
  if (tag === 'true' || tag === 'typ1') return 'font/ttf'
  // TrueType's sfnt version is the fixed-point number 1.0.
  if (bytes[0] === 0x00 && bytes[1] === 0x01 && bytes[2] === 0x00 && bytes[3] === 0x00) return 'font/ttf'
  return null
}

// Every font embedded in the package, as data URLs. Returns [] when the document
// embeds none, which is the common case for documents written with fonts the
// reader is expected to already have.
export async function parseEmbeddedFonts(
  fontTableXml: string,
  rels: RelationshipMap,
  zip: JSZip,
): Promise<EmbeddedFont[]> {
  let doc: Record<string, unknown>
  try {
    doc = parser.parse(fontTableXml)
  } catch {
    return [] // a malformed font table must never fail the whole document
  }
  const entries = ((doc?.fonts as Record<string, unknown>)?.font ?? []) as Record<string, unknown>[]
  const out: EmbeddedFont[] = []

  for (const font of entries) {
    const family = String(font?.name ?? '').trim()
    if (!family) continue
    for (const slot of SLOTS) {
      const embed = font[slot.tag] as Record<string, string> | undefined
      if (!embed?.id) continue
      const target = rels[embed.id]?.target
      if (!target) continue
      // Relationship targets are usually relative to word/ ("fonts/x.ttf") but
      // can be package-absolute ("/word/fonts/x.ttf"). Same two-try shape
      // resolveImage uses, with the leading slash stripped first.
      const t = target.replace(/^\/+/, '')
      const entry = zip.file(`word/${t}`) ?? zip.file(t)
      if (!entry) continue
      const raw = await entry.async('uint8array')
      const bytes = deobfuscate(raw, embed.fontKey)
      const mime = fontMime(bytes)
      // Not something a browser can load (a corrupt part, or a format we do not
      // recognise). Skipping keeps the rest of the document's fonts usable.
      if (!mime) continue
      out.push({
        family,
        weight: slot.weight,
        style: slot.style,
        src: `data:${mime};base64,${uint8ArrayToBase64(bytes)}`,
      })
    }
  }
  return out
}

// Register the fonts with the browser so text measured during rendering uses
// them. This has to happen — and be awaited — before the renderer paginates:
// laying out against a fallback font and having the real one arrive afterwards
// reflows the text and invalidates every page break already decided.
//
// A no-op outside the browser (the parser runs in Node too) and best-effort
// per font: one font that fails to decode must not take the document with it.
export async function registerFonts(fonts: EmbeddedFont[]): Promise<void> {
  if (fonts.length === 0) return
  const g = globalThis as { document?: { fonts?: FontFaceSet } }
  const set = g.document?.fonts
  if (!set || typeof FontFace === 'undefined') return
  // document.fonts is global and outlives any one parse, so parsing the same
  // file twice would stack a second copy of every face. Skip what is already
  // there rather than growing the set on every call.
  // A registered face reports its family as a CSS <family-name>, which means a
  // name containing a space comes back QUOTED ("DM Sans") while a single-word
  // one does not (Poppins). Comparing raw strings therefore deduplicated only
  // the unquoted families and let the others pile up 4 per parse.
  const famKey = (name: string): string => name.replace(/^['"]|['"]$/g, '')
  const key = (family: string, weight: string, style: string): string =>
    `${famKey(family)}|${weight}|${style}`
  const present = new Set<string>()
  set.forEach(f => present.add(key(f.family, f.weight, f.style)))
  await Promise.all(
    fonts.filter(f => !present.has(key(f.family, f.weight, f.style))).map(async f => {
      try {
        const face = new FontFace(f.family, `url(${f.src})`, { weight: f.weight, style: f.style })
        await face.load()
        set.add(face)
      } catch {
        // Leave this one to the fallback rather than failing the parse.
      }
    }),
  )
}
