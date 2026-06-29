import type JSZip from 'jszip'
import type { RelationshipMap } from './relationships.js'

const MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
}

const SKIP_EXTS = new Set(['emf', 'wmf'])

function uint8ArrayToBase64(bytes: Uint8Array): string {
  let b64 = ''
  for (let i = 0; i < bytes.length; i += 8192) {
    b64 += String.fromCharCode(...bytes.subarray(i, i + 8192))
  }
  return btoa(b64)
}

export async function resolveImage(
  rId: string,
  relationshipMap: RelationshipMap,
  zip: JSZip,
): Promise<{ src: string } | null> {
  const rel = relationshipMap[rId]
  if (!rel) return null

  if (!rel.type.includes('image')) return null

  // External (linked) image: the relationship target is a URL, not a zip entry.
  // Hand the URL straight to the <img> so the picture is not lost.
  if (/^https?:\/\//i.test(rel.target)) return { src: rel.target }

  const ext = rel.target.split('.').pop()?.toLowerCase() ?? ''
  if (SKIP_EXTS.has(ext)) return null

  const mime = MIME[ext]
  if (!mime) {
    console.warn(`[simple-super-doc] unknown image extension ".${ext}" for "${rel.target}"`)
    return null
  }

  const entry = zip.file(`word/${rel.target}`) ?? zip.file(rel.target)
  if (!entry) {
    console.warn(`[simple-super-doc] image not found in zip: "${rel.target}"`)
    return null
  }

  const bytes = await entry.async('uint8array')
  const src = `data:${mime};base64,${uint8ArrayToBase64(bytes)}`
  return { src }
}
