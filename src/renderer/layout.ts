// Page-layout heuristics for documents that paginate themselves with full-page
// background images (cover slides, framed body pages, closing slides) but carry
// NO explicit page breaks. These are pure functions (no DOM) so they can be
// unit-tested. The thresholds are tuned for proposal/letter-style templates
// (e.g. the Clicksign commercial proposal); see test/layout-heuristics.test.ts.
//
// Why heuristics: such templates place section headings, frames, and watermarks
// as floating images positioned by absolute page coordinates, decoupled from the
// text flow. Word/LibreOffice start each section on a new page even though the
// .docx has zero <w:br type="page"/> or <w:lastRenderedPageBreak/>. We recover
// that structure from the shape and size of the images, not the XML order.

import type { Block, ParagraphBlock, ImageRun, Run, TextRun } from '../types.js'

// A paragraph mark with no run renders 0px in a browser but a full line in Word.
// Empty paragraphs are used as vertical spacers (notably on the cover), so they
// get this min line height. 1.7 matches how the document's fallback fonts size
// an empty line and keeps the cover's customer-name field aligned with Word.
export const EMPTY_LINE_EM = 1.7

// Word docDefaults for this template family: w:line="276" w:lineRule="auto" = 1.15.
export const LINE_HEIGHT = 1.15

// A heading is considered "large text" at or above this point size.
export const HEADING_MIN_PT = 24

// The isPageBackground ImageRun carried by a paragraph (a behindDoc=1 anchor), or null.
export function extractPageBackground(block: ParagraphBlock): ImageRun | null {
  for (const run of block.runs) {
    if (run.type === 'image' && (run as ImageRun).isPageBackground) return run as ImageRun
  }
  return null
}

// A large floating image (the document's watermark / decorative frame) that
// overlays the page rather than flowing with text. Excludes page backgrounds,
// small heading/icon images, and the near-full-width closing slide (which is
// handled as its own edge-to-edge page via fullPageImage).
export function isWatermark(img: ImageRun, pw: number, ph: number): boolean {
  if (img.isPageBackground) return false
  return img.heightPx >= ph * 0.4 && img.widthPx < pw * 0.85
}

// A "page-level" image is either a behindDoc background or a watermark — neither
// participates in text flow.
export function isPageLevelImage(run: Run, pw: number, ph: number): boolean {
  return run.type === 'image' &&
    ((run as ImageRun).isPageBackground || isWatermark(run as ImageRun, pw, ph))
}

// The paragraph with page-level images (background + watermark) stripped, leaving
// the content that flows with text. A paragraph that was ONLY a page-level image
// carrier is dropped (null); an originally-empty paragraph is kept because it
// acts as a vertical spacer (e.g. on the cover).
export function flowOnly(block: ParagraphBlock, pw: number, ph: number): ParagraphBlock | null {
  const hadPageImage = block.runs.some(r => isPageLevelImage(r, pw, ph))
  const runs = block.runs.filter(r => !isPageLevelImage(r, pw, ph))
  if (hadPageImage) {
    const hasContent = runs.some(r => r.type !== 'run' || (r as TextRun).text.length > 0)
    if (!hasContent) return null
  }
  return { ...block, runs }
}

// All watermark images carried by a block (for overlay rendering).
export function watermarksOf(block: Block, pw: number, ph: number): ImageRun[] {
  if (block.type !== 'paragraph') return []
  return block.runs.filter(
    r => r.type === 'image' && isWatermark(r as ImageRun, pw, ph),
  ) as ImageRun[]
}

// Does the block carry any visible content (non-whitespace text or an image)?
export function isBlockVisible(block: Block): boolean {
  if (block.type === 'table') return true
  for (const run of block.runs) {
    if (run.type === 'image') return true
    if ((run as TextRun).text.trim().length > 0) return true
  }
  return false
}

// A wide-short "text-as-image" banner used as a section heading (e.g. the
// "Condições Comerciais" / "Outras informações" headings are images, not text).
export function headingImage(block: Block): boolean {
  if (block.type !== 'paragraph') return false
  for (const run of block.runs) {
    if (run.type !== 'image') continue
    const img = run as ImageRun
    if (img.isPageBackground) continue
    if (img.widthPx >= 300 && img.heightPx <= 70 && img.widthPx / img.heightPx >= 5) {
      return true
    }
  }
  return false
}

// A section heading: a large-text paragraph or a wide-short heading banner.
// Section headings force a new page (they start each framed body section).
export function isHeadingBlock(block: Block): boolean {
  if (block.type !== 'paragraph') return false
  if (headingImage(block)) return true
  for (const run of block.runs) {
    if (run.type === 'run' && (run as TextRun).text.trim() && (run.style.fontSize ?? 0) >= HEADING_MIN_PT) {
      return true
    }
  }
  return false
}

// A small standalone icon (e.g. the 70x70 section glyph) that precedes a text
// heading should travel with it, so the page break goes before the icon.
export function isIconOnly(block: Block): boolean {
  if (block.type !== 'paragraph') return false
  let img: ImageRun | null = null
  for (const run of block.runs) {
    if (run.type === 'image') {
      if (img) return false
      img = run as ImageRun
    } else if ((run as TextRun).text.trim().length > 0) {
      return false
    }
  }
  return !!img && img.widthPx <= 150 && img.heightPx <= 150 && !img.isPageBackground
}

// If a page's only visible content is one near-full-width image (e.g. a full-
// bleed closing slide), return it so it can be rendered edge-to-edge.
export function fullPageImage(blocks: Block[], pageWidthPx: number): ImageRun | null {
  let found: ImageRun | null = null
  for (const block of blocks) {
    if (block.type !== 'paragraph') {
      if (isBlockVisible(block)) return null
      continue
    }
    for (const run of block.runs) {
      if (run.type === 'image') {
        if (found) return null
        found = run as ImageRun
      } else if ((run as TextRun).text.trim().length > 0) {
        return null
      }
    }
  }
  return found && found.widthPx >= pageWidthPx * 0.85 ? found : null
}
