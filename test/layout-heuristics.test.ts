import { describe, it, expect } from 'vitest'
import type { Block, ParagraphBlock, TableBlock, ImageRun, TextRun } from '../src/types.js'
import {
  isWatermark, isPageLevelImage, isHeadingBlock, headingImage, isIconOnly,
  fullPageImage, isBlockVisible, flowOnly, extractPageBackground,
} from '../src/renderer/layout.js'

// A4 landscape at 96 DPI, matching the Clicksign template's page box.
const PW = 1123
const PH = 794

// ── builders ───────────────────────────────────────────────────────────────
function img(widthPx: number, heightPx: number, isPageBackground = false): ImageRun {
  return { type: 'image', src: 'data:,', widthPx, heightPx, ...(isPageBackground ? { isPageBackground } : {}) }
}
function text(t: string, fontSize?: number): TextRun {
  return { type: 'run', text: t, style: fontSize != null ? { fontSize } : {} }
}
function para(...runs: Array<ImageRun | TextRun>): ParagraphBlock {
  return { type: 'paragraph', style: {}, runs }
}
const table: TableBlock = { type: 'table', rows: [{ cells: [{ rowSpan: 1, colSpan: 1, blocks: [] }] }] }

// Representative blocks from template_fotos_marca_dagua.docx.
const coverBg = img(1130, 802, true)          // B0 orange cover background
const frameBg = img(1128, 796, true)          // B38 frame background
const condicoesHeadingImg = img(487, 42)      // B19 "Condições Comerciais" banner
const cobrancaHeadingImg = img(515, 45)       // B29 "Cobrança e Pagamento" banner
const sectionIcon = img(70, 70)               // B40 SLA section icon
const watermarkFrame = img(931, 657)          // B33 rId9 decorative frame
const closingSlide = img(1122, 793)           // B96 closing full-page image
const nivesHeadingText = para(text('Níveis de Serviço', 29)) // B41 large text heading

describe('isWatermark', () => {
  it('flags the large decorative frame (931x657)', () => {
    expect(isWatermark(watermarkFrame, PW, PH)).toBe(true)
  })
  it('does NOT flag the near-full-width closing slide (handled separately)', () => {
    expect(isWatermark(closingSlide, PW, PH)).toBe(false)
  })
  it('does NOT flag small heading/icon images', () => {
    expect(isWatermark(condicoesHeadingImg, PW, PH)).toBe(false)
    expect(isWatermark(sectionIcon, PW, PH)).toBe(false)
  })
  it('does NOT flag a page background', () => {
    expect(isWatermark(coverBg, PW, PH)).toBe(false)
  })
})

describe('isPageLevelImage', () => {
  it('treats backgrounds and watermarks as page-level (out of flow)', () => {
    expect(isPageLevelImage(coverBg, PW, PH)).toBe(true)
    expect(isPageLevelImage(watermarkFrame, PW, PH)).toBe(true)
  })
  it('treats heading banners and text as in-flow', () => {
    expect(isPageLevelImage(condicoesHeadingImg, PW, PH)).toBe(false)
    expect(isPageLevelImage(text('hi'), PW, PH)).toBe(false)
  })
})

describe('headingImage', () => {
  it('flags wide-short text-as-image banners', () => {
    expect(headingImage(para(condicoesHeadingImg))).toBe(true)
    expect(headingImage(para(cobrancaHeadingImg))).toBe(true)
  })
  it('rejects square icons and large frames', () => {
    expect(headingImage(para(sectionIcon))).toBe(false)
    expect(headingImage(para(watermarkFrame))).toBe(false)
  })
})

describe('isHeadingBlock', () => {
  it('flags large-text headings and banner-image headings', () => {
    expect(isHeadingBlock(nivesHeadingText)).toBe(true)
    expect(isHeadingBlock(para(condicoesHeadingImg))).toBe(true)
  })
  it('rejects normal body text and tables', () => {
    expect(isHeadingBlock(para(text('Plano contratado', 13)))).toBe(false)
    expect(isHeadingBlock(table)).toBe(false)
  })
})

describe('isIconOnly', () => {
  it('flags a small standalone icon', () => {
    expect(isIconOnly(para(sectionIcon))).toBe(true)
  })
  it('rejects icon+text, large images, and page backgrounds', () => {
    expect(isIconOnly(para(sectionIcon, text('Níveis')))).toBe(false)
    expect(isIconOnly(para(watermarkFrame))).toBe(false)
    expect(isIconOnly(para(coverBg))).toBe(false)
  })
})

describe('fullPageImage', () => {
  it('returns the closing slide when it is the only content on a page', () => {
    expect(fullPageImage([para(closingSlide)], PW)).toBe(closingSlide)
  })
  it('ignores trailing empty paragraphs around the image', () => {
    expect(fullPageImage([para(), para(closingSlide), para()], PW)).toBe(closingSlide)
  })
  it('returns null when the page also has text or a narrow image', () => {
    expect(fullPageImage([para(closingSlide), para(text('hi'))], PW)).toBeNull()
    expect(fullPageImage([para(condicoesHeadingImg)], PW)).toBeNull()
  })
})

describe('isBlockVisible', () => {
  it('true for text, images, and tables; false for empty paragraphs', () => {
    expect(isBlockVisible(para(text('x')))).toBe(true)
    expect(isBlockVisible(para(img(10, 10)))).toBe(true)
    expect(isBlockVisible(table)).toBe(true)
    expect(isBlockVisible(para(text('   ')))).toBe(false)
    expect(isBlockVisible(para())).toBe(false)
  })
})

describe('flowOnly', () => {
  it('keeps an originally-empty paragraph as a spacer', () => {
    const empty = para(text(''))
    expect(flowOnly(empty, PW, PH)).not.toBeNull()
  })
  it('drops a paragraph that was purely a page-level image carrier', () => {
    expect(flowOnly(para(coverBg), PW, PH)).toBeNull()
    expect(flowOnly(para(watermarkFrame), PW, PH)).toBeNull()
  })
  it('keeps text but strips the page-level image (e.g. bg + text)', () => {
    const mixed = para(frameBg, text('Fidelização'))
    const result = flowOnly(mixed, PW, PH)
    expect(result).not.toBeNull()
    expect(result!.runs.some(r => r.type === 'image')).toBe(false)
    expect(result!.runs.some(r => r.type === 'run' && (r as TextRun).text === 'Fidelização')).toBe(true)
  })
  it('keeps an in-flow heading banner image', () => {
    const result = flowOnly(para(condicoesHeadingImg), PW, PH)
    expect(result).not.toBeNull()
    expect(result!.runs.some(r => r.type === 'image')).toBe(true)
  })
})

describe('extractPageBackground', () => {
  it('returns the background ImageRun, or null', () => {
    expect(extractPageBackground(para(coverBg))).toBe(coverBg)
    expect(extractPageBackground(para(text('x')))).toBeNull()
    expect(extractPageBackground(para(condicoesHeadingImg))).toBeNull()
  })
})

describe('regression: distinct image classes are mutually exclusive', () => {
  // Each real image in the template should match exactly one role so the
  // renderer routes it correctly. Guards against threshold drift.
  const cases: Array<[string, Block, { heading: boolean; icon: boolean; watermark: boolean }]> = [
    ['cover bg', para(coverBg), { heading: false, icon: false, watermark: false }],
    ['heading banner', para(condicoesHeadingImg), { heading: true, icon: false, watermark: false }],
    ['section icon', para(sectionIcon), { heading: false, icon: true, watermark: false }],
    ['watermark frame', para(watermarkFrame), { heading: false, icon: false, watermark: true }],
  ]
  for (const [name, block, want] of cases) {
    it(name, () => {
      expect(isHeadingBlock(block)).toBe(want.heading)
      expect(isIconOnly(block)).toBe(want.icon)
      const wm = block.type === 'paragraph' &&
        block.runs.some(r => r.type === 'image' && isWatermark(r as ImageRun, PW, PH))
      expect(wm).toBe(want.watermark)
    })
  }
})
