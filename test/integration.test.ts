import { describe, it, expect, beforeAll } from 'vitest'
import { readFileSync } from 'fs'
import { dirname, join } from 'path'
import { fileURLToPath } from 'url'
import { parse } from '../src/index.js'
import type { DocxDocument, ParagraphBlock, TableBlock } from '../src/types.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const fixturePath = join(__dirname, 'fixtures', 'sample.docx')

let doc: DocxDocument

beforeAll(async () => {
  const buf = readFileSync(fixturePath).buffer as ArrayBuffer
  doc = await parse(buf)
})

describe('integration: sample.docx', () => {
  it('parses without throwing', () => {
    expect(doc).toBeDefined()
    expect(Array.isArray(doc.blocks)).toBe(true)
  })

  it('has at least 10 blocks', () => {
    expect(doc.blocks.length).toBeGreaterThanOrEqual(10)
  })

  describe('style cascade', () => {
    it('Heading1 block: bold=true, fontSize=20 (sz=40 half-points)', () => {
      const h1 = doc.blocks[0] as ParagraphBlock
      expect(h1.type).toBe('paragraph')
      expect(h1.style.bold).toBe(true)
      expect(h1.style.fontSize).toBe(20)
    })

    it('Heading1 block: color from style (1F3864)', () => {
      const h1 = doc.blocks[0] as ParagraphBlock
      expect(h1.style.color).toBe('1F3864')
    })

    it('bold=false override: Heading1 style + w:b val=0 → run.style.bold=false', () => {
      const boldOverride = doc.blocks.find(b => {
        if (b.type !== 'paragraph') return false
        const p = b as ParagraphBlock
        return p.runs.some(r => r.type === 'run' && r.style.bold === false)
      }) as ParagraphBlock | undefined
      expect(boldOverride).toBeDefined()
    })

    it('color auto → no color property in run style', () => {
      const autoColorPara = doc.blocks.findLast(b => {
        if (b.type !== 'paragraph') return false
        const p = b as ParagraphBlock
        return p.runs.some(r => r.type === 'run' && (r as any).text?.includes('Auto color'))
      }) as ParagraphBlock | undefined
      expect(autoColorPara).toBeDefined()
      const run = autoColorPara!.runs[0]
      expect(run.type).toBe('run')
      if (run.type === 'run') {
        expect(run.style.color).toBeUndefined()
      }
    })
  })

  describe('mixed runs', () => {
    it('finds a bold run', () => {
      const hasBold = doc.blocks.some(b => {
        if (b.type !== 'paragraph') return false
        return (b as ParagraphBlock).runs.some(r => r.type === 'run' && r.style.bold === true)
      })
      expect(hasBold).toBe(true)
    })

    it('finds an italic run', () => {
      const hasItalic = doc.blocks.some(b => {
        if (b.type !== 'paragraph') return false
        return (b as ParagraphBlock).runs.some(r => r.type === 'run' && r.style.italic === true)
      })
      expect(hasItalic).toBe(true)
    })

    it('finds an underline run', () => {
      const hasUnderline = doc.blocks.some(b => {
        if (b.type !== 'paragraph') return false
        return (b as ParagraphBlock).runs.some(r => r.type === 'run' && r.style.underline === true)
      })
      expect(hasUnderline).toBe(true)
    })

    it('finds a red run (color=C00000)', () => {
      const hasRed = doc.blocks.some(b => {
        if (b.type !== 'paragraph') return false
        return (b as ParagraphBlock).runs.some(r => r.type === 'run' && r.style.color === 'C00000')
      })
      expect(hasRed).toBe(true)
    })
  })

  describe('lists', () => {
    it('has bullet list paragraphs (numId=1, ordered=false)', () => {
      const bullets = doc.blocks.filter(b =>
        b.type === 'paragraph' && (b as ParagraphBlock).list?.ordered === false
      )
      expect(bullets.length).toBe(3)
    })

    it('has ordered list paragraphs (numId=2, ordered=true)', () => {
      const ordered = doc.blocks.filter(b =>
        b.type === 'paragraph' && (b as ParagraphBlock).list?.ordered === true
      )
      expect(ordered.length).toBe(2)
    })
  })

  describe('table with vMerge', () => {
    it('has a table block', () => {
      const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock | undefined
      expect(tbl).toBeDefined()
    })

    it('table has 3 rows', () => {
      const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock
      expect(tbl.rows.length).toBe(3)
    })

    it('row 0 col 0: rowSpan=2 (vMerge restart)', () => {
      const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock
      expect(tbl.rows[0].cells[0].rowSpan).toBe(2)
    })

    it('row 0 col 1: colSpan=2 (gridSpan)', () => {
      const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock
      expect(tbl.rows[0].cells[1].colSpan).toBe(2)
    })

    it('row 1 has 2 cells (vMerge continue removed)', () => {
      const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock
      expect(tbl.rows[1].cells.length).toBe(2)
    })

    it('row 2 has 3 cells', () => {
      const tbl = doc.blocks.find(b => b.type === 'table') as TableBlock
      expect(tbl.rows[2].cells.length).toBe(3)
    })
  })

  describe('w:ins tracked change', () => {
    it('tracked-change runs are included in output', () => {
      const hasIns = doc.blocks.some(b => {
        if (b.type !== 'paragraph') return false
        return (b as ParagraphBlock).runs.some(r =>
          r.type === 'run' && (r as any).text?.includes('tracked change')
        )
      })
      expect(hasIns).toBe(true)
    })
  })
})
