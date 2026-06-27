import { describe, it, expect } from 'vitest'
import { parseStyles, extractRPr } from '../src/parser/styles.js'

// Minimal styles.xml with docDefaults + two styles
const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:sz w:val="24"/>
        <w:rFonts w:ascii="Calibri"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:styleId="Normal" w:type="paragraph">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:styleId="Heading1" w:type="paragraph">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:styleId="Strong" w:type="character">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
  <w:style w:styleId="CycleA" w:type="paragraph">
    <w:basedOn w:val="CycleB"/>
  </w:style>
  <w:style w:styleId="CycleB" w:type="paragraph">
    <w:basedOn w:val="CycleA"/>
  </w:style>
  <w:style w:styleId="GhostStyle" w:type="paragraph">
    <w:basedOn w:val="DoesNotExist"/>
    <w:rPr><w:i/></w:rPr>
  </w:style>
</w:styles>`

describe('parseStyles', () => {
  it('parses docDefaults: sz=24 → fontSize=12', () => {
    const { docDefaults } = parseStyles(STYLES_XML)
    expect(docDefaults.fontSize).toBe(12)
  })

  it('parses docDefaults: rFonts ascii → fontFamily', () => {
    const { docDefaults } = parseStyles(STYLES_XML)
    expect(docDefaults.fontFamily).toBe('Calibri')
  })

  it('resolves basedOn chain: Heading1 basedOn Normal → includes bold', () => {
    const { styleMap } = parseStyles(STYLES_XML)
    expect(styleMap['Heading1']?.bold).toBe(true)
  })

  it('resolves Heading1 sz=32 → fontSize=16', () => {
    const { styleMap } = parseStyles(STYLES_XML)
    expect(styleMap['Heading1']?.fontSize).toBe(16)
  })

  it('resolves character style Strong → bold:true', () => {
    const { styleMap } = parseStyles(STYLES_XML)
    expect(styleMap['Strong']?.bold).toBe(true)
  })

  it('cycle detection: does not infinite loop', () => {
    // Should complete without throwing
    const { styleMap } = parseStyles(STYLES_XML)
    expect(styleMap['CycleA']).toBeDefined()
  })

  it('missing basedOn style: falls back gracefully', () => {
    const { styleMap } = parseStyles(STYLES_XML)
    // GhostStyle basedOn a non-existent style — should still resolve own rPr
    expect(styleMap['GhostStyle']?.italic).toBe(true)
  })
})

describe('extractRPr', () => {
  it('sz=24 → fontSize=12 (half-point conversion)', () => {
    const s = extractRPr({ sz: { val: '24' } })
    expect(s.fontSize).toBe(12)
  })

  it('b with no val → bold:true', () => {
    const s = extractRPr({ b: {} })
    expect(s.bold).toBe(true)
  })

  it('b val="0" → bold:false (explicit negation)', () => {
    const s = extractRPr({ b: { val: '0' } })
    expect(s.bold).toBe(false)
  })

  it('b val="false" → bold:false', () => {
    const s = extractRPr({ b: { val: 'false' } })
    expect(s.bold).toBe(false)
  })

  it('u val="none" → underline:false (explicit negation)', () => {
    const s = extractRPr({ u: { val: 'none' } })
    expect(s.underline).toBe(false)
  })

  it('u val="single" → underline:true', () => {
    const s = extractRPr({ u: { val: 'single' } })
    expect(s.underline).toBe(true)
  })

  it('color val="auto" → color:undefined (not "#auto")', () => {
    const s = extractRPr({ color: { val: 'auto' } })
    expect(s.color).toBeUndefined()
  })

  it('color val="FF0000" → color:"FF0000"', () => {
    const s = extractRPr({ color: { val: 'FF0000' } })
    expect(s.color).toBe('FF0000')
  })

  it('rFonts w:ascii → fontFamily', () => {
    const s = extractRPr({ rFonts: { ascii: 'Times New Roman' } })
    expect(s.fontFamily).toBe('Times New Roman')
  })

  it('rFonts no ascii, w:hAnsi → fontFamily fallback', () => {
    const s = extractRPr({ rFonts: { hAnsi: 'Arial' } })
    expect(s.fontFamily).toBe('Arial')
  })
})
