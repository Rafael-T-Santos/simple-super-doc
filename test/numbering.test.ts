import { describe, it, expect } from 'vitest'
import { parseNumbering, emptyNumbering } from '../src/parser/numbering.js'

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
  <w:num w:numId="3">
    <w:abstractNumId w:val="0"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="5"/>
    </w:lvlOverride>
  </w:num>
</w:numbering>`

describe('parseNumbering', () => {
  it('parses abstractNum levels', () => {
    const { abstractNumMap } = parseNumbering(NUMBERING_XML)
    expect(abstractNumMap['0'][0].format).toBe('decimal')
    expect(abstractNumMap['0'][0].start).toBe(1)
  })

  it('parses ordered num', () => {
    const { numMap } = parseNumbering(NUMBERING_XML)
    expect(numMap['1'].abstractNumId).toBe('0')
  })

  it('parses bullet num', () => {
    const { numMap, abstractNumMap } = parseNumbering(NUMBERING_XML)
    const entry = numMap['2']
    const level = abstractNumMap[entry.abstractNumId][0]
    expect(level.format).toBe('bullet')
  })

  it('resolves lvlOverride startOverride', () => {
    const { numMap } = parseNumbering(NUMBERING_XML)
    expect(numMap['3'].startOverride).toBe(5)
  })
})

describe('lvlText (bullet glyph)', () => {
  const XML = `<?xml version="1.0"?>
  <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl>
      <w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl>
      <w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`

  it('captures the literal lvlText for each level', () => {
    const { abstractNumMap } = parseNumbering(XML)
    expect(abstractNumMap['0'][0].text).toBe('-')
    expect(abstractNumMap['0'][1].text).toBe('%2.')
  })

  it('leaves text undefined when the level has no lvlText', () => {
    const { abstractNumMap } = parseNumbering(XML)
    expect(abstractNumMap['0'][2].text).toBeUndefined()
  })
})

describe('emptyNumbering', () => {
  it('returns empty maps', () => {
    const { abstractNumMap, numMap } = emptyNumbering()
    expect(Object.keys(abstractNumMap)).toHaveLength(0)
    expect(Object.keys(numMap)).toHaveLength(0)
  })
})
