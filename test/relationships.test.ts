import { describe, it, expect } from 'vitest'
import { parseRelationships } from '../src/parser/relationships.js'

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

describe('parseRelationships', () => {
  it('parses all relationships', () => {
    const map = parseRelationships(RELS_XML)
    expect(Object.keys(map)).toHaveLength(3)
  })

  it('maps rId to type and target', () => {
    const map = parseRelationships(RELS_XML)
    expect(map['rId1'].target).toBe('media/image1.png')
    expect(map['rId1'].type).toContain('image')
  })

  it('returns empty map for empty relationships', () => {
    const xml = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
    expect(parseRelationships(xml)).toEqual({})
  })
})
