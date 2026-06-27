import { XMLParser } from 'fast-xml-parser'

export type RelationshipMap = Record<string, { type: string; target: string }>

const parser = new XMLParser({
  removeNSPrefix: true,
  attributeNamePrefix: '',
  ignoreAttributes: false,
  parseAttributeValue: false,
  isArray: (name) => name === 'Relationship',
})

export function parseRelationships(xml: string): RelationshipMap {
  const doc = parser.parse(xml)
  const rels: RelationshipMap = {}
  const relationships = doc?.Relationships?.Relationship ?? []
  for (const rel of relationships) {
    if (rel.Id && rel.Type && rel.Target) {
      rels[rel.Id] = { type: rel.Type, target: rel.Target }
    }
  }
  return rels
}
