import { XMLParser } from 'fast-xml-parser'

type LevelInfo = { format: string; start: number }

export type AbstractNumMap = Record<string, Record<number, LevelInfo>>
export type NumEntry = { abstractNumId: string; startOverride?: number }
export type NumMap = Record<string, NumEntry>

const parser = new XMLParser({
  removeNSPrefix: true,
  attributeNamePrefix: '',
  ignoreAttributes: false,
  parseAttributeValue: false,
  isArray: (name) => ['abstractNum', 'num', 'lvl', 'lvlOverride'].includes(name),
})

export function parseNumbering(xml: string): { abstractNumMap: AbstractNumMap; numMap: NumMap } {
  const doc = parser.parse(xml)
  const numbering = doc?.numbering ?? {}

  const abstractNumMap: AbstractNumMap = {}
  for (const an of (numbering.abstractNum ?? []) as Record<string, unknown>[]) {
    const id = String((an as Record<string, string>).abstractNumId)
    abstractNumMap[id] = {}
    for (const lvl of (an.lvl ?? []) as Record<string, unknown>[]) {
      const ilvl = parseInt(String((lvl as Record<string, string>).ilvl), 10)
      const fmtNode = lvl.numFmt as Record<string, string> | undefined
      const startNode = lvl.start as Record<string, string> | undefined
      const format = fmtNode?.val ?? 'bullet'
      const start = parseInt(startNode?.val ?? '1', 10)
      abstractNumMap[id][ilvl] = { format, start: isNaN(start) ? 1 : start }
    }
  }

  const numMap: NumMap = {}
  for (const n of (numbering.num ?? []) as Record<string, unknown>[]) {
    const numId = String((n as Record<string, string>).numId)
    const abstractRef = n.abstractNumId as Record<string, string> | undefined
    const abstractNumId = abstractRef?.val ?? ''

    let startOverride: number | undefined
    for (const override of (n.lvlOverride ?? []) as Record<string, unknown>[]) {
      const ilvl = parseInt(String((override as Record<string, string>).ilvl), 10)
      if (ilvl === 0) {
        const soNode = override.startOverride as Record<string, string> | undefined
        if (soNode?.val !== undefined) {
          startOverride = parseInt(soNode.val, 10)
        }
      }
    }

    numMap[numId] = { abstractNumId, ...(startOverride !== undefined ? { startOverride } : {}) }
  }

  return { abstractNumMap, numMap }
}

export function emptyNumbering(): { abstractNumMap: AbstractNumMap; numMap: NumMap } {
  return { abstractNumMap: {}, numMap: {} }
}
