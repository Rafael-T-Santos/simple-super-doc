import { describe, it, expect } from 'vitest'
import { resolveVMerge } from '../src/parser/table.js'

describe('resolveVMerge', () => {
  it('no merge: passes through unchanged', () => {
    const grid = [
      [{ colSpan: 1, vMerge: 'none' as const, rawData: 'A' }],
      [{ colSpan: 1, vMerge: 'none' as const, rawData: 'B' }],
    ]
    const result = resolveVMerge(grid)
    expect(result).toHaveLength(2)
    expect(result[0][0].rowSpan).toBe(1)
    expect(result[1][0].rowSpan).toBe(1)
  })

  it('restart + continue: rowSpan=2, continuation removed', () => {
    const grid = [
      [{ colSpan: 1, vMerge: 'restart' as const, rawData: 'A' }],
      [{ colSpan: 1, vMerge: 'continue' as const, rawData: '' }],
      [{ colSpan: 1, vMerge: 'none' as const, rawData: 'B' }],
    ]
    const result = resolveVMerge(grid)
    expect(result[0]).toHaveLength(1)
    expect(result[0][0].rowSpan).toBe(2)
    expect(result[1]).toHaveLength(0) // continuation removed
    expect(result[2][0].rawData).toBe('B')
  })

  it('three-row merge: rowSpan=3', () => {
    const grid = [
      [{ colSpan: 1, vMerge: 'restart' as const, rawData: 'top' }],
      [{ colSpan: 1, vMerge: 'continue' as const, rawData: '' }],
      [{ colSpan: 1, vMerge: 'continue' as const, rawData: '' }],
    ]
    const result = resolveVMerge(grid)
    expect(result[0][0].rowSpan).toBe(3)
    expect(result[1]).toHaveLength(0)
    expect(result[2]).toHaveLength(0)
  })

  it('colSpan preserved on resolved cell', () => {
    const grid = [
      [{ colSpan: 2, vMerge: 'restart' as const, rawData: 'wide' }],
      [{ colSpan: 2, vMerge: 'continue' as const, rawData: '' }],
    ]
    const result = resolveVMerge(grid)
    expect(result[0][0].colSpan).toBe(2)
    expect(result[0][0].rowSpan).toBe(2)
  })

  it('two independent merge regions in same table', () => {
    // Col 0: merge rows 0-1; Col 1: independent
    const grid = [
      [
        { colSpan: 1, vMerge: 'restart' as const, rawData: 'A' },
        { colSpan: 1, vMerge: 'none' as const, rawData: 'B' },
      ],
      [
        { colSpan: 1, vMerge: 'continue' as const, rawData: '' },
        { colSpan: 1, vMerge: 'none' as const, rawData: 'C' },
      ],
    ]
    const result = resolveVMerge(grid)
    expect(result[0][0].rowSpan).toBe(2) // merged
    expect(result[0][1].rawData).toBe('B') // independent
    expect(result[1]).toHaveLength(1) // only col 1 remains
    expect(result[1][0].rawData).toBe('C')
  })

  it('empty grid: returns empty', () => {
    expect(resolveVMerge([])).toEqual([])
  })
})
