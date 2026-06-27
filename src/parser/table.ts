// Two-pass vMerge resolution for OOXML tables.
//
// Pass 1: collect raw cells with colSpan and vMerge tag.
// Pass 2: walk column-by-column; restart cells keep rowSpan:1; continue cells
//         increment the restart cell's rowSpan and are removed from the grid.

export type RawCell = {
  colSpan: number
  vMerge: 'restart' | 'continue' | 'none'
  rawData: unknown // opaque — passed back to DocumentParser for content extraction
}

export type ResolvedCell = {
  colSpan: number
  rowSpan: number
  rawData: unknown
}

export function resolveVMerge(grid: RawCell[][]): ResolvedCell[][] {
  if (grid.length === 0) return []

  const rowCount = grid.length
  // Build a virtual column map; track the restart cell per physical column.
  // colIndex accounts for colSpan so each physical col has a slot.

  // We need the max physical columns to size our tracking array.
  let maxCols = 0
  for (const row of grid) {
    let count = 0
    for (const cell of row) count += cell.colSpan
    maxCols = Math.max(maxCols, count)
  }

  // restartCell[col] points to the ResolvedCell that started a vMerge in that column.
  const restartCell: Array<ResolvedCell | null> = new Array(maxCols).fill(null)

  const output: ResolvedCell[][] = []

  for (let r = 0; r < rowCount; r++) {
    const outRow: ResolvedCell[] = []
    let col = 0

    for (const raw of grid[r]) {
      const resolved: ResolvedCell = {
        colSpan: raw.colSpan,
        rowSpan: 1,
        rawData: raw.rawData,
      }

      if (raw.vMerge === 'restart') {
        restartCell[col] = resolved
        outRow.push(resolved)
      } else if (raw.vMerge === 'continue') {
        const starter = restartCell[col]
        if (starter) {
          starter.rowSpan++
          // Do not push — continuation cells are removed from output
        } else {
          // Malformed DOCX: continue without restart; treat as normal cell
          outRow.push(resolved)
        }
      } else {
        restartCell[col] = null
        outRow.push(resolved)
      }

      col += raw.colSpan
    }

    output.push(outRow)
  }

  return output
}
