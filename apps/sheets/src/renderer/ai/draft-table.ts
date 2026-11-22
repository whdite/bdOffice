import { hasExplicitOfficeFormatting } from '@genoffice/agent-core'
import {
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { draftTableLayout } from '@genoffice/xlsx-gateway/domain/draft-table'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

export function defaultTableOperations(
  operations: readonly WorkbookOperation[],
  instruction: string,
  isPristineSheet: (sheetId: string) => boolean,
): WorkbookOperation[] {
  if (!instruction || hasExplicitOfficeFormatting(instruction)) return []
  const ranges = operations.filter((op) => op.op === 'set_range')
  const grid = ranges[0]
  if (ranges.length !== 1 || !grid) return []
  if (
    operations.some((op) => {
      if ('sheetId' in op && op.sheetId !== grid.sheetId) return false
      if (op.op === 'format_range')
        return Object.keys(op.format).some((property) => property !== 'numberFormat')
      return ![
        'set_range',
        'set_cell',
        'set_formula',
        'rename_sheet',
        'add_conditional_format',
        'add_chart',
        'set_filter',
        'set_freeze',
        'set_data_validation',
      ].includes(op.op)
    })
  )
    return []
  const layout = draftTableLayout(grid.values)
  if (!layout || !isPristineSheet(grid.sheetId)) return []
  const bounds = grid.range ? parseRange(grid.range) : undefined
  const origin = grid.start
    ? parseAddress(grid.start)
    : bounds && { row: bounds.startRow, column: bounds.startColumn }
  // Let the regular planner report missing or disagreeing write coordinates.
  if (
    !origin ||
    (bounds && (bounds.startRow !== origin.row || bounds.startColumn !== origin.column))
  )
    return []

  const { sheetId } = grid
  const { row, column } = origin
  const lastColumn = column + layout.widths.length - 1
  if (lastColumn > 16_383 || row + grid.values.length > 1_048_576) return []
  const additions: WorkbookOperation[] = layout.rowStyles.flatMap((style, index) => {
    if (!style) return []
    const { bold, fillColor, fontColor, fontSize, verticalAlignment, wrapText } = style
    const format = {
      bold,
      fillColor,
      fontColor,
      fontSize,
      verticalAlign: verticalAlignment,
      wrapText,
    }
    const from = formatAddress(row + index, column)
    const to = formatAddress(row + index, index < layout.headerRow ? column : lastColumn)
    return [{ op: 'format_range', sheetId, range: from === to ? from : `${from}:${to}`, format }]
  })
  layout.widths.forEach((widthPx, index) =>
    additions.push({
      op: 'set_col_width',
      sheetId,
      column: columnLabel(column + index),
      count: 1,
      widthPx,
    }),
  )
  return additions
}
