import type { CellScalar } from './workbook.types'
import type { WorkbookStyleEdit } from '../shared/edit-schemas'

const title: WorkbookStyleEdit = { fontSize: 15, bold: true, fontColor: '#174C45' }
const header: WorkbookStyleEdit = {
  bold: true,
  fillColor: '#215A51',
  fontColor: '#FFFFFF',
  verticalAlignment: 'center',
  wrapText: true,
}
const stripe: WorkbookStyleEdit = { fillColor: '#EFF6F4' }
const total: WorkbookStyleEdit = { fillColor: '#DCECE7', bold: true }

function isLabel(value: CellScalar | undefined): value is string {
  return typeof value === 'string' && /[^\d\s.,/%-]/.test(value) && !/^\s*=/.test(value)
}

/** Recognize a small rectangular table with an optional single-cell title row. */
export function draftTableLayout(
  rows: readonly (readonly CellScalar[])[],
):
  | { headerRow: number; widths: number[]; rowStyles: (WorkbookStyleEdit | undefined)[] }
  | undefined {
  if (rows.length === 0 || rows.length > 200) return
  const first = rows[0]!
  const headerRow =
    isLabel(first[0]) && first.slice(1).every((value) => value === null || value === '') ? 1 : 0
  const headings = rows[headerRow]
  if (!headings) return
  const columns = headings.length
  if (
    columns < 2 ||
    columns > 24 ||
    rows.length * columns > 2000 ||
    !headings.every(isLabel) ||
    new Set(headings.map((value) => value.trim())).size !== columns ||
    rows.slice(headerRow + 1).some((row) => row.length !== columns) ||
    first.length > columns
  )
    return
  const widths = headings.map((label, column) => {
    const measured = [label, ...rows.slice(headerRow + 1).map((row) => row[column])]
      .filter((value) => value != null && !(typeof value === 'string' && value.startsWith('=')))
      .map((value) =>
        [...String(value).slice(0, 64)].reduce(
          (px, ch) => px + (/[^\x00-\x7F]/.test(ch) ? 14 : 8),
          0,
        ),
      )
    return Math.min(248, Math.max(88, 26 + Math.max(...measured)))
  })
  const rowStyles = rows.map((row, index) => {
    if (index < headerRow) return title
    if (index === headerRow) return header
    if (/^(?:合计|总计|总平均|grand total|total)$/i.test(String(row[0] ?? '').trim())) return total
    return (index - headerRow) % 2 === 0 ? stripe : undefined
  })
  return { headerRow, widths, rowStyles }
}
