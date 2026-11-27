import type { WordLayout } from '@genoffice/agent-core'
import type { PmNode } from '../editor/convert'

function textOf(n: PmNode): string {
  return n.text ?? n.content?.map(textOf).join('') ?? ''
}
function sizeRuns(n: PmNode, size: number): PmNode {
  if (n.type !== 'text')
    return n.content ? { ...n, content: n.content.map((c) => sizeRuns(c, size)) } : n
  const marks = n.marks ?? [],
    style = marks.find((m) => m.type === 'docTextStyle')
  return {
    ...n,
    marks: [
      ...marks.filter((m) => m !== style),
      { type: 'docTextStyle', attrs: { ...style?.attrs, sizeHalfPoints: Math.round(size * 2) } },
    ],
  }
}
function columnProportions(node: PmNode, element?: Element, requested?: number[]) {
  const rows = node.content ?? [],
    cols = rows[0]?.content?.length ?? 0
  if (
    cols < 2 ||
    rows.some(
      (row) =>
        row.content?.length !== cols ||
        row.content.some((c) => Number(c.attrs?.colspan ?? 1) !== 1),
    )
  )
    return
  if (requested?.length === cols) return requested
  const columns = element?.querySelectorAll('colgroup > col')
  const entries = columns?.length
    ? Array.from(columns)
    : Array.from(element?.querySelector('tr')?.children ?? [])
  const widths = entries.map((col) => {
    const value = (col as HTMLElement).style.width || col.getAttribute('width') || ''
    return /^\d+(?:\.\d+)?%$/.test(value) ? parseFloat(value) : 0
  })
  if (
    widths.length === cols &&
    widths.every((n) => n > 0) &&
    Math.abs(widths.reduce((a, b) => a + b) - 100) < 1
  )
    return widths
  // Widen narrative columns in new tables; keep balanced columns equal.
  const lengths = Array.from({ length: cols }, (_, i) => {
    const values = rows
      .slice(0, 64)
      .map((r) => Array.from(textOf(r.content![i]!)).length)
      .sort((a, b) => a - b)
    return Math.max(6, Math.min(58, values[Math.floor(values.length * 0.75)] ?? 0))
  })
  if (Math.max(...lengths) < Math.min(...lengths) * 2.5) return
  const sum = lengths.reduce((a, b) => a + b)
  return lengths.map((n) => (100 * n) / sum)
}
/** Format newly authored material only. Local rewrites retain their own formatting contract. */
export function formatGeneratedWord(
  nodes: PmNode[],
  layout: WordLayout = {},
  html?: Document,
): PmNode[] {
  const first = nodes.find((n) => textOf(n).trim())
  const explicitTitle = !!layout.titleAlign || !!layout.titleSize
  const apparentTitle =
    first &&
    nodes.length > 1 &&
    /(?:周报|月报|纪要|备忘录|分析报告|评审报告|工作总结|工作计划|briefing|report|minutes)$/i.test(
      textOf(first).trim(),
    )
  const title =
    first?.type === 'docParagraph' &&
    (explicitTitle || apparentTitle) &&
    textOf(first).length <= 120
      ? first
      : nodes.find((n) => n.type === 'docHeading')
  const titleEl = html?.querySelector<HTMLElement>('h1,h2,h3,h4,h5,h6')
  const cssSize =
    titleEl?.style.fontSize || titleEl?.querySelector<HTMLElement>('span')?.style.fontSize || ''
  const size =
    layout.titleSize ?? (/^\d+(?:\.\d+)?pt$/i.test(cssSize) ? parseFloat(cssSize) : undefined)
  const align = layout.titleAlign ?? titleEl?.style.textAlign
  const tables = Array.from(html?.querySelectorAll('table') ?? [])
  let tableIndex = 0
  return nodes.map((node) => {
    if (node === title) {
      const heading =
        node.type === 'docParagraph'
          ? { ...node, type: 'docHeading', attrs: { ...node.attrs, level: 1, styleId: null } }
          : node
      const styled = size && size >= 1 && size <= 400 ? sizeRuns(heading, size) : heading
      return align && ['left', 'right', 'center', 'justify'].includes(align)
        ? { ...styled, attrs: { ...styled.attrs, align } }
        : styled
    }
    if (node.type !== 'docTable') return node
    const widths = columnProportions(node, tables[tableIndex++], layout.columns)
    if (!widths) return node
    return {
      ...node,
      attrs: { ...node.attrs, colWidthsPct: widths, widthPct: 100, widthPx: null },
      content: node.content!.map((row) => ({
        ...row,
        content: row.content!.map((cell) => ({
          ...cell,
          attrs: { ...cell.attrs, colwidth: null },
        })),
      })),
    }
  })
}
