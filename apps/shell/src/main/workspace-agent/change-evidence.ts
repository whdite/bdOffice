import { createHash, randomUUID } from 'node:crypto'
import { parseDocx } from '@genoffice/docx-engine'
import { redactNode } from '../../../../docs/src/renderer/ai/redact-view'
import {
  openPptx,
  elementCNvPrId,
  elementDurableId,
  slideDurableId,
  type SlideElement,
} from '@genoffice/pptx-engine'
import { readBasicWorkbook } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { pdfPagesToText } from '@genoffice/file-parse'
import type { DocumentSnapshot } from './service'
import type { ChangeLocation, LocatedChange } from '../../shared/change-evidence'
import { officePackageEvidence, pdfPackageEvidence } from './package-evidence'

type RecordValue = Record<string, unknown>
interface Unit {
  key: string
  location: ChangeLocation
  text: string
  digest: string
  properties: string
  headingLevel?: number
  cellValue?: string
}
export interface EvidenceObservation {
  wordLayout?: { title?: { align: string; sizes: number[] }; tables: number[][] }
  version: string
  units: Unit[]
  incomplete: boolean
  /** Comparable content, excluding engine-specific model metadata. */
  content: string
}
const MAX_UNITS = 20_000
const MAX_DETAILS = 80
const MAX_EXCERPT = 1_200
const MAX_PARSE_BYTES = 16 * 1024 * 1024
const record = (value: unknown): RecordValue =>
  value && typeof value === 'object' && !Array.isArray(value) ? (value as RecordValue) : {}
const array = (value: unknown): unknown[] => (Array.isArray(value) ? value : [])
const text = (value: unknown): string =>
  value === null || value === undefined ? '' : String(value)
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
function stable(value: unknown): string {
  if (Array.isArray(value)) return '[' + value.map(stable).join(',') + ']'
  if (value && typeof value === 'object')
    return (
      '{' +
      Object.entries(record(value))
        .filter(([, v]) => v !== undefined)
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([k, v]) => JSON.stringify(k) + ':' + stable(v))
        .join(',') +
      '}'
    )
  return JSON.stringify(value) ?? 'null'
}
/** Saving clears this Docs review decoration; text, marks and paragraph formatting remain significant. */
function docxContentModel(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(docxContentModel)
  if (!value || typeof value !== 'object') return value
  return Object.fromEntries(
    Object.entries(record(value)).map(([key, item]) => [
      key,
      key === 'attrs'
        ? Object.fromEntries(
            Object.entries(record(item))
              .filter(([name]) => name !== 'aiChanged')
              .map(([name, entry]) => [name, docxContentModel(entry)]),
          )
        : docxContentModel(item),
    ]),
  )
}
export function evidenceSnapshotDigest(snapshot: DocumentSnapshot, live: boolean): string {
  const bytes = Buffer.from(snapshot.bytesBase64, 'base64')
  return live && snapshot.type === 'docx'
    ? hash(stable(docxContentModel(JSON.parse(bytes.toString('utf8')))))
    : hash(bytes)
}
function nodeText(value: unknown): string {
  const node = record(value)
  if (typeof node.text === 'string') return node.text
  if (node.type === 'docProtected') return text(record(node.attrs).previewText)
  if (node.type === 'hardBreak') return '\n'
  const children = array(node.content)
  return children
    .map(nodeText)
    .join(['docTable', 'table', 'tableRow', 'docTableRow'].includes(text(node.type)) ? '\n' : '')
}
function renderText(value: unknown): string {
  const node = record(value)
  const lines = array(record(node.text).lines)
  if (lines.length)
    return lines
      .map((line) =>
        array(record(line).runs)
          .filter((run) => record(run).isBullet !== true)
          .map((run) => text(record(run).text))
          .join(''),
      )
      .join('\n')
  if (node.type === 'group') return array(node.children).map(renderText).join('\n')
  if (node.type === 'table') return array(node.cells).map(renderText).join(' | ')
  return ''
}
function packageSlideText(value: unknown): string {
  const element = record(value)
  if (element.type === 'group') return array(element.children).map(packageSlideText).join('\n')
  if (element.type === 'table')
    return array(element.rows)
      .flatMap((row) => array(row).map((cell) => packageSlideText({ text: record(cell).text })))
      .join(' | ')
  return array(record(element.text).paragraphs)
    .map((paragraph) =>
      array(record(paragraph).runs)
        .map((run) => text(record(run).text))
        .join(''),
    )
    .join('\n')
}
function column(index: number): string {
  let value = index + 1,
    name = ''
  while (value > 0) {
    name = String.fromCharCode(65 + ((value - 1) % 26)) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

/** Extract actual model/package content; no tool input or assistant summary is used. */
/** Corrective messages enter the model context, so use the same withheld-span projection as Word. */
export async function observeDeliveryDocument(
  snapshot: DocumentSnapshot,
  live: boolean,
  layout = false,
): Promise<EvidenceObservation> {
  if (live && snapshot.type === 'docx') {
    const bytes = Buffer.from(snapshot.bytesBase64, 'base64')
    if (bytes.length <= MAX_PARSE_BYTES && bytes.includes('"redaction"')) {
      const projected = redactNode(JSON.parse(bytes.toString('utf8')))
      snapshot = {
        ...snapshot,
        bytesBase64: Buffer.from(JSON.stringify(projected)).toString('base64'),
      }
    }
  }
  return observeDocument(snapshot, live, layout)
}

export async function observeDocument(
  snapshot: DocumentSnapshot,
  live: boolean,
  inspectLayout = false,
): Promise<EvidenceObservation> {
  const bytes = Buffer.from(snapshot.bytesBase64, 'base64')
  if (bytes.length > MAX_PARSE_BYTES)
    return { version: hash(bytes), units: [], incomplete: true, content: hash('unavailable') }
  const units: Unit[] = []
  const layout: EvidenceObservation['wordLayout'] =
    inspectLayout && snapshot.type === 'docx' ? { tables: [] } : undefined
  let incomplete = false
  const add = (
    key: string,
    location: ChangeLocation,
    value: string,
    properties: unknown = {},
    headingLevel?: number,
    cellValue?: string,
  ) => {
    if (units.length >= MAX_UNITS) {
      incomplete = true
      return
    }
    const props = stable(properties)
    units.push({
      key,
      location,
      text: value,
      properties: props,
      digest: hash(stable([value, properties])),
      ...(headingLevel ? { headingLevel } : {}),
      ...(cellValue !== undefined ? { cellValue } : {}),
    })
  }
  const json: unknown = live ? JSON.parse(bytes.toString('utf8')) : undefined
  if (snapshot.type === 'docx') {
    if (live) {
      if (!Array.isArray(record(json).content)) incomplete = true
      array(record(json).content).forEach((node, index) => {
        const raw = record(node),
          attrs = record(raw.attrs)
        const blockId = Number.isInteger(attrs.docxIndex)
          ? `block-${attrs.docxIndex}`
          : `live-block-${index}`
        add(
          blockId,
          { kind: 'block', blockId, index: index + 1 },
          nodeText(node),
          docxContentModel(node),
          record(node).type === 'docHeading' ? Number(attrs.level) : undefined,
        )
        if (layout && record(node).type === 'docHeading' && !layout.title) {
          const runs = array(record(node).content)
            .map(record)
            .filter((r) => r.type === 'text' && text(r.text).trim())
          layout.title = {
            align: text(attrs.align || 'left'),
            sizes: runs.map(
              (r) =>
                Number(
                  record(
                    array(r.marks)
                      .map(record)
                      .find((m) => m.type === 'docTextStyle')?.attrs,
                  ).sizeHalfPoints,
                ) / 2 || 0,
            ),
          }
        }
        if (layout && record(node).type === 'docTable') {
          const row = record(array(record(node).content)[0])
          const widths = array(row.content).flatMap((c) =>
            array(record(record(c).attrs).colwidth).map(Number),
          )
          layout.tables.push(
            widths.length === array(row.content).length && widths.every((w) => w > 0)
              ? widths
              : array(attrs.colWidthsPct).map(Number),
          )
        }
      })
    } else {
      const document = await parseDocx(bytes)
      document.blocks.forEach((block, index) => {
        if (block.docxIndex === undefined) return
        const blockId = `block-${block.docxIndex}`
        const value =
          block.type === 'table'
            ? (block.table?.rows ?? [])
                .map((row) => row.map((cell) => cell.paras.join('\n')).join(' | '))
                .join('\n')
            : (block.runs ?? []).map((run) => run.text).join('')
        add(
          blockId,
          { kind: 'block', blockId, index: index + 1 },
          value,
          {
            type: block.type,
            level: block.level,
            runs: block.runs,
            table: block.table,
          },
          block.type === 'heading' ? block.level : undefined,
        )
        if (layout && block.type === 'heading' && !layout.title) {
          const style = block.styleId ? document.styles.get(block.styleId)?.display : undefined
          layout.title = {
            align: block.format?.align ?? style?.align ?? 'left',
            sizes: (block.runs ?? [])
              .filter((r) => r.text.trim())
              .map((r) => (r.sizeHalfPoints ?? style?.sizeHalfPoints ?? 0) / 2),
          }
        }
        if (layout && block.type === 'table')
          layout.tables.push(block.table?.colWidthsTwips ?? block.table?.colWidthsPct ?? [])
      })
    }
  } else if (snapshot.type === 'xlsx' || snapshot.type === 'xlsm') {
    if (live) {
      const book = record(json),
        sheets = record(book.sheets),
        styles = record(book.styles)
      if (!book.sheets || typeof book.sheets !== 'object') incomplete = true
      const order = Array.isArray(book.sheetOrder) ? book.sheetOrder.map(text) : Object.keys(sheets)
      for (const sheetId of order) {
        const sheet = record(sheets[sheetId]),
          name = text(sheet.name)
        add(`sheet:${sheetId}`, { kind: 'sheet', sheet: name, sheetId }, name, {
          rowCount: sheet.rowCount,
          columnCount: sheet.columnCount,
          mergeData: sheet.mergeData,
          rowData: sheet.rowData,
          columnData: sheet.columnData,
          hidden: sheet.hidden,
        })
        for (const [row, values] of Object.entries(record(sheet.cellData))) {
          for (const [col, raw] of Object.entries(record(values))) {
            if (!/^\d+$/.test(row) || !/^\d+$/.test(col)) continue
            const cell = record(raw),
              range = column(Number(col)) + (Number(row) + 1)
            const value = cell.f
              ? '=' + text(cell.f).replace(/^=/, '')
              : text(cell.v ?? record(record(cell.p).body).dataStream).replace(/\r?\n$/, '')
            add(
              `cell:${sheetId}:${range}`,
              { kind: 'cell', sheet: name, sheetId, range },
              value,
              typeof cell.s === 'string' ? styles[cell.s] : cell.s,
              undefined,
              cell.v == null ? undefined : text(cell.v),
            )
          }
        }
      }
    } else {
      const book = await readBasicWorkbook(bytes)
      for (const sheet of book.snapshot.sheets) {
        add(
          `sheet:${sheet.id}`,
          { kind: 'sheet', sheet: sheet.name, sheetId: sheet.id },
          sheet.name,
          { merges: sheet.merges, rowHeights: sheet.rowHeights, colWidths: sheet.colWidths },
        )
        for (const [range, cell] of Object.entries(sheet.cells))
          add(
            `cell:${sheet.id}:${range}`,
            { kind: 'cell', sheet: sheet.name, sheetId: sheet.id, range },
            cell.formula ? '=' + cell.formula.replace(/^=/, '') : text(cell.rawValue ?? cell.value),
            sheet.styles?.[range],
            undefined,
            cell.rawValue == null && cell.value == null
              ? undefined
              : text(cell.rawValue ?? cell.value),
          )
      }
    }
  } else if (snapshot.type === 'pptx') {
    if (live) {
      if (!Array.isArray(json)) incomplete = true
      array(json).forEach((slide, index) => {
        const visit = (value: unknown) => {
          const node = record(value),
            elementId = text(node.durableId ?? node.sourceId ?? node.id)
          if (!elementId) {
            incomplete = true
            return
          }
          const { id: _id, sourceId: _sourceId, durableId: _durableId, ...properties } = node
          add(
            `${index + 1}:${text(node.sourceId ?? node.id)}`,
            { kind: 'slide', slide: index + 1, elementId },
            renderText(node),
            properties,
          )
          if (node.type === 'group') array(node.children).forEach(visit)
        }
        array(record(slide).nodes).forEach(visit)
      })
    } else {
      const opened = await openPptx(bytes)
      opened.deck.slides.forEach((slide, index) => {
        const visit = (value: unknown) => {
          const element = record(value),
            engineElement = value as SlideElement,
            alias = elementCNvPrId(engineElement),
            elementId = elementDurableId(engineElement) ?? alias ?? text(element.id)
          if (!elementId) {
            incomplete = true
            return
          }
          add(
            `${slideDurableId(slide)}:${alias ?? elementId}`,
            { kind: 'slide', slide: index + 1, slideId: slideDurableId(slide), elementId },
            packageSlideText(element),
            {
              type: element.type,
              transform: element.transform,
              text: element.text,
              shape: element.shape,
              rows: element.rows,
              chart: element.chart,
            },
          )
          if (element.type === 'group') array(element.children).forEach(visit)
        }
        slide.elements.forEach(visit)
      })
    }
  } else if (snapshot.type === 'md' || snapshot.type === 'html') {
    const value = live
      ? text(
          record(json).html ??
            record(json).body ??
            record(json).text ??
            record(json).content ??
            json,
        )
      : bytes.toString('utf8')
    value
      .split(/\r?\n/)
      .forEach((line, index) => add(`line:${index + 1}`, { kind: 'line', line: index + 1 }, line))
  } else if (snapshot.type === 'pdf' && !live) {
    const pages = await pdfPagesToText(bytes)
    pages.forEach((page, index) =>
      add(
        `page:${index + 1}`,
        { kind: 'page', page: index + 1 },
        typeof page === 'string' ? page : text(record(page).text),
      ),
    )
    // Text extraction cannot establish all annotation/image/layout changes.
    incomplete = true
  } else if (snapshot.type === 'pdf' && live) {
    for (const [kind, values] of Object.entries(record(json))) {
      if (!Array.isArray(values)) continue
      values.forEach((value, index) => {
        const item = record(value),
          page =
            typeof item.pageIndex === 'number'
              ? item.pageIndex + 1
              : typeof item.page === 'number'
                ? item.page
                : undefined
        add(
          `${kind}:${index}`,
          page && page > 0 ? { kind: 'page', page } : { kind: 'file' },
          stable(value),
          { kind },
        )
      })
    }
    incomplete = true
  } else {
    incomplete = true
  }
  return {
    version: evidenceSnapshotDigest(snapshot, live),
    units,
    ...(layout ? { wordLayout: layout } : {}),
    incomplete,
    content: hash(
      stable(
        units
          .map((unit) => {
            const location = unit.location
            const address =
              location.kind === 'cell'
                ? ['cell', location.sheet, location.range]
                : location.kind === 'sheet'
                  ? ['sheet', location.sheet]
                  : location.kind === 'block'
                    ? ['block', location.index]
                    : location.kind === 'slide'
                      ? ['slide', location.slide, location.elementId]
                      : location.kind === 'page'
                        ? ['page', location.page]
                        : location.kind === 'line'
                          ? ['line', location.line]
                          : ['file', unit.key]
            return stable([address, unit.text])
          })
          .sort(),
      ),
    ),
  }
}

/** Bounded excerpts centered on the real difference, including equal-length replacements. */
function excerpts(
  before: string,
  after: string,
): { before: string; after: string; truncated: boolean } {
  let prefix = 0
  while (prefix < Math.min(before.length, after.length) && before[prefix] === after[prefix])
    prefix++
  const start = Math.max(0, prefix - 160)
  const take = (value: string) =>
    (start ? '…' : '') +
    value.slice(start, start + MAX_EXCERPT) +
    (value.length > start + MAX_EXCERPT ? '…' : '')
  return {
    before: take(before),
    after: take(after),
    truncated: start > 0 || before.length > MAX_EXCERPT || after.length > MAX_EXCERPT,
  }
}
export function documentChanges(
  before: EvidenceObservation | undefined,
  after: EvidenceObservation | undefined,
  scope: 'content' | 'resources' = 'content',
): { details: LocatedChange[]; omitted: number; incomplete: boolean } {
  let left = before?.units ?? [],
    right = after?.units ?? []
  if (
    scope === 'content' &&
    left.every((item) => item.location.kind === 'line') &&
    right.every((item) => item.location.kind === 'line')
  ) {
    // Match unchanged lines at both edges. An insertion may move a paragraph
    // down; its new line number alone does not make it a content change.
    let prefix = 0,
      suffix = 0
    while (
      prefix < Math.min(left.length, right.length) &&
      left[prefix]!.digest === right[prefix]!.digest
    )
      prefix++
    while (
      suffix < Math.min(left.length, right.length) - prefix &&
      left[left.length - suffix - 1]!.digest === right[right.length - suffix - 1]!.digest
    )
      suffix++
    left = left.slice(prefix, left.length - suffix)
    right = right.slice(prefix, right.length - suffix)
  }
  const previous = new Map(left.map((unit) => [unit.key, unit]))
  const next = new Map(right.map((unit) => [unit.key, unit]))
  const details: LocatedChange[] = []
  let omitted = 0
  for (const key of new Set([...previous.keys(), ...next.keys()])) {
    const a = previous.get(key),
      b = next.get(key)
    if (a?.digest === b?.digest) continue
    if (details.length >= MAX_DETAILS) {
      omitted++
      continue
    }
    const kind = !a || !b ? 'structure' : a.text !== b.text ? 'text' : 'format'
    const excerpt = excerpts(
      kind === 'format' ? (a?.properties ?? '') : (a?.text ?? ''),
      kind === 'format' ? (b?.properties ?? '') : (b?.text ?? ''),
    )
    details.push({
      id: randomUUID(),
      location: (b ?? a)!.location,
      kind,
      ...excerpt,
      beforeDigest: a?.digest,
      afterDigest: b?.digest,
    })
  }
  if (scope === 'content' && !details.length && before?.version !== after?.version)
    details.push({
      id: randomUUID(),
      location: { kind: 'file' },
      kind: 'content',
      beforeDigest: before?.version,
      afterDigest: after?.version,
    })
  return {
    details,
    omitted,
    incomplete:
      !!before?.incomplete ||
      !!after?.incomplete ||
      omitted > 0 ||
      details.some((detail) => detail.location.kind === 'file'),
  }
}

export async function observeResourceChanges(
  before: DocumentSnapshot | undefined,
  after: DocumentSnapshot | undefined,
): Promise<{ details: LocatedChange[]; omitted: number; incomplete: boolean }> {
  const inspect = async (
    snapshot: DocumentSnapshot | undefined,
  ): Promise<EvidenceObservation | undefined> => {
    if (!snapshot || !['docx', 'xlsx', 'xlsm', 'pptx', 'pdf'].includes(snapshot.type)) return
    const bytes = Buffer.from(snapshot.bytesBase64, 'base64'),
      version = hash(bytes)
    if (bytes.length > MAX_PARSE_BYTES)
      return { version, units: [], incomplete: true, content: version }
    try {
      const values =
        snapshot.type === 'pdf'
          ? await pdfPackageEvidence(bytes)
          : await officePackageEvidence(bytes)
      const units = values.slice(0, MAX_UNITS).map((value): Unit => ({
        key: value.key,
        location: value.location,
        text: value.text,
        properties: stable(value.data),
        digest: hash(stable([value.text, value.data])),
      }))
      return { version, units, incomplete: values.length > MAX_UNITS, content: version }
    } catch {
      return { version, units: [], incomplete: true, content: version }
    }
  }
  const [left, right] = await Promise.all([inspect(before), inspect(after)])
  // An unreadable source cannot establish that every part in the other file is new.
  if ((left?.incomplete && !left.units.length) || (right?.incomplete && !right.units.length))
    return { details: [], omitted: 0, incomplete: true }
  return documentChanges(left, right, 'resources')
}
