/** Pure Agent boundary validation, shared by live editors and closed-file sessions. */
import { WORD_OPERATIONS } from '@genoffice/agent-core'
import { validateToolInput } from './tool-input-schema'
import { CHART_TYPES, WORKBOOK_FORMAT_SCHEMA } from './workbook-tool-schema'
import { DOCX_STYLE_SCHEMA } from './docx-tool-schema'
type Command = Record<string, unknown>
const has = (command: Command, key: string) => command[key] !== undefined
const nonempty = (value: unknown) => typeof value === 'string' && value.trim().length > 0
const scalar = (value: unknown) =>
  value === null ||
  typeof value === 'string' ||
  typeof value === 'boolean' ||
  (typeof value === 'number' && Number.isFinite(value))

export function documentCommandKind(command: Command): string {
  return String(command.kind ?? command.operation ?? 'replace_text')
}

export function documentCommandError(type: string, command: Command): string | null {
  const kind = documentCommandKind(command)
  if (type === 'docx' && kind === 'word_tool') {
    return (command.name === 'guide' || WORD_OPERATIONS.some((name) => name === command.name)) &&
      typeof command.runId === 'string' &&
      command.args !== null &&
      typeof command.args === 'object' &&
      !Array.isArray(command.args)
      ? null
      : 'Invalid Word tool or arguments. Read docx.tools for available operations.'
  }
  const allowed: Record<string, readonly string[]> = {
    docx: ['edit', 'replace_text', 'replace_blocks', 'insert_content', 'format'],
    pptx: ['edit', 'replace_text', 'regenerate_slide'],
    xlsx: [
      'edit',
      'replace_text',
      'write_range',
      'set_formula',
      'create_chart',
      'edit_chart',
      'format_range',
      'insert_sheet',
      'delete_sheet',
      'sort',
      'filter',
    ],
    md: ['edit', 'replace_text', 'insert', 'insert_content'],
    pdf: ['pdf_annotate'],
  }
  if (!allowed[type === 'xlsm' ? 'xlsx' : type]?.includes(kind))
    return `Unsupported ${type.toUpperCase()} operation: ${kind}.`
  if (has(command, 'kind') && has(command, 'operation') && command.kind !== command.operation)
    return 'kind and operation disagree.'
  for (const aliases of [
    ['replace', 'text', 'content'],
    ['find', 'query'],
    ['range', 'address', 'dataRange'],
    ['sheet', 'sheetName', 'sheetId'],
    ['slide', 'slideIndex'],
  ]) {
    const supplied = aliases.filter((key) => has(command, key))
    if (supplied.length > 1 && supplied.some((key) => command[key] !== command[supplied[0]!]))
      return `Conflicting aliases: ${supplied.join(', ')}. Supply one unambiguous value.`
  }
  if (type === 'xlsx' || type === 'xlsm') {
    if (has(command, 'chartType') && !CHART_TYPES.includes(String(command.chartType)))
      return 'Unsupported chartType.'
    if (kind === 'insert_sheet')
      return nonempty(command.name) ? null : 'Provide the new worksheet name.'
    if (kind === 'edit_chart') {
      if (!nonempty(command.chartPath)) return 'Provide chartPath from workbook inspection.'
      return ['title', 'chartType', 'legend', 'seriesColors'].some((key) => has(command, key))
        ? null
        : 'Provide at least one chart property to change.'
    }
    if (!nonempty(command.sheet ?? command.sheetName ?? command.sheetId))
      return 'Provide an explicit worksheet (sheet); inspect the workbook first.'
    if (kind === 'delete_sheet') return null
    if (kind === 'filter') {
      const action = command.action ?? 'set'
      if (!['set', 'clear', 'criteria'].includes(String(action)))
        return 'Unsupported filter action.'
      if (action === 'clear') return null
      if (action === 'criteria')
        return nonempty(command.column) &&
          (command.values === null ||
            (Array.isArray(command.values) && command.values.every((v) => typeof v === 'string')))
          ? null
          : 'Filter criteria require column and values (an array of strings or null).'
    }
    if (!nonempty(command.range ?? command.address ?? command.dataRange))
      return 'Provide an explicit worksheet range; no default cell will be written.'
    if (kind === 'format_range')
      return validateToolInput(WORKBOOK_FORMAT_SCHEMA, command.format, 'format')
    if (kind === 'create_chart') return nonempty(command.chartType) ? null : 'Provide chartType.'
    if (kind === 'sort')
      return nonempty(command.byColumn) && ['asc', 'desc'].includes(String(command.order))
        ? null
        : 'Sort requires byColumn and order (asc or desc).'
    if (kind === 'filter') return null
    if (has(command, 'find') || has(command, 'query'))
      return 'Spreadsheet writes accept value, values, or formula; find/replace is not supported by both backends.'
    const payloads = ['value', 'values', 'formula'].filter((key) => has(command, key))
    if (payloads.length !== 1)
      return 'Provide exactly one of value, values, or formula. Use value: null to explicitly clear cells.'
    if (payloads[0] === 'value' && !scalar(command.value))
      return 'value must be a cell scalar or null.'
    if (payloads[0] === 'formula' && !nonempty(command.formula)) return 'formula must be nonempty.'
    if (
      payloads[0] === 'values' &&
      (!Array.isArray(command.values) ||
        !command.values.length ||
        command.values.some((row) => !Array.isArray(row) || !row.length || !row.every(scalar)))
    )
      return 'values must be a nonempty two-dimensional array of cell scalars.'
    return null
  }
  if (type === 'pdf' || kind === 'regenerate_slide') return null
  if (
    has(command, 'blockIds') &&
    (!Array.isArray(command.blockIds) ||
      !command.blockIds.length ||
      command.blockIds.some((id) => !nonempty(id)) ||
      new Set(command.blockIds).size !== command.blockIds.length)
  )
    return 'blockIds must be a nonempty list of unique block IDs.'
  if (type === 'docx' && kind === 'format') {
    if (!Array.isArray(command.blockIds) || !command.blockIds.length)
      return 'Formatting requires explicit blockIds from docx inspection.'
    if (['find', 'replace', 'text', 'content'].some((key) => has(command, key)))
      return 'Formatting must not include text-replacement parameters.'
    if (
      command.headingLevel !== undefined &&
      (!Number.isInteger(command.headingLevel) ||
        Number(command.headingLevel) < 0 ||
        Number(command.headingLevel) > 6)
    )
      return 'headingLevel must be 0 (body) or 1 through 6.'
    if (command.style !== undefined)
      return validateToolInput(DOCX_STYLE_SCHEMA, command.style, 'style')
    return command.headingLevel !== undefined ? null : 'Provide a style patch or headingLevel.'
  }
  if (kind === 'insert' || kind === 'insert_content') {
    if (['find', 'query', 'replace'].some((key) => has(command, key)))
      return 'Insertion cannot include find/replace parameters.'
    if (!nonempty(command.content ?? command.text)) return 'Provide content to insert.'
    if (type === 'md' && (!Number.isInteger(command.afterLine) || Number(command.afterLine) < 0))
      return 'Provide afterLine explicitly (0 inserts at the start).'
    return null
  }
  if (!['replace', 'text', 'content'].some((key) => typeof command[key] === 'string'))
    return 'Provide replacement text explicitly; an empty string means intentional deletion.'
  const find = command.find ?? command.query
  if (type === 'md' && (has(command, 'startLine') || has(command, 'endLine'))) {
    if (has(command, 'find') || has(command, 'query'))
      return 'Choose either find/replace or a line range, not both.'
    const start = command.startLine
    const end = command.endLine ?? start
    if (
      !Number.isInteger(start) ||
      Number(start) < 1 ||
      !Number.isInteger(end) ||
      Number(end) < Number(start)
    )
      return 'Line ranges require 1 <= startLine <= endLine.'
    return null
  }
  if (!nonempty(find))
    return 'Provide nonempty find/query text; use a formatting tool for style changes, not text replacement.'
  if (type === 'pptx') {
    if (has(command, 'matchCase') && typeof command.matchCase !== 'boolean')
      return 'matchCase must be a boolean.'
    if (has(command, 'slides') && (has(command, 'slide') || has(command, 'slideIndex')))
      return 'Choose slide or slides, not both.'
    const slides =
      command.slides ??
      (has(command, 'slide') || has(command, 'slideIndex')
        ? [command.slide ?? command.slideIndex]
        : undefined)
    if (
      slides !== undefined &&
      (!Array.isArray(slides) ||
        !slides.length ||
        slides.some((page) => !Number.isInteger(page) || Number(page) < 1) ||
        new Set(slides).size !== slides.length)
    )
      return 'slides must contain unique existing 1-based slide numbers.'
  }
  return null
}
