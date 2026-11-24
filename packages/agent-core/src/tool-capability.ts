import type { AgentToolCall, ToolExecution } from './types'

export type ToolCapability = 'read' | 'preview' | 'edit' | 'external'
export type ToolEffect = 'not_applied' | 'applied' | 'unknown'

const readers = new Set([
  'get_document_context',
  'read_blocks',
  'read_revisions',
  'read_comments',
  'read_notes',
  'list_styles',
  'search_knowledge_base',
  'read_knowledge_file',
  'get_deck_context',
  'read_slide',
  'list_style_templates',
  'ask_clarification',
  'get_workbook_context',
  'read_range',
  'aggregate_range',
  'load_guide',
  'read_formats',
  'read_sheet_features',
  'read_cells',
  'find_cells',
  'select_range',
  'trace_precedents',
  'trace_dependents',
  'read_pages',
  'search_text',
  'goto_page',
  'read_annotations',
  'list_page_images',
  'list_form_fields',
  'get_outline',
  'read_frontmatter',
  'read_source',
  'read_attachment',
  'web_search',
  'image_search',
  'search_files',
  'list_attachments',
  'read_file',
  'search_document',
  'get_selection',
  'read_context',
  'list_tools',
])

// Tools on the current document. File creation, network/media and extension tools keep their own gate.
const edits = new Set([
  'apply_ops',
  'replace_selection',
  'replace_blocks',
  'insert_content',
  'write_document',
  'save_document',
  'set_header_footer',
  'set_page_setup',
  'insert_section_break',
  'define_style',
  'set_watermark',
  'add_comment',
  'reply_comment',
  'resolve_comment',
  'delete_comment',
  'insert_footnote',
  'insert_endnote',
  'edit_note',
  'delete_note',
  'insert_chart',
  'edit_chart',
  'insert_text_box',
  'accept_changes',
  'reject_changes',
  'propose_operations',
  'save_workbook',
  'save',
  'execute_slide_script',
  'edit_text',
  'edit_block',
  'insert_text',
  'markup_text',
  'set_form_field',
  'replace_page_image',
  'edit_annotation',
  'remove_annotation',
  'plan_page',
])

export function editorToolCapability(
  call: Pick<AgentToolCall, 'name' | 'input'>,
  editor?: string,
): ToolCapability {
  if (readers.has(call.name)) return 'read'
  // Only the Docs executor guarantees this flag validates without writing.
  if (editor === 'docs' && call.name === 'apply_ops' && call.input.dryRun === true) return 'preview'
  if (call.name === 'set_watermark' && typeof call.input.image === 'string') return 'external'
  return edits.has(call.name) ? 'edit' : 'external'
}

// The visible Word rollback restores the body snapshot. Independent Word parts and external effects require a separate decision.
const roundWordEdits = new Set([
  'apply_ops',
  'replace_blocks',
  'replace_selection',
  'insert_content',
  'write_document',
  'insert_chart',
  'edit_chart',
  'save_document',
])
export function isRoundWordEdit(call: Pick<AgentToolCall, 'name' | 'input'>): boolean {
  return editorToolCapability(call, 'docs') === 'edit' && roundWordEdits.has(call.name)
}

export function isObservation(capability: ToolCapability): boolean {
  return capability === 'read' || capability === 'preview'
}

export function recoveryForCapability(
  capability: ToolCapability,
): 'repeat_read' | 'verify_effect' | 'user_reconcile' {
  if (isObservation(capability)) return 'repeat_read'
  return capability === 'edit' ? 'verify_effect' : 'user_reconcile'
}

/** The public Word reader for each component lives beside its tool capability. */
export function wordPartReader(name: string): string | undefined {
  if (name.endsWith('_comment')) return 'read_comments'
  if (name === 'define_style') return 'list_styles'
  if (name.endsWith('_note') || name === 'insert_footnote' || name === 'insert_endnote')
    return 'read_notes'
  if (name === 'accept_changes' || name === 'reject_changes') return 'read_revisions'
  if (
    ['set_header_footer', 'set_page_setup', 'insert_section_break', 'set_watermark'].includes(name)
  )
    return 'get_document_context'
  return undefined
}

/** Body changes, changes to Word parts and external output have distinct receipts. */
export function wordToolReceipt(result: ToolExecution, name: string): ToolExecution {
  if (result.externalEffect) return result
  const effect =
    result.effect ?? (result.mutated ? 'applied' : result.isError ? 'unknown' : 'not_applied')
  const reader = wordPartReader(name)
  if (effect === 'applied' && reader && !result.wordPart?.expected)
    return { ...result, effect, wordPart: { expected: { [reader]: 'read' } } }
  if (
    !result.isError &&
    ['get_document_context', 'read_notes', 'read_comments', 'read_revisions'].includes(name)
  )
    return {
      ...result,
      effect,
      wordPart: { ...result.wordPart, observed: { ...result.wordPart?.observed, [name]: 'read' } },
    }
  return { ...result, effect }
}

export function explicitlyRequestsSave(instruction: string): boolean {
  // Match imperative clauses, including a trailing “and save”. Quoted document contents never enter this check.
  const request = instruction.startsWith('The user attached the following skill')
    ? instruction.split('\n\nUser request:\n').at(-1)!
    : instruction
  const clauses = request
    .replace(/“[^”]*”|"[^"]*"|\x60[^\x60]*\x60/g, '')
    .split(/[。！？.!?\n；;，,]/u)
  return clauses.some((raw) => {
    const clause = raw.trim()
    if (
      /(?:不要|不用|无需|无须|不必|暂不|先别|别|禁止)\s*(?:自动|再)?(?:保存|另存|存盘)|\b(?:do not|don't|without|never|no need to)\b.*\bsav/i.test(
        clause,
      )
    )
      return false
    if (
      /(?:如何|怎么|怎样|为什么|是否|解释|分析|检查|查看|测试|设计).*(?:保存|另存|存盘)|(?:保存|另存)(?:流程|逻辑|机制|策略|按钮|功能|记录|历史|了吗|了没)|\b(?:how (?:do|to|can)|was .*saved)\b/i.test(
        clause,
      )
    )
      return false
    if (
      /(?:^|然后|最后|接着|并且|并|完成后|修改后|记得|再)\s*(?:(?:请你?|帮我|麻烦你?|可以帮我|直接|先|立即|务必)\s*)*(?:(?:把|将).{1,120}?)?(?:保存|另存为|存盘)/u.test(
        clause,
      )
    )
      return true
    return (
      /(?:^|^(?:can|could|would) you\s+(?:please\s+)?|\b(?:please|then|and|also)\s+)save\b/i.test(
        clause,
      ) ||
      /^(?:enregistrez|sauvegardez|guarda|guardar|salva|salve|speichere|speichern|сохрани(?:те)?|kaydet|simpan|احفظ)\s/iu.test(
        clause,
      )
    )
  })
}
