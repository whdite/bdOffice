import type { AgentSkill } from './skill'
import { validateToolInput } from './tool-input-schema'
import type { AgentToolCall, ToolExecution, WordPartEvidence } from './types'

export const WORD_OPERATIONS = [
  'get_document_context',
  'read_blocks',
  'list_styles',
  'read_comments',
  'read_notes',
  'read_revisions',
  'apply_ops',
  'insert_content',
  'replace_blocks',
  'replace_selection',
  'write_document',
  'set_header_footer',
  'define_style',
  'set_page_setup',
  'insert_section_break',
  'add_comment',
  'reply_comment',
  'resolve_comment',
  'delete_comment',
  'insert_chart',
  'edit_chart',
  'insert_footnote',
  'insert_endnote',
  'edit_note',
  'delete_note',
] as const

/** Every requested property needs a later observation for the same document and property. */
export function unverifiedWordParts(calls: readonly { wordPart?: WordPartEvidence }[]): string[] {
  const missing = new Map<string, string>()
  for (const { wordPart } of calls) updateWordChecks(missing, wordPart)
  return [...missing.keys()].map((k) => (JSON.parse(k) as string[]).filter(Boolean).join(' · '))
}

export function updateWordChecks(missing: Map<string, string>, evidence?: WordPartEvidence) {
  if (!evidence) return
  const key = (part: string) => JSON.stringify([evidence.path ?? '', part])
  for (const [part, value] of Object.entries(evidence.expected ?? {})) missing.set(key(part), value)
  for (const [part, value] of Object.entries(evidence.observed ?? {}))
    if (missing.get(key(part)) === value) missing.delete(key(part))
}

/** Only document-local capabilities may cross the workspace bridge. */
export async function executeWordOperation(
  skill: AgentSkill,
  call: AgentToolCall,
  live: boolean,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const tools = skill.tools.filter(
    (tool) =>
      WORD_OPERATIONS.some((name) => name === tool.name) &&
      (live || !['replace_selection', 'write_document'].includes(tool.name)),
  )
  if (call.name === 'guide')
    return {
      summary: 'Word editing tools',
      effect: 'not_applied',
      output: JSON.stringify({
        instructions: skill.systemPrompt.split('# Tool usage')[1]?.split('# Comments')[0],
        operations: skill.systemPrompt.split('# apply_ops guide')[1],
        tools,
        ...(live ? {} : { selection: 'Use explicit block indexes for a closed document.' }),
      }),
    }
  const definition = tools.find((tool) => tool.name === call.name)
  const error = definition
    ? validateToolInput(definition.inputSchema, call.input)
    : 'This operation is unavailable on this document. Read docx.tools for available operations.'
  if (error) return { summary: call.name, output: error, isError: true, effect: 'not_applied' }
  return skill.executeTool(call, signal)
}
