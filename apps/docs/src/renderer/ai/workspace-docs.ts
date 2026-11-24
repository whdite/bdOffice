import type { Editor } from '@tiptap/core'
import {
  executeWordOperation,
  editorToolCapability,
  isObservation,
  type AgentSkill,
} from '@genoffice/agent-core'

type Source = { create(instruction: string): AgentSkill; parts(): unknown; busy(): boolean }
const sources = new WeakMap<Editor, Source>()
const runs = new WeakMap<Editor, { id: string; skill: AgentSkill; stop: AbortController }>()
const stopped = new WeakMap<Editor, string>()
export function cancelWorkspaceWord(editor: Editor, id: string) {
  stopped.set(editor, id)
  const current = runs.get(editor)
  if (current?.id === id) current.stop.abort()
}

export function registerWorkspaceDocs(editor: Editor, source: Source) {
  sources.set(editor, source)
  return () => {
    if (sources.get(editor) === source) {
      sources.delete(editor)
      runs.get(editor)?.stop.abort()
      runs.delete(editor)
      stopped.delete(editor)
    }
  }
}

export function wordSnapshot(editor: Editor) {
  const parts = sources.get(editor)?.parts()
  return { ...editor.getJSON(), ...(parts ? { wordParts: parts } : {}) }
}

/** Translate saved workspace calls to the same Word tools the sidebar uses. */
export function workspaceWordCommand(
  editor: Editor,
  command: Record<string, unknown>,
): Record<string, unknown> | string {
  const ids = Array.isArray(command.blockIds) ? command.blockIds : []
  const indexes: number[] = []
  editor.state.doc.forEach((node, _offset, index) => {
    const id =
      typeof node.attrs.docxIndex === 'number'
        ? `block-${node.attrs.docxIndex}`
        : `live-block-${index}`
    if (ids.includes(id) && ['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name))
      indexes.push(index)
  })
  if (indexes.length !== new Set(ids).size)
    return 'Every blockId must identify an editable document block.'
  const target = indexes.length ? { blockIndexes: indexes } : { scope: 'document' }
  const base = { kind: 'word_tool', runId: command.runId ?? 'workspace-legacy' }
  if (command.kind === 'insert_content') {
    const text = String(command.text ?? command.content ?? '')
    if (!text) return 'Provide content to insert.'
    const html = text
      .split(/\r?\n/)
      .map(
        (line) =>
          '<p>' + line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</p>',
      )
      .join('')
    return {
      ...base,
      name: 'insert_content',
      args: { html, afterBlockIndex: indexes.at(-1) ?? editor.state.doc.childCount - 1 },
    }
  }
  const ops: Record<string, unknown>[] = []
  if (command.kind === 'format') {
    if (command.headingLevel !== undefined)
      ops.push({ op: 'setHeadingLevel', target, level: command.headingLevel })
    const { fontColor, fontSizePt, ...flags } = (command.style ?? {}) as Record<string, unknown>
    if (Object.keys(command.style ?? {}).length)
      ops.push({
        op: 'setFont',
        target,
        ...flags,
        ...(fontColor !== undefined ? { color: fontColor } : {}),
        ...(fontSizePt !== undefined ? { fontSize: fontSizePt } : {}),
      })
  } else {
    const find = command.find ?? command.query
    if (!find) return 'Provide find/query text.'
    ops.push({
      op: 'findReplace',
      target,
      find,
      replace: command.replace ?? command.text ?? '',
      matchCase: command.matchCase !== false,
    })
  }
  return { ...base, name: 'apply_ops', args: { ops } }
}

export async function workspaceWordOperation(editor: Editor, command: Record<string, unknown>) {
  const source = sources.get(editor)
  const name = String(command.name),
    input = command.args as Record<string, unknown>
  const reject = (description: string) => ({
    ok: false,
    changed: 0,
    effect: 'not_applied' as const,
    description,
  })
  if (!source || source.busy())
    return reject('The Word editor is busy. Read its current context again.')
  if (
    name !== 'guide' &&
    !isObservation(editorToolCapability({ name, input }, 'docs')) &&
    !editor.isEditable
  )
    return reject('The Word document is read-only.')
  const id = String(command.runId)
  if (stopped.get(editor) === id) return reject('The Word operation was stopped before execution.')
  let run = runs.get(editor)
  if (!run || run.id !== id) {
    run = {
      id,
      skill: source.create(String(command.instruction ?? '')),
      stop: new AbortController(),
    }
    run.skill.buildContext?.()
    runs.set(editor, run)
  }
  const result = await executeWordOperation(run.skill, { id, name, input }, true, run.stop.signal)
  return {
    ok: !result.isError,
    changed: result.mutated || result.effect === 'applied' ? 1 : 0,
    effect: result.effect,
    wordPart: result.wordPart,
    description: result.output,
  }
}
