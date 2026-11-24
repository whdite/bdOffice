import type { ChatMessage } from '@genoffice/project-store'
import type { PersistedAgentMessage, WorkspaceAgentEvent } from '../../shared/workspace-agent-api'
export interface DisplayTranscript {
  messages: ChatMessage[]
  cursor: number
  active?: number
  errorRun?: string
}
const FAILURE = '任务未完成，部分操作可能已完成。请查看任务状态与成果记录。'
/** Human history is separate from the model context, which deliberately rolls failed turns back. */
export function displayTranscript(
  history: PersistedAgentMessage[],
  events: WorkspaceAgentEvent[],
  saved?: DisplayTranscript,
): DisplayTranscript {
  let state: DisplayTranscript
  if (saved && Array.isArray(saved.messages) && Number.isSafeInteger(saved.cursor))
    state = { ...saved, messages: saved.messages.map((m) => ({ ...m })) }
  else {
    const visible = history
      .filter(
        (m): m is Extract<PersistedAgentMessage, { role: 'user' | 'assistant' }> =>
          m.role === 'user' || m.role === 'assistant',
      )
      .filter((m) => m.text.trim())
    const users = events.filter(
      (e) => e.type === 'assistant.message' && e.role === 'user' && e.text,
    )
    const overlap = visible.findIndex(
      (m) => m.role === 'user' && users.some((e) => e.text === m.text),
    )
    const prefix = overlap >= 0 ? visible.slice(0, overlap) : visible
    state = {
      messages: prefix.map((m, seq) => ({ seq, ts: '', role: m.role, text: m.text })),
      cursor: 0,
    }
    if (!users.length) state.cursor = events.at(-1)?.sequence ?? 0
  }
  for (const event of events)
    if ((event.sequence ?? 0) > state.cursor) applyDisplayEvent(state, event)
  return state
}
export function applyDisplayEvent(state: DisplayTranscript, event: WorkspaceAgentEvent): void {
  if (event.type === 'assistant.message' && (event.text || event.replaceText)) {
    if (event.role === 'user' && event.text) {
      state.messages.push({
        seq: (state.messages.at(-1)?.seq ?? -1) + 1,
        ts: event.timestamp,
        role: 'user',
        text: event.text.slice(0, 100000),
      })
      state.active = undefined
    } else {
      const active =
        state.active === undefined ? undefined : state.messages.find((m) => m.seq === state.active)
      if (active && event.replaceText && !event.text) {
        state.messages = state.messages.filter(message => message !== active)
        state.active = undefined
      } else if (active) active.text = (event.replaceText ? event.text ?? '' : active.text + event.text).slice(0, 100000)
      else if (event.text) {
        const seq = (state.messages.at(-1)?.seq ?? -1) + 1
        state.messages.push({
          seq,
          ts: event.timestamp,
          role: 'assistant',
          text: event.text.slice(0, 100000),
        })
        state.active = seq
      }
    }
  } else if (event.type === 'error')
    markDisplayFailure(state, event.runId ?? event.id, event.timestamp)
  else if (
    event.type === 'tool.started' ||
    event.type === 'agent.run.review' ||
    (event.type === 'agent.status' &&
      ['completed', 'cancelled', 'error'].includes(event.status ?? ''))
  )
    state.active = undefined
  state.cursor = Math.max(state.cursor, event.sequence ?? 0)
  while (state.messages.length > 500) state.messages.shift()
  while (
    state.messages.length > 1 &&
    Buffer.byteLength(JSON.stringify(state.messages)) > 1024 * 1024
  )
    state.messages.shift()
}
export function markDisplayFailure(state: DisplayTranscript, runId: string, ts: string): void {
  if (state.errorRun !== runId) {
    state.messages.push({
      seq: (state.messages.at(-1)?.seq ?? -1) + 1,
      ts,
      role: 'assistant',
      text: FAILURE,
    })
    state.errorRun = runId
  }
  state.active = undefined
}
