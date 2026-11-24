import type {
  AgentPlanStep,
  AgentRunReview,
  PersistedAgentMessage,
  ToolPermissionLevel,
  WorkspaceAgentEvent,
} from '../../shared/workspace-agent-api'
import { dogVisualForEvent, type DogEvent } from './dog-assets'
import { withoutCompactionPreamble } from '@genoffice/agent-core'
import { advanceMemePolicy, type MemePolicyState } from './workspace-chat-memes'

export interface TimelineMessage {
  kind: 'message'
  id: string
  role: 'user' | 'assistant'
  text: string
}

export interface TimelineReasoning {
  kind: 'reasoning'
  id: string
  text: string
}

export interface TimelineToolActivity {
  kind: 'tool'
  id: string
  callId: string
  runId?: string
  name: string
  state: 'running' | 'completed' | 'failed'
  summary?: string
  output?: string
  mutated?: boolean
  display?: NonNullable<WorkspaceAgentEvent['tool']>['display']
}

export interface TimelineFileActivity {
  kind: 'file'
  id: string
  runId?: string
  path: string
  operation: 'read' | 'modified' | 'created'
}

export interface TimelinePlan {
  kind: 'plan'
  id: string
  runId?: string
  steps: AgentPlanStep[]
}

export interface TimelinePermission {
  kind: 'permission'
  id: string
  requestId: string
  level: ToolPermissionLevel
  toolName: string
  title: string
  description: string
  state: 'pending' | 'allowed' | 'denied'
  reason?: 'expired' | 'cancelled'
  round?: NonNullable<WorkspaceAgentEvent['permission']>['round']
}

export interface TimelineRunReview {
  kind: 'run'
  id: string
  run: AgentRunReview
}

export interface TimelineDogMeme {
  kind: 'meme'
  id: string
  event: DogEvent
  mood: ReturnType<typeof dogVisualForEvent>['mood']
  frame: string
  permissionId?: string
}

export type WorkspaceChatTimelineItem =
  | TimelineMessage
  | TimelineReasoning
  | TimelineToolActivity
  | TimelineFileActivity
  | TimelinePlan
  | TimelinePermission
  | TimelineRunReview
  | TimelineDogMeme

export interface TimelineEventOptions {
  /** Add event-driven mascot replies without changing the durable Agent log. */
  includeDogMemes?: boolean
}

export interface WorkspaceChatTimelineState {
  usage?: { inputTokens: number; outputTokens: number }
  items: WorkspaceChatTimelineItem[]
  activeAssistantId: string | null
  activeReasoningId: string | null
  activeRunId?: string | null
  memePolicy?: MemePolicyState
}

export function emptyTimelineState(): WorkspaceChatTimelineState {
  return { items: [], activeAssistantId: null, activeReasoningId: null }
}

export function timelineStateFromHistory(
  history: PersistedAgentMessage[],
): WorkspaceChatTimelineState {
  const items: WorkspaceChatTimelineItem[] = []
  const pendingTools = new Map<string, number[]>()

  const first = history.length - withoutCompactionPreamble(history).length
  for (let messageIndex = first; messageIndex < history.length; messageIndex += 1) {
    const message = history[messageIndex]!
    if (message.role === 'user' || message.role === 'assistant') {
      if (message.text.trim()) {
        items.push({
          kind: 'message',
          id: `history-message-${messageIndex}`,
          role: message.role,
          text: message.text,
        })
      }
      if (message.role === 'assistant') {
        if (message.reasoning?.trim()) {
          items.push({
            kind: 'reasoning',
            id: `history-reasoning-${messageIndex}`,
            text: message.reasoning,
          })
        }
        for (let callIndex = 0; callIndex < (message.toolCalls?.length ?? 0); callIndex += 1) {
          const call = message.toolCalls![callIndex]!
          const itemIndex = items.length
          items.push({
            kind: 'tool',
            id: `history-tool-${messageIndex}-${callIndex}-${call.id}`,
            callId: call.id,
            name: call.name,
            state: 'running',
          })
          const queue = pendingTools.get(call.id) ?? []
          queue.push(itemIndex)
          pendingTools.set(call.id, queue)
        }
      }
      continue
    }

    for (let resultIndex = 0; resultIndex < message.results.length; resultIndex += 1) {
      const result = message.results[resultIndex]!
      const queue = pendingTools.get(result.id)
      const itemIndex = queue?.shift()
      if (queue?.length === 0) pendingTools.delete(result.id)
      if (itemIndex !== undefined) {
        const pending = items[itemIndex]
        if (pending?.kind === 'tool') {
          items[itemIndex] = {
            ...pending,
            state: result.isError ? 'failed' : 'completed',
            output: result.output,
          }
          continue
        }
      }

      items.push({
        kind: 'tool',
        id: `history-result-${messageIndex}-${resultIndex}-${result.id}`,
        callId: result.id,
        name: result.name,
        state: result.isError ? 'failed' : 'completed',
        output: result.output,
      })
    }
  }

  const last = items[items.length - 1]
  return {
    items,
    activeAssistantId: last?.kind === 'message' && last.role === 'assistant' ? last.id : null,
    activeReasoningId: last?.kind === 'reasoning' ? last.id : null,
  }
}

function resetActiveStreams(state: WorkspaceChatTimelineState): WorkspaceChatTimelineState {
  if (state.activeAssistantId === null && state.activeReasoningId === null) return state
  return { ...state, activeAssistantId: null, activeReasoningId: null }
}

function reduceTimelineEventBase(
  state: WorkspaceChatTimelineState,
  event: WorkspaceAgentEvent,
): WorkspaceChatTimelineState {
  if (event.type === 'assistant.message' && (event.text || event.replaceText)) {
    if (event.role === 'user' && event.text) {
      return {
        items: [
          ...state.items,
          { kind: 'message', id: `message-${event.id}`, role: 'user', text: event.text },
        ],
        activeAssistantId: null,
        activeReasoningId: null,
      }
    }

    if (state.activeAssistantId) {
      const index = state.items.findIndex((item) => item.id === state.activeAssistantId)
      const current = state.items[index]
      if (index >= 0 && current?.kind === 'message' && current.role === 'assistant') {
        const items = [...state.items]
        if (event.replaceText && !event.text) {
          items.splice(index,1)
          return {items,activeAssistantId:null,activeReasoningId:null}
        }
        items[index] = { ...current, text: event.replaceText ? event.text ?? '' : `${current.text}${event.text}` }
        return { items, activeAssistantId: current.id, activeReasoningId: null }
      }
    }

    if (!event.text) return state
    const id = `message-${event.id}`
    return {
      items: [...state.items, { kind: 'message', id, role: 'assistant', text: event.text }],
      activeAssistantId: id,
      activeReasoningId: null,
    }
  }

  if (event.type === 'assistant.reasoning_summary' && event.text) {
    if (state.activeReasoningId) {
      const index = state.items.findIndex((item) => item.id === state.activeReasoningId)
      const current = state.items[index]
      if (index >= 0 && current?.kind === 'reasoning') {
        const items = [...state.items]
        items[index] = { ...current, text: `${current.text}${event.text}`.slice(-8_000) }
        return { items, activeAssistantId: null, activeReasoningId: current.id }
      }
    }

    const id = `reasoning-${event.id}`
    return {
      items: [...state.items, { kind: 'reasoning', id, text: event.text.slice(-8_000) }],
      activeAssistantId: null,
      activeReasoningId: id,
    }
  }

  if (event.type === 'tool.started' && event.tool) {
    const callId = event.tool.callId || event.id
    // A history snapshot can already contain the call while its live event is
    // still in the IPC buffer. Do not duplicate a card from the current turn;
    // a reused provider id after a later user message remains valid.
    let lastUserIndex = -1
    let existingCallIndex = -1
    for (let index = 0; index < state.items.length; index += 1) {
      const item = state.items[index]
      if (item?.kind === 'message' && item.role === 'user') lastUserIndex = index
      if (item?.kind === 'tool' && item.callId === callId) existingCallIndex = index
    }
    if (existingCallIndex > lastUserIndex) return state
    return {
      items: [
        ...state.items,
        {
          kind: 'tool',
          id: `tool-${event.id}`,
          callId,
          ...(event.runId ? { runId: event.runId } : {}),
          name: event.tool.name,
          state: 'running',
          summary: event.tool.summary,
        },
      ],
      activeAssistantId: null,
      activeReasoningId: null,
    }
  }

  if ((event.type === 'tool.completed' || event.type === 'tool.failed') && event.tool) {
    const callId = event.tool.callId ?? event.id
    let index = -1
    for (let candidate = state.items.length - 1; candidate >= 0; candidate -= 1) {
      const item = state.items[candidate]
      if (item?.kind === 'tool' && item.callId === callId && item.state === 'running') {
        index = candidate
        break
      }
    }
    if (index < 0 && !event.tool.callId) {
      for (let candidate = state.items.length - 1; candidate >= 0; candidate -= 1) {
        const item = state.items[candidate]
        if (item?.kind === 'tool' && item.name === event.tool.name && item.state === 'running') {
          index = candidate
          break
        }
      }
    }
    if (index < 0) {
      const completed = state.items.find(
        (item) =>
          item.kind === 'tool' &&
          item.callId === callId &&
          item.state !== 'running' &&
          (event.tool!.output === undefined || item.output === event.tool!.output),
      )
      if (completed) return state
    }
    const existing = index >= 0 ? state.items[index] : undefined
    const value: TimelineToolActivity = {
      kind: 'tool',
      id: existing?.id ?? `tool-${event.id}`,
      callId: existing?.kind === 'tool' ? existing.callId : callId,
      ...(event.runId
        ? { runId: event.runId }
        : existing?.kind === 'tool' && existing.runId
          ? { runId: existing.runId }
          : {}),
      name: event.tool.name,
      state: event.type === 'tool.failed' ? 'failed' : 'completed',
      summary: event.tool.summary,
      output: event.tool.output,
      mutated: event.tool.mutated,
      display: event.tool.display,
    }
    if (index < 0) {
      return {
        items: [...state.items, value],
        activeAssistantId: null,
        activeReasoningId: null,
      }
    }
    const items = [...state.items]
    items[index] = value
    return { items, activeAssistantId: null, activeReasoningId: null }
  }

  if (
    (event.type === 'file.read' ||
      event.type === 'file.modified' ||
      event.type === 'document.created') &&
    event.file?.path
  ) {
    // Event ids are durable; replaying a buffered IPC event after loading the
    // persisted log must never create a second activity row.
    if (state.items.some((item) => item.kind === 'file' && item.id === `file-${event.id}`))
      return state
    const operation =
      event.type === 'file.read'
        ? 'read'
        : event.type === 'document.created'
          ? 'created'
          : 'modified'
    return {
      items: [
        ...state.items,
        {
          kind: 'file',
          id: `file-${event.id}`,
          path: event.file.path,
          operation,
          ...(event.runId ? { runId: event.runId } : {}),
        },
      ],
      activeAssistantId: null,
      activeReasoningId: null,
    }
  }

  if (event.type === 'agent.plan' && event.plan) {
    let index = -1
    for (let candidate = state.items.length - 1; candidate >= 0; candidate -= 1) {
      const item = state.items[candidate]
      if (item?.kind === 'message' && item.role === 'user') break
      if (item?.kind === 'plan') {
        index = candidate
        break
      }
    }
    const value: TimelinePlan = {
      kind: 'plan',
      id: index >= 0 ? state.items[index]!.id : `plan-${event.id}`,
      ...(event.runId ? { runId: event.runId } : {}),
      steps: event.plan.steps.map((step) => ({ ...step })),
    }
    if (index < 0) {
      return {
        items: [...state.items, value],
        activeAssistantId: null,
        activeReasoningId: null,
      }
    }
    const items = [...state.items]
    items[index] = value
    return { items, activeAssistantId: null, activeReasoningId: null }
  }

  if (
    (event.type === 'permission.requested' || event.type === 'permission.resolved') &&
    event.permission
  ) {
    const index = state.items.findIndex(
      (item) => item.kind === 'permission' && item.requestId === event.permission!.id,
    )
    const value: TimelinePermission = {
      kind: 'permission',
      id: index >= 0 ? state.items[index]!.id : `permission-${event.permission.id}`,
      requestId: event.permission.id,
      level: event.permission.level,
      toolName: event.permission.toolName,
      title: event.permission.title,
      description: event.permission.description,
      state: event.permission.state,
      ...(event.permission.reason ? { reason: event.permission.reason } : {}),
      ...(event.permission.round ? { round: event.permission.round } : {}),
    }
    if (index < 0) {
      return {
        items: [...state.items, value],
        activeAssistantId: null,
        activeReasoningId: null,
      }
    }
    const items = [...state.items]
    items[index] = value
    return { items, activeAssistantId: null, activeReasoningId: null }
  }

  if (event.type === 'agent.run.review' && event.run) {
    const index = state.items.findIndex(
      (item) => item.kind === 'run' && item.run.id === event.run!.id,
    )
    const value: TimelineRunReview = {
      kind: 'run',
      id: index >= 0 ? state.items[index]!.id : `run-${event.run.id}`,
      run: {
        ...event.run,
        changes: event.run.changes.map((change) => ({
          ...change,
          tools: [...change.tools],
          summaries: [...change.summaries],
        })),
      },
    }
    if (index < 0) {
      return {
        items: [...state.items, value],
        activeAssistantId: null,
        activeReasoningId: null,
      }
    }
    const items = [...state.items]
    items[index] = value
    return { items, activeAssistantId: null, activeReasoningId: null }
  }

  if (event.type === 'error' || (event.type === 'agent.status' && event.status !== 'running')) {
    return resetActiveStreams(state)
  }

  return state
}

function appendDogMeme(
  state: WorkspaceChatTimelineState,
  event: WorkspaceAgentEvent,
  previousPolicy: MemePolicyState | undefined,
): WorkspaceChatTimelineState {
  const { policy, reaction: dogEvent } = advanceMemePolicy(previousPolicy, event)
  state = { ...state, memePolicy: policy }
  if (!dogEvent) return state
  const id = `meme-${event.id}`
  if (state.items.some((item) => item.kind === 'meme' && item.id === id)) return state
  const visual = dogVisualForEvent(dogEvent)
  const nextItems = [
    ...state.items,
    {
      kind: 'meme' as const,
      id,
      event: dogEvent,
      mood: visual.mood,
      frame: visual.frame,
      ...(dogEvent === 'approval' && event.permission ? { permissionId: event.permission.id } : {}),
    },
  ]
  // A restored session can contain hundreds of historical tool events. Keep
  // the transcript intact while bounding image nodes to the most recent
  // progress replies.
  let memeCount = 0
  for (const item of nextItems) if (item.kind === 'meme') memeCount += 1
  if (memeCount > 12) {
    const first = nextItems.findIndex((item) => item.kind === 'meme')
    if (first >= 0) nextItems.splice(first, 1)
  }
  return {
    ...state,
    items: nextItems,
  }
}

export function reduceTimelineEvent(
  state: WorkspaceChatTimelineState,
  event: WorkspaceAgentEvent,
  options: TimelineEventOptions = {},
): WorkspaceChatTimelineState {
  const next = {
    ...reduceTimelineEventBase(state, event),
    ...(state.usage ? { usage: state.usage } : {}),
  }
  const activeRun =
    (event.type === 'agent.status' && event.status !== 'running' && event.status !== 'waiting_user') ||
    event.type === 'error'
      ? null
      : (event.runId ?? state.activeRunId)
  if (activeRun !== undefined) next.activeRunId = activeRun
  if (event.type === 'usage.updated' && event.usage)
    next.usage = {
      inputTokens: (state.usage?.inputTokens ?? 0) + (event.usage.inputTokens ?? 0),
      outputTokens: (state.usage?.outputTokens ?? 0) + (event.usage.outputTokens ?? 0),
    }
  return options.includeDogMemes ? appendDogMeme(next, event, state.memePolicy) : next
}

export type TimelineWorkStep = TimelineToolActivity | TimelineFileActivity | TimelinePlan

export interface TimelineWorkGroup {
  kind: 'work'
  id: string
  steps: TimelineWorkStep[]
  running: boolean
}

export function timelineForDisplay(state: WorkspaceChatTimelineState, running: boolean) {
  const rows: (Exclude<WorkspaceChatTimelineItem, TimelineWorkStep> | TimelineWorkGroup)[] = []
  const groups = new Map<string, TimelineWorkGroup>()
  const pending = new Set(
    state.items.flatMap((item) =>
      item.kind === 'permission' && item.state === 'pending' ? [item.requestId] : [],
    ),
  )
  let turn = 'history'
  for (const item of state.items) {
    if (item.kind === 'message' && item.role === 'user') turn = item.id
    if (item.kind === 'meme' && item.event === 'approval' && !pending.has(item.permissionId ?? ''))
      continue
    if (item.kind !== 'tool' && item.kind !== 'file' && item.kind !== 'plan') {
      rows.push(item)
      continue
    }
    const id = item.runId ? 'work-run-' + item.runId : 'work-turn-' + turn
    let group = groups.get(id)
    if (!group) {
      group = { kind: 'work', id, steps: [], running: false }
      groups.set(id, group)
      rows.push(group)
    }
    group.steps.push(item)
  }
  if (running) {
    const group = groups.get(
      state.activeRunId ? 'work-run-' + state.activeRunId : 'work-turn-' + turn,
    )
    if (group) group.running = true
  }
  return rows
}

/** Apply events buffered while the persisted transcript was loading. A run may
 * have been persisted before the snapshot arrived, so discard only events that
 * are already represented by that snapshot and replay the rest in order. */
export function replayTimelineEvents(
  state: WorkspaceChatTimelineState,
  events: WorkspaceAgentEvent[],
  options: TimelineEventOptions = {},
): WorkspaceChatTimelineState {
  let next = state
  const seenEventIds = new Set<string>()
  const consumedHistoryUsers = new Set<number>()
  const orderedEvents = events
    .map((value, index) => ({ value, index }))
    .sort((a, b) => {
      const left = a.value.sequence
      const right = b.value.sequence
      if (typeof left === 'number' && typeof right === 'number' && left !== right)
        return left - right
      return a.index - b.index
    })
    .map(({ value }) => value)
  for (const event of orderedEvents) {
    // A live IPC event can arrive just before the durable event-log request
    // resolves. The same event is then present in both arrays; event ids are
    // the stable de-duplication key across that race.
    if (seenEventIds.has(event.id)) continue
    seenEventIds.add(event.id)
    if (next.items.some((item) => item.id === `file-${event.id}`)) continue
    if (event.type === 'assistant.message' && event.role === 'user' && event.text) {
      let duplicate = -1
      for (let index = 0; index < next.items.length; index += 1) {
        const item = next.items[index]
        if (
          item?.kind === 'message' &&
          item.role === 'user' &&
          item.id.startsWith('history-') &&
          !consumedHistoryUsers.has(index) &&
          item.text === event.text
        ) {
          duplicate = index
          break
        }
      }
      if (duplicate >= 0) {
        consumedHistoryUsers.add(duplicate)
        continue
      }
    }
    if (event.type === 'assistant.message' && event.role === 'assistant' && event.text) {
      const last = next.items[next.items.length - 1]
      if (
        last?.kind === 'message' &&
        last.role === 'assistant' &&
        last.id.startsWith('history-') &&
        last.text.endsWith(event.text)
      )
        continue
    }
    next = reduceTimelineEvent(next, event, options)
  }
  return next
}
