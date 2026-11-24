import type { AgentSessionBudget, AgentSessionSnapshot } from './session'
import type { AgentMessage, AgentToolCall, AgentTokenUsage, ToolExecution } from './types'

/** Host-owned execution. Checkpoints contain complete message/tool pairs, never an executable queue. */
export interface TaskStartInput {
  goal: string
  budget: AgentSessionBudget
  snapshot: AgentSessionSnapshot
}
export interface TaskExecutionStart {
  id: string
  /** Fresh host-selected historical data, excluded from recoverable conversation messages. */
  memoryContext?: string
  checkpoint?: AgentSessionSnapshot
  remainingMs: number
}
export interface TaskApprovalScope {
  covered?: boolean
  round?: { path: string; baseVersion: string }
}

export interface TaskToolPermit {
  stepId: string
  reused?: ToolExecution
}
export interface TaskExecutionEnd {
  state: 'completed' | 'cancelled' | 'failed' | 'waiting_user'
  text?: string
  error?: string
  question?: string
}
/** Every method may fail closed. This port belongs to one run, not a mutable current tab. */
export interface AgentTaskExecution {
  readonly taskId?: string
  supportsClarification?: boolean
  start(input: TaskStartInput, signal: AbortSignal): Promise<TaskExecutionStart | null>
  reserveModel(checkpoint: AgentSessionSnapshot, signal: AbortSignal): Promise<string>
  reserveAuxiliaryModel?(signal?: AbortSignal): Promise<string>
  usage(modelCallId: string, usage: AgentTokenUsage): Promise<void>
  beforeApproval?(
    call: AgentToolCall,
    required?: boolean,
    parentCall?: AgentToolCall,
  ): Promise<void | TaskApprovalScope>
  afterApproval?(call: AgentToolCall, allowed: boolean, round?: boolean): Promise<void>
  prepareTool(call: AgentToolCall, signal: AbortSignal): Promise<TaskToolPermit>
  finishTool(
    permit: TaskToolPermit,
    result: ToolExecution,
    checkpoint: AgentSessionSnapshot,
  ): Promise<void>
  /** Undefined when the task has no on-disk deliverable requirement. */
  checkSaved?(): Promise<boolean | undefined>
  pendingWordParts?(): Promise<string[]>
  finish(result: TaskExecutionEnd, checkpoint: AgentSessionSnapshot): Promise<void>
}

/** Build a provider-valid prefix when a crash occurs in a multi-call model turn. */
export function taskMessagePrefix(messages: readonly AgentMessage[]): AgentMessage[] {
  const result: AgentMessage[] = []
  for (let index = 0; index < messages.length; index++) {
    const message = messages[index]!
    if (message.role === 'tool') continue
    if (message.role !== 'assistant' || !message.toolCalls?.length) {
      result.push(message)
      continue
    }
    const next = messages[index + 1]
    if (next?.role !== 'tool') {
      if (message.text) result.push({ role: 'assistant', text: message.text })
      continue
    }
    const results = new Map(next.results.map((item) => [item.id, item]))
    const calls = message.toolCalls.filter((call) => results.has(call.id))
    if (calls.length) {
      result.push({ ...message, toolCalls: calls })
      result.push({ role: 'tool', results: calls.map((call) => results.get(call.id)!) })
    } else if (message.text) result.push({ role: 'assistant', text: message.text })
    index++
  }
  return result
}
