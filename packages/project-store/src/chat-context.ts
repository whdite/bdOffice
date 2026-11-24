import type {
  AgentLoop,
  AgentMessage,
  AgentSessionPersistence,
  AgentSessionSnapshot,
  AgentTaskExecution,
} from '@genoffice/agent-core'
import type { ChatMessage, ConversationApi, ResolveChatResult } from './ipc.js'
import { missingFileTaskChannel } from './file-task-api.js'
export { stripMemoryContextEcho } from './memory-context-text.js'
export type { FileTaskCall, FileTaskItem, FileTaskPanelState } from './file-task-api.js'
export type { PortableResultRecord } from './result-evidence.js'
export interface ChatIdentityRef {
  current: ResolveChatResult | null
}
const continuations = new WeakMap<ChatIdentityRef, string>()
export function selectFileTaskContinuation(ids: ChatIdentityRef, taskId: string): void {
  continuations.set(ids, taskId)
}
export function clearFileTaskContinuation(ids: ChatIdentityRef): void {
  continuations.delete(ids)
}
const INTERRUPTED_NOTE =
  'The previous request was interrupted. No pending tool or approval has been resumed. Inspect current document state before retrying.'
export function recoverConversation(snapshot: AgentSessionSnapshot): AgentSessionSnapshot {
  if (!['interrupted', 'cancelled', 'error'].includes(snapshot.phase)) return snapshot
  const messages = [...snapshot.messages]
  const last = messages.at(-1),
    before = messages.at(-2)
  if (last?.role === 'assistant' && last.text === INTERRUPTED_NOTE && before?.role === 'user')
    return snapshot
  if (snapshot.pendingInput && messages.at(-1)?.role !== 'user')
    messages.push({ role: 'user', text: snapshot.pendingInput })
  if (messages.at(-1)?.role === 'user') messages.push({ role: 'assistant', text: INTERRUPTED_NOTE })
  return { ...snapshot, messages }
}
/** A run captures its identity; late writes never follow a mutable current tab. */
export function createChatPersistence(
  api: () => ConversationApi | undefined,
  ids: ChatIdentityRef,
): AgentSessionPersistence {
  let owner: ResolveChatResult | null = null
  const writerId = crypto.randomUUID()
  let newChat: Promise<void> | null = null
  let saves: Promise<void> = Promise.resolve()
  let taskSupported = false
  const waitForIds = async (signal?: AbortSignal): Promise<ResolveChatResult> => {
    const start = Date.now()
    while (!ids.current) {
      signal?.throwIfAborted()
      if (Date.now() - start > 10_000)
        throw new Error('Conversation storage is not ready. Please try again.')
      await new Promise((resolve) => setTimeout(resolve, 25))
    }
    return { ...ids.current }
  }
  return {
    available: () => !!api()?.loadContext && !!api()?.saveContext,
    async prepare(signal) {
      if (newChat) await newChat
      const current = await waitForIds(signal)
      signal.throwIfAborted()
      owner = current
      const snapshot = await api()!.loadContext!({ ...current, writerId })
      const call = api()?.fileTask
      taskSupported = false
      if (call) {
        try {
          taskSupported = (await call('probe', current)) === true
        } catch (error) {
          if (!missingFileTaskChannel(error)) throw error
        }
      }
      signal.throwIfAborted()
      return snapshot ? recoverConversation(snapshot) : null
    },
    taskExecution(): AgentTaskExecution | undefined {
      const call = api()?.fileTask,
        bound = owner && { ...owner, writerId }
      if (!taskSupported || !call || !bound) return
      let token = ''
      const args = () => ({ ...bound, token })
      return {
        supportsClarification: true,
        async start(input, signal) {
          const requestId = crypto.randomUUID(),
            resumeId = continuations.get(ids)
          continuations.delete(ids)
          const begin = { ...bound, requestId, ...(resumeId ? { resumeId } : {}) }
          const cancel = () => {
            void call('abortBegin', begin).catch(() => {})
          }
          signal.addEventListener('abort', cancel, { once: true })
          try {
            signal.throwIfAborted()
            const result = await call('begin', { ...begin, input })
            token = result.token
            return result
          } finally {
            signal.removeEventListener('abort', cancel)
          }
        },
        reserveModel: (checkpoint, signal) => {
          signal.throwIfAborted()
          return call('reserveModel', { ...args(), checkpoint })
        },
        usage: (modelCallId, usage) => call('usage', { ...args(), modelCallId, usage }),
        beforeApproval: (proposal) => call('beforeApproval', { ...args(), call: proposal }),
        afterApproval: (proposal, allowed, round) =>
          call('afterApproval', {
            ...args(),
            call: proposal,
            allowed,
            ...(round ? { round: true } : {}),
          }),
        prepareTool: (proposal, signal) => {
          signal.throwIfAborted()
          return call('prepareTool', { ...args(), call: proposal })
        },
        finishTool: (permit, result, checkpoint) =>
          call('finishTool', { ...args(), permit, result, checkpoint }),
        pendingWordParts: () => call('pendingWordParts', args()),
        checkSaved: () => call('checkSaved', args()),
        finish: (result, checkpoint) => call('finish', { ...args(), result, checkpoint }),
      }
    },
    save(snapshot) {
      const current = owner ? { ...owner } : null
      if (!current || !api()?.saveContext) return
      const write = () => api()!.saveContext!({ ...current, snapshot, writerId })
      const result = saves.catch(() => {}).then(write)
      saves = result.catch(() => {})
      return result
    },
    async newConversation() {
      if (!api()?.newConversation) return
      const current = await waitForIds()
      continuations.delete(ids)
      ids.current = null
      const change = async () => {
        await saves.catch(() => {})
        try {
          ids.current = await api()!.newConversation!(current)
          owner = null
        } catch (error) {
          ids.current = current
          throw error
        }
      }
      newChat = change()
      try {
        await newChat
      } finally {
        newChat = null
      }
    },
  }
}
export async function restoreChatContext<T>(
  api: ConversationApi,
  ids: ResolveChatResult,
  loop: AgentLoop<T> | null,
  rows: ChatMessage[],
): Promise<void> {
  if (!loop) return
  const fallback: AgentMessage[] = rows.map((m) => ({ role: m.role, text: m.text }))
  if (!api.loadContext) {
    loop.restore(fallback)
    return
  }
  const context = await api.loadContext(ids)
  if (context) loop.restoreSession(recoverConversation(context))
  else loop.restore(fallback)
}
