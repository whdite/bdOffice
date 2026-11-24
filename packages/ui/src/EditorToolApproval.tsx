import { editorToolCapability, isObservation } from '@genoffice/agent-core'
import { useEffect, useSyncExternalStore } from 'react'
import { EditorApprovalControl } from './EditorApprovalControl'
import './editor-tool-approval.css'

interface Call {
  id: string
  name: string
  input: Record<string, unknown>
}
interface Pending {
  id: string
  call: Call
  round?: { path: string; baseVersion: string }
  decide(allow: boolean | 'allow-round'): void
}
interface ApprovalBridge {
  getMode(): Promise<'auto' | 'per-action'>
  setMode?(mode: 'auto' | 'per-action'): Promise<'auto' | 'per-action'>
  onModeChanged(listener: (mode: 'auto' | 'per-action') => void): () => void
}
declare global {
  interface Window {
    editorApproval?: ApprovalBridge
  }
}

export const isEditorReadTool = (name: string): boolean =>
  editorToolCapability({ name, input: {} }) === 'read'
let pending: readonly Pending[] = []
const listeners = new Set<() => void>()
const publish = () => {
  for (const listener of listeners) listener()
}
const subscribe = (listener: () => void) => {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}
const snapshot = () => pending

export async function requestEditorToolApproval(
  call: Call,
  signal?: AbortSignal,
  context?: { editor?: string; round?: { path: string; baseVersion: string } },
): Promise<boolean | 'allow-round'> {
  if (signal?.aborted) return false
  if (isObservation(editorToolCapability(call, context?.editor))) return true
  const bridge = window.editorApproval
  if (!bridge) throw new Error('Editor approval bridge is unavailable; the tool was not executed.')
  const mode = await bridge.getMode()
  if (signal?.aborted) return false
  if (mode === 'auto') return true
  return new Promise<boolean | 'allow-round'>((resolve) => {
    const id = crypto.randomUUID()
    let settled = false
    let unsubscribe = () => {}
    const decide = (allow: boolean | 'allow-round') => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      signal?.removeEventListener('abort', abort)
      unsubscribe()
      pending = pending.filter((request) => request.id !== id)
      publish()
      resolve(signal?.aborted ? false : allow)
    }
    const abort = () => decide(false)
    const timer = setTimeout(abort, 5 * 60_000)
    signal?.addEventListener('abort', abort, { once: true })
    try {
      unsubscribe = bridge.onModeChanged((next) => {
        if (next === 'auto') decide(true)
      })
    } catch {
      abort()
      return
    }
    if (signal?.aborted) {
      abort()
      return
    }
    pending = [
      ...pending,
      { id, call, decide, ...(context?.round ? { round: context.round } : {}) },
    ]
    publish()
    // Close the race between the first mode read and registering the listener.
    void bridge
      .getMode()
      .then((next) => {
        if (next === 'auto') decide(true)
      })
      .catch(() => decide(false))
  })
}

/** Inline, not a modal: stays beside the assistant's composer/activity stream. */
export function EditorToolApproval() {
  const requests = useSyncExternalStore(subscribe, snapshot, snapshot)
  useEffect(
    () => () => {
      for (const request of [...pending]) request.decide(false)
    },
    [],
  )
  const zh =
    typeof document !== 'undefined' &&
    (document.documentElement.lang || navigator.language).startsWith('zh')
  return (
    <>
      <EditorApprovalControl zh={zh} />
      {requests.map((request) => (
        <section
          className="editor-tool-approval"
          role="region"
          aria-label={zh ? '工具执行审批' : 'Tool approval'}
          key={request.id}
        >
          <div className="editor-tool-approval-title">
            {zh ? '等待批准此操作' : 'Approve this action'}
          </div>
          <code>{request.call.name}</code>
          <details>
            <summary>{zh ? '查看参数' : 'View arguments'}</summary>
            <pre>{JSON.stringify(request.call.input, null, 2)}</pre>
          </details>
          {request.round && (
            <small>
              {request.round.path} · {request.round.baseVersion.slice(0, 12)}
            </small>
          )}
          <div className="editor-tool-approval-actions">
            <button type="button" onClick={() => request.decide(false)}>
              {zh ? '拒绝' : 'Deny'}
            </button>
            <button type="button" onClick={() => request.decide(true)}>
              {zh ? '允许本次' : 'Allow once'}
            </button>
            {request.round && (
              <button type="button" onClick={() => request.decide('allow-round')}>
                {zh ? '允许本轮' : 'Allow this round'}
              </button>
            )}
          </div>
        </section>
      ))}
    </>
  )
}
