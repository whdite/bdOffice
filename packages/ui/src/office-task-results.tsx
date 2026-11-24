import { useEffect, useRef, useState } from 'react'
import type { ConversationApi, PortableResultRecord } from '@genoffice/project-store'
import {
  selectFileTaskContinuation,
  clearFileTaskContinuation,
  type ChatIdentityRef,
  type FileTaskItem,
  type FileTaskPanelState,
} from '@genoffice/project-store/browser'
import './office-task-results.css'

export function officeTaskCopy(lang: string) {
  const zh = lang.startsWith('zh')
  const states: Record<string, [string, string]> = {
    queued: ['排队中', 'Queued'],
    waiting_device: ['等待打开原文档', 'Waiting for the editor'],
    waiting_resource: ['等待其他任务', 'Waiting for another task'],
    preparing: ['准备中', 'Preparing'],
    running: ['正在处理', 'Running'],
    waiting_user: ['等待确认', 'Waiting for confirmation'],
    verifying: ['正在核验', 'Verifying'],
    pausing: ['正在安全暂停', 'Pausing safely'],
    paused: ['已暂停', 'Paused'],
    cancelling: ['正在安全停止', 'Stopping safely'],
    interrupted: ['已中断', 'Interrupted'],
    completed: ['已完成', 'Completed'],
    failed: ['需要处理', 'Needs attention'],
    cancelled: ['已停止', 'Stopped'],
    partial: ['部分完成', 'Partially completed'],
    undone: ['已撤销', 'Undone'],
    'undo-failed': ['撤销需要处理', 'Undo needs attention'],
  }
  const reasons: Record<string, [string, string]> = {
    approval_expired: [
      '操作确认已过期，这次操作未执行。可重试并重新确认。',
      'Approval expired before execution. Retry for a new confirmation.',
    ],
    operation_not_completed: [
      '请求中的操作未完成，请核对对话中的原因后重试。',
      'The requested operation did not complete. Review the conversation before retrying.',
    ],
    task_provider_transport_failed: [
      '连接模型服务失败，请检查网络和模型连接后重试。已完成的工具步骤会先核对再继续。',
      'The model connection failed. Check the network and model connection, then retry. Completed tool steps will be checked before continuing.',
    ],
    execution_failed: [
      '任务执行未完成，请查看对话中的错误信息后重试。',
      'The task did not finish. Review the conversation error before retrying.',
    ],
    outcome_unknown: [
      '上一次操作的结果尚未确认。请核对应用或文件的当前状态，并确认旧执行已经停止。',
      'The previous action has an uncertain result. Review the current app or file state and make sure the old execution has stopped.',
    ],
    task_resource_requires_review: [
      '该文件或操作在另一项任务中有待核对的结果。请打开该任务的对话。',
      'This file or action has a pending review in another task. Open that conversation to continue.',
    ],
    task_effect_requires_review: [
      '上一次操作需要先核对。',
      'Review the previous action before continuing.',
    ],
    task_context_changed: [
      '文件已有其他修改，请核对后继续。',
      'The file has changed. Review the current version before continuing.',
    ],
    task_unsaved_context_unavailable: [
      '原来未保存的编辑内容已不可用，请核对当前文档。',
      'The original unsaved content is unavailable. Review the current document.',
    ],
    task_provider_changed: [
      '当前模型配置与开始时不同，请检查连接。',
      'The model configuration changed. Check the connection.',
    ],
    result_not_verified: [
      '请求的成果尚未核验通过。',
      'The requested result has not passed verification.',
    ],
    budget_exhausted: ['任务用量已达上限。', 'The task reached its execution limit.'],
    task_budget_exhausted: ['任务用量已达上限。', 'The task reached its execution limit.'],
    fresh_confirmation_required: [
      '恢复后需要重新确认操作。',
      'The action needs a fresh confirmation.',
    ],
    task_conversation_changed: [
      '此对话已有后续任务，旧任务不会覆盖新上下文。',
      'This conversation has newer work. Its context will be preserved.',
    ],
  }
  return {
    status: (state: string) => states[state]?.[zh ? 0 : 1] ?? (zh ? '需要处理' : 'Needs attention'),
    reason: (reason: string | undefined) => (reason ? reasons[reason]?.[zh ? 0 : 1] : undefined),
    tasks: zh ? '任务和成果' : 'Tasks and results',
    localTasks: zh ? '本机任务' : 'Local tasks',
    before: zh ? '修改前' : 'Before',
    after: zh ? '修改后' : 'After',
    pause: zh ? '暂停' : 'Pause',
    resume: zh ? '继续' : 'Continue',
    retry: zh ? '重试' : 'Retry',
    cancel: zh ? '停止' : 'Stop',
    reviewed: zh ? '已核对结果并确认旧执行停止' : 'Outcome reviewed and previous execution stopped',
    reconcile: zh ? '已核对，继续任务' : 'Reviewed, continue task',
    resumeMessage: (title: string) => (zh ? `继续任务：${title}` : `Continue task: ${title}`),
    imported: zh ? '迁移的历史成果' : 'Imported result history',
    saved: zh ? '已保存' : 'Saved',
    pending: zh ? '待保存' : 'Pending save',
    memory: zh ? '仅在编辑器中' : 'In editor only',
    unknown: zh ? '待核实' : 'Needs verification',
    verified: zh ? '已核验' : 'Verified',
    sourceVerified: zh ? '来源记录已核验' : 'Verified on the source device',
    editorVerified: zh ? '编辑器内容已核验' : 'Editor content verified',
    open: zh ? '打开成果副本' : 'Open result copy',
    limited: zh ? '此处仅展示可核对的差异。' : 'Only observed differences are shown here.',
    unavailable: zh ? '任务记录暂不可用，请重试。' : 'Task history is unavailable. Try again.',
    empty: zh ? '暂无任务' : 'No tasks yet',
    acknowledge: zh ? '已读' : 'Mark read',
    notice: zh ? '有任务结果或需要处理的事项' : 'Task results or items need your attention',
    block: zh ? '文档块' : 'Document block',
    slide: zh ? '幻灯片' : 'Slide',
    page: zh ? '页' : 'Page',
    line: zh ? '行' : 'Line',
    reply: zh ? '回复并继续' : 'Reply and continue',
    replyLabel: zh ? '任务回复' : 'Task reply',
  }
}
export function TaskQuestionReply({
  question,
  lang,
  answer,
}: {
  question: string
  lang: string
  answer(text: string): Promise<void>
}) {
  const copy = officeTaskCopy(lang),
    [value, setValue] = useState(''),
    [busy, setBusy] = useState(false),
    [failed, setFailed] = useState(false)
  const submit = async () => {
    if (busy || !value.trim()) return
    setBusy(true)
    setFailed(false)
    try {
      await answer(value.trim())
      setValue('')
    } catch {
      setFailed(true)
    } finally {
      setBusy(false)
    }
  }
  return (
    <form
      className="office-task-question"
      onSubmit={(event) => {
        event.preventDefault()
        void submit()
      }}
    >
      <p>{question}</p>
      <textarea
        aria-label={copy.replyLabel}
        value={value}
        maxLength={8000}
        onChange={(event) => setValue(event.target.value)}
      />
      <button type="submit" disabled={busy || !value.trim()}>
        {copy.reply}
      </button>
      {failed && <p role="alert">{copy.unavailable}</p>}
    </form>
  )
}
export function PortableOfficeResults({
  records,
  lang,
  openArtifact,
}: {
  records: PortableResultRecord[]
  lang: string
  openArtifact?: ((id: string) => void) | undefined
}) {
  const copy = officeTaskCopy(lang)
  return (
    <>
      {records.map((record) => (
        <details
          key={record.id}
          className="office-result-record"
          data-result-origin={record.origin}
        >
          <summary>
            {record.origin === 'imported' ? copy.imported : copy.status(record.outcome)} ·{' '}
            {new Date(record.occurredAt).toLocaleString()}
          </summary>
          {record.files.map((file) => (
            <section key={file.id} className="office-result-file">
              <strong>{file.name}</strong>
              <div className="office-result-status">
                <span>
                  {file.saveState === 'saved'
                    ? copy.saved
                    : file.saveState === 'pending_save'
                      ? copy.pending
                      : file.saveState === 'memory_only'
                        ? copy.memory
                        : copy.unknown}
                </span>
                <span>
                  {file.verification === 'passed'
                    ? record.origin === 'imported'
                      ? copy.sourceVerified
                      : file.scope === 'editor_model'
                        ? copy.editorVerified
                        : copy.verified
                    : copy.unknown}
                </span>
              </div>
              {file.artifactId && openArtifact && (
                <button type="button" onClick={() => openArtifact(file.artifactId!)}>
                  {copy.open}
                </button>
              )}
              {file.details.map((detail) => {
                const at = detail.location
                const label =
                  at.kind === 'cell'
                    ? `${at.sheet ?? ''}!${at.range ?? ''}`
                    : at.kind === 'block'
                      ? `${copy.block} ${at.index ?? ''}`
                      : at.kind === 'slide'
                        ? `${copy.slide} ${at.slide ?? ''}`
                        : at.kind === 'page'
                          ? `${copy.page} ${at.page ?? ''}`
                          : at.kind === 'line'
                            ? `${copy.line} ${at.line ?? ''}`
                            : (at.name ?? at.sheet ?? file.name)
                return (
                  <div key={detail.id} className="office-result-difference">
                    <small>{label}</small>
                    <div>
                      <div>
                        <small>{copy.before}</small>
                        <pre>{detail.before ?? '—'}</pre>
                      </div>
                      <div>
                        <small>{copy.after}</small>
                        <pre>{detail.after ?? '—'}</pre>
                      </div>
                    </div>
                  </div>
                )
              })}
              {file.incomplete && <p>{copy.limited}</p>}
            </section>
          ))}
        </details>
      ))}
    </>
  )
}

/** All six editor panels share the same task controls and historical evidence view. */
export function FileTaskPanel({
  api,
  ids,
  lang,
  busy,
  resume,
}: {
  api?: ConversationApi
  ids: ChatIdentityRef
  lang: string
  busy: boolean
  resume(instruction: string): void
}) {
  const copy = officeTaskCopy(lang),
    call = api?.fileTask
  const [state, setState] = useState<FileTaskPanelState>({ tasks: [], results: [] })
  const [error, setError] = useState(''),
    [working, setWorking] = useState(false)
  const lastAuto = useRef(''),
    currentResume = useRef(resume)
  currentResume.current = resume
  useEffect(() => {
    if (!call) return
    let alive = true,
      pending = false,
      unavailable = false
    const poll = async () => {
      const source = ids.current
      if (!source || pending || unavailable || document.visibilityState === 'hidden') return
      pending = true
      try {
        const next = await call('panel', source)
        if (alive && ids.current?.contextToken === source.contextToken) setState(next)
      } catch (error) {
        if (error instanceof Error && error.message.includes('No handler registered'))
          unavailable = true
      } finally {
        pending = false
      }
    }
    void poll()
    const timer = setInterval(() => void poll(), 1_100)
    return () => {
      alive = false
      clearInterval(timer)
      clearFileTaskContinuation(ids)
    }
  }, [call, ids])
  const next = state.tasks.find(
    (task) =>
      task.resumeRequested && ['queued', 'waiting_device', 'waiting_resource'].includes(task.state),
  )
  const key = next ? `${next.id}:${next.attempts}` : ''
  useEffect(() => {
    if (!key || busy || !next || lastAuto.current === key || !ids.current) return
    lastAuto.current = key
    selectFileTaskContinuation(ids, next.id)
    currentResume.current(officeTaskCopy(lang).resumeMessage(next.title))
  }, [key, busy, next, ids, lang])
  const act = async (
    task: FileTaskItem,
    action: 'pause' | 'resume' | 'cancel' | 'retry' | 'reconcile',
  ) => {
    if (!call || !ids.current || working) return
    setWorking(true)
    setError('')
    try {
      if (
        action === 'resume' &&
        task.resumeRequested &&
        ['queued', 'waiting_device', 'waiting_resource'].includes(task.state)
      ) {
        selectFileTaskContinuation(ids, task.id)
        resume(copy.resumeMessage(task.title))
        return
      }
      const result = await call('control', {
        ...ids.current,
        taskId: task.id,
        requestId: crypto.randomUUID(),
        expectedRevision: task.revision,
        action,
        reviewed: action === 'reconcile',
        ...(action === 'reconcile' ? { continueAfterReview: true } : {}),
      })
      setState((current) => ({
        ...current,
        tasks: current.tasks.map((value) => (value.id === result.id ? result : value)),
      }))
    } catch (error) {
      setError(
        copy.reason(error instanceof Error ? error.message.split(/[: ]/).at(-1) : undefined) ??
          copy.unavailable,
      )
    } finally {
      setWorking(false)
    }
  }
  if (!state.tasks.length && !state.results.length) return null
  return (
    <details
      className="office-task-panel"
      data-file-task-panel="true"
      open={state.tasks.some((task) => !!task.reviewTargets) || undefined}
    >
      <summary>
        {copy.tasks}
        {state.tasks[0] ? ` · ${copy.status(state.tasks[0].state)}` : ''}
      </summary>
      {[...state.tasks]
        .sort((a, b) => Number(!!b.reviewTargets) - Number(!!a.reviewTargets))
        .slice(0, 8)
        .map((task) => (
          <div key={task.id} className="office-task-item" data-task-state={task.state}>
            <strong>{task.title}</strong> <small>{copy.status(task.state)}</small>
            {copy.reason(task.reason) && <p>{copy.reason(task.reason)}</p>}
            {task.confirmation?.kind === 'clarification' && call && (
              <TaskQuestionReply
                key={task.confirmation.id}
                question={task.confirmation.text}
                lang={lang}
                answer={async (answer) => {
                  if (!ids.current) throw new Error('Conversation unavailable')
                  const result = await call('answer', {
                    ...ids.current,
                    taskId: task.id,
                    confirmationId: task.confirmation!.id,
                    requestId: crypto.randomUUID(),
                    expectedRevision: task.revision,
                    answer,
                  })
                  setState((current) => ({
                    ...current,
                    tasks: current.tasks.map((item) => (item.id === result.id ? result : item)),
                  }))
                }}
              />
            )}
            <div className="office-task-actions">
              {['queued', 'waiting_resource', 'running', 'waiting_user'].includes(task.state) && (
                <button disabled={working} type="button" onClick={() => void act(task, 'pause')}>
                  {copy.pause}
                </button>
              )}
              {(['paused', 'interrupted', 'waiting_device'].includes(task.state) &&
                task.reason !== 'outcome_unknown') ||
              (task.resumeRequested && task.state === 'queued') ? (
                <button
                  disabled={working || busy}
                  type="button"
                  onClick={() => void act(task, 'resume')}
                >
                  {copy.resume}
                </button>
              ) : null}
              {task.state === 'failed' && (
                <button
                  disabled={working || busy}
                  type="button"
                  onClick={() => void act(task, 'retry')}
                >
                  {copy.retry}
                </button>
              )}
              {!['completed', 'failed', 'cancelled'].includes(task.state) && (
                <button disabled={working} type="button" onClick={() => void act(task, 'cancel')}>
                  {copy.cancel}
                </button>
              )}
              {[
                'outcome_unknown',
                'task_context_changed',
                'task_unsaved_context_unavailable',
              ].includes(task.reason ?? '') && (
                <>
                  {task.reviewTargets?.map((path) => (
                    <p key={path}>{path}</p>
                  ))}
                  <button
                    type="button"
                    disabled={working}
                    onClick={() => void act(task, 'reconcile')}
                  >
                    {task.pendingControl === 'cancel' ? copy.acknowledge : copy.reconcile}
                  </button>
                </>
              )}
            </div>
          </div>
        ))}
      {error && <p role="alert">{error}</p>}
      <PortableOfficeResults
        records={state.results}
        lang={lang}
        openArtifact={
          call
            ? (artifactId) => {
                if (ids.current)
                  void call('openArtifact', { ...ids.current, artifactId }).catch(() =>
                    setError(copy.unavailable),
                  )
              }
            : undefined
        }
      />
    </details>
  )
}
