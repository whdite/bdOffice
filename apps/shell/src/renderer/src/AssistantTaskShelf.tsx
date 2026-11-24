import { useEffect, useState } from 'react'
import { officeTaskCopy, PortableOfficeResults, TaskQuestionReply } from '@genoffice/ui'
import type { PortableResultRecord } from '@genoffice/project-store'
import type {
  AssistantTaskControl,
  AssistantTaskFeed,
  AssistantTasksApi,
  AssistantTaskView,
} from '../../shared/assistant-task-api'

declare global {
  interface Window {
    aiOfficeTasks?: AssistantTasksApi
  }
}

export interface TaskActivity {
  running: number
  waiting: number
  attention: number
  total: number
  unread: number
}

export function AssistantTaskShelf({
  lang,
  embedded = false,
  visible = true,
  onActivityChange,
  onNavigate,
  sessionId,
  onOpenFile,
}: {
  lang: string
  sessionId?: string
  onOpenFile?(path: string): void
  embedded?: boolean
  visible?: boolean
  onActivityChange?(activity: TaskActivity): void
  onNavigate?(): void
}) {
  const api = window.aiOfficeTasks,
    copy = officeTaskCopy(lang)
  const [open, setOpen] = useState(false),
    [feed, setFeed] = useState<AssistantTaskFeed>({ tasks: [], notices: [], throughCursor: 0 })
  const expanded = !!sessionId || (embedded ? visible : open)
  const [loaded, setLoaded] = useState(false),
    [feedError, setFeedError] = useState(false)
  const [working, setWorking] = useState(''),
    [error, setError] = useState('')
  const [results, setResults] = useState<{
    taskId: string
    records: PortableResultRecord[]
  } | null>(null)
  useEffect(() => {
    if (!api) return
    let alive = true,
      pending = false
    const poll = async () => {
      if (pending || document.visibilityState === 'hidden') return
      pending = true
      try {
        const value = await api.feed(sessionId)
        if (alive) {
          setFeed(value)
          setLoaded(true)
          setFeedError(false)
        }
      } catch {
        if (alive) setFeedError(true)
      } finally {
        pending = false
      }
    }
    void poll()
    const timer = setInterval(() => void poll(), expanded ? 700 : 2_000)
    return () => {
      alive = false
      clearInterval(timer)
    }
  }, [api, expanded, sessionId])
  const needsAttention = (task: AssistantTaskView) =>
    task.state === 'waiting_user' ||
    task.state === 'interrupted' ||
    ['outcome_unknown', 'task_context_changed', 'task_unsaved_context_unavailable'].includes(
      task.reason ?? '',
    ) ||
    (task.state === 'failed' && feed.notices.some((notice) => notice.taskId === task.id))
  const running = feed.tasks.filter((task) =>
    ['preparing', 'running', 'verifying', 'pausing', 'cancelling'].includes(task.state),
  ).length
  const waiting = feed.tasks.filter((task) =>
    ['queued', 'waiting_resource', 'waiting_device'].includes(task.state),
  ).length
  const attention = feed.tasks.filter(needsAttention).length
  useEffect(() => {
    onActivityChange?.({
      running,
      waiting,
      attention,
      total: feed.tasks.length,
      unread: feed.notices.length,
    })
  }, [running, waiting, attention, feed.tasks.length, feed.notices.length, onActivityChange])
  const run = async (task: AssistantTaskView, action: AssistantTaskControl['action']) => {
    if (!api || working) return
    setWorking(task.id)
    setError('')
    try {
      const result = await api.control({
        taskId: task.id,
        requestId: crypto.randomUUID(),
        expectedRevision: task.revision,
        action,
        reviewed: action === 'reconcile',
        ...(action === 'reconcile' ? { continueAfterReview: true } : {}),
      })
      setFeed((current) => ({
        ...current,
        tasks: current.tasks.map((value) => (value.id === result.id ? result : value)),
      }))
    } catch (error) {
      setError(
        copy.reason(error instanceof Error ? error.message.split(/[: ]/).at(-1) : undefined) ??
          copy.unavailable,
      )
    } finally {
      setWorking('')
    }
  }
  if (!api) return embedded && visible ? <p role="alert">{copy.unavailable}</p> : null
  if (sessionId) {
    const tasks = feed.tasks.filter(
      (task) =>
        task.sessionId === sessionId &&
        (task.reviewTargets ||
          ['paused', 'interrupted', 'failed'].includes(task.state) ||
          task.confirmation?.kind === 'clarification'),
    )
    return (
      <>
        {tasks.slice(0, 8).map((task) => (
          <section className="office-task-item" key={task.id}>
            <strong>{task.title}</strong> · {copy.status(task.state)}
            {task.reason && <p>{copy.reason(task.reason)}</p>}
            {task.reviewTargets?.map((path) => (
              <p key={path}>
                {onOpenFile ? (
                  <button type="button" onClick={() => onOpenFile(path)}>
                    {path}
                  </button>
                ) : (
                  path
                )}
              </p>
            ))}
            {task.steps
              .filter((step) => step.state === 'awaiting_reconciliation')
              .map((step) => (
                <code key={step.id}>{step.label}</code>
              ))}
            {task.confirmation?.kind === 'clarification' ? (
              <TaskQuestionReply
                question={task.confirmation.text}
                lang={lang}
                answer={async (answer) => {
                  await api.answer({
                    taskId: task.id,
                    confirmationId: task.confirmation!.id,
                    expectedRevision: task.revision,
                    requestId: crypto.randomUUID(),
                    answer,
                  })
                }}
              />
            ) : (
              <button
                type="button"
                disabled={!!working}
                onClick={() =>
                  void run(
                    task,
                    task.reviewTargets ? 'reconcile' : task.state === 'failed' ? 'retry' : 'resume',
                  )
                }
              >
                {task.reviewTargets
                  ? task.pendingControl === 'cancel'
                    ? copy.acknowledge
                    : copy.reconcile
                  : task.state === 'failed'
                    ? copy.retry
                    : copy.resume}
              </button>
            )}
          </section>
        ))}
        {error && <p role="alert">{error}</p>}
      </>
    )
  }
  const actionable = feed.tasks.filter(
    (task) => !['completed', 'failed', 'cancelled'].includes(task.state),
  ).length
  const priority = (task: AssistantTaskView) =>
    needsAttention(task) ? 0 : !['completed', 'failed', 'cancelled'].includes(task.state) ? 1 : 2
  const content = expanded ? (
    <div>
      {feed.notices.length > 0 && (
        <div className="office-task-actions">
          <button
            type="button"
            disabled={!!working}
            onClick={() => {
              const ids = new Set(feed.notices.map((notice) => notice.id))
              setWorking('acknowledging')
              void Promise.all([...ids].map((id) => api.acknowledge(id)))
                .then(() =>
                  setFeed((current) => ({
                    ...current,
                    notices: current.notices.filter((notice) => !ids.has(notice.id)),
                  })),
                )
                .catch(() => setError(copy.unavailable))
                .finally(() => setWorking(''))
            }}
          >
            {lang.startsWith('zh') ? '全部标为已读' : 'Mark all as read'}
          </button>
        </div>
      )}
      {feedError && <p role="alert">{copy.unavailable}</p>}
      {!loaded && !feedError && (
        <p role="status">{lang.startsWith('zh') ? '正在读取任务…' : 'Loading tasks…'}</p>
      )}
      {loaded && !feedError && !feed.tasks.length && (
        <p>{lang.startsWith('zh') ? '还没有任务记录。' : 'No tasks yet.'}</p>
      )}
      {[...feed.tasks]
        .sort((a, b) => priority(a) - priority(b) || b.updatedAt - a.updatedAt)
        .slice(0, 30)
        .map((task) => (
          <div key={task.id} className="office-task-item" data-task-state={task.state}>
            <strong>{task.title}</strong> <small>{copy.status(task.state)}</small>
            {copy.reason(task.reason) && <p>{copy.reason(task.reason)}</p>}
            {task.confirmation && <p>{copy.status('waiting_user')}</p>}
            <div className="office-task-actions">
              <button
                type="button"
                onClick={() =>
                  void api
                    .openSession(task.id)
                    .then(() => onNavigate?.())
                    .catch(() => setError(copy.unavailable))
                }
              >
                {lang.startsWith('zh') ? '查看对话' : 'Open conversation'}
              </button>
              {task.evidenceIds.length > 0 && (
                <button
                  type="button"
                  onClick={() =>
                    void api
                      .evidence(task.id)
                      .then((records) => setResults({ taskId: task.id, records }))
                      .catch(() => setError(copy.unavailable))
                  }
                >
                  {lang.startsWith('zh') ? '查看成果' : 'View results'}
                </button>
              )}
              {feed.notices.some((item) => item.taskId === task.id) && (
                <button
                  type="button"
                  onClick={() => {
                    void Promise.all(
                      feed.notices
                        .filter((item) => item.taskId === task.id)
                        .map((item) => api.acknowledge(item.id)),
                    )
                      .then(() =>
                        setFeed((current) => ({
                          ...current,
                          notices: current.notices.filter((item) => item.taskId !== task.id),
                        })),
                      )
                      .catch(() => setError(copy.unavailable))
                  }}
                >
                  {copy.acknowledge}
                </button>
              )}
            </div>
            {results?.taskId === task.id && (
              <PortableOfficeResults records={results.records} lang={lang} />
            )}
          </div>
        ))}
      {error && <p role="alert">{error}</p>}
    </div>
  ) : null
  if (embedded)
    return (
      <section
        className="office-task-panel assistant-task-records"
        aria-label={lang.startsWith('zh') ? '所有对话的任务' : 'Tasks across conversations'}
      >
        {content}
      </section>
    )
  if (!feed.tasks.length && !open) return null
  return (
    <details
      className="office-task-panel assistant-task-shelf"
      onToggle={(event) => setOpen(event.currentTarget.open)}
    >
      <summary>
        {copy.localTasks}
        {actionable ? ` · ${actionable}` : ''}
        {feed.notices.length ? ` · ${copy.notice}` : ''}
      </summary>
      {content}
    </details>
  )
}
