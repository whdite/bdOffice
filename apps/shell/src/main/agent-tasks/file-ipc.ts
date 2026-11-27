import { randomUUID } from 'node:crypto'
import type { IpcMain, IpcMainInvokeEvent, WebContents } from 'electron'
import {
  editorToolCapability,
  recoveryForCapability,
  explicitlyRequestsSave,
  type AgentSessionSnapshot,
  type AgentTaskExecution,
  type AgentToolCall,
  type EditorHarnessKind,
  type TaskStartInput,
  type ToolExecution,
} from '@genoffice/agent-core'
import {
  FILE_TASK_CHANNEL,
  normalizeSessionContext,
  type ChatContextController,
  type FileTaskItem,
  type ProjectStore,
  type PortableResultRecord,
} from '@genoffice/project-store'
import type {
  AssistantTaskAnswer,
  AssistantTaskControl,
  AssistantTaskRequest,
  TaskTargetVersion,
} from '../../shared/assistant-task-api'
import type { AssistantTaskService, TaskExecutorBinding } from './service'
import { taskId } from './assistant-store'
import type { DocumentObservation } from './document-observation'

const object = (value: unknown): Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('invalid_file_task_request')
  return value as Record<string, unknown>
}
export function parseAssistantTaskControl(value: unknown): AssistantTaskControl {
  const row = object(value)
  if (
    !['pause', 'resume', 'cancel', 'retry', 'reconcile'].includes(String(row.action)) ||
    typeof row.expectedRevision !== 'number' ||
    !Number.isSafeInteger(row.expectedRevision) ||
    row.expectedRevision < 0
  )
    throw new Error('invalid_task_control')
  return {
    taskId: taskId(row.taskId),
    requestId: taskId(row.requestId),
    expectedRevision: row.expectedRevision,
    action: row.action as AssistantTaskControl['action'],
    ...(row.reviewed === true ? { reviewed: true } : {}),
    ...(row.action === 'reconcile' && row.continueAfterReview === true
      ? { continueAfterReview: true }
      : {}),
  }
}
export function parseAssistantTaskAnswer(value: unknown): AssistantTaskAnswer {
  const row = object(value)
  if (
    typeof row.answer !== 'string' ||
    !row.answer.trim() ||
    row.answer.length > 8000 ||
    typeof row.expectedRevision !== 'number' ||
    !Number.isSafeInteger(row.expectedRevision) ||
    row.expectedRevision < 0
  )
    throw new Error('invalid_task_answer')
  return {
    taskId: taskId(row.taskId),
    confirmationId: taskId(row.confirmationId),
    requestId: taskId(row.requestId),
    expectedRevision: row.expectedRevision,
    answer: row.answer,
  }
}
function proposal(value: unknown): AgentToolCall {
  const row = object(value)
  if (
    typeof row.name !== 'string' ||
    !/^[a-zA-Z0-9_.:-]{1,128}$/.test(row.name) ||
    typeof row.id !== 'string' ||
    row.id.length > 256
  )
    throw new Error('invalid_task_tool')
  const input = object(row.input)
  if (Buffer.byteLength(JSON.stringify(input)) > 2 * 1024 * 1024)
    throw new Error('task_payload_too_large')
  return { id: row.id, name: row.name, input }
}
function fileWordEvidence(value: unknown): NonNullable<ToolExecution['wordPart']> {
  const data = object(value),
    result: NonNullable<ToolExecution['wordPart']> = {}
  for (const field of ['expected', 'observed'] as const) {
    if (!data[field]) continue
    const values = object(data[field])
    if (
      Object.values(values).some((v) => typeof v !== 'string') ||
      Buffer.byteLength(JSON.stringify(values)) > 2 * 1024 * 1024
    )
      throw new Error('invalid_word_observation')
    result[field] = values as Record<string, string>
  }
  return result
}

interface Run {
  token: string
  taskId: string
  senderId: number
  sessionId: string
  contextToken: string
  writerId: string
  execution: AgentTaskExecution
  alive(): boolean
}

/** IPC owns the editor and source binding. Requests cannot supply an executor path or lease. */
export function registerFileTaskIpc(options: {
  ipc: IpcMain
  service(): AssistantTaskService
  chats: ChatContextController
  projects(): ProjectStore
  describe(sender: WebContents): { kind: EditorHarnessKind; path?: string } | null
  observe(sender: WebContents, sessionId: string): Promise<DocumentObservation>
  inspect(
    sender: WebContents,
    sessionId: string,
    expected: TaskTargetVersion,
  ): Promise<DocumentObservation>
  results(sessionId: string): Promise<PortableResultRecord[]>
  openArtifact?(sessionId: string, artifactId: string): Promise<void>
  approveEach(): boolean
}): void {
  const runs = new Map<string, Run>(),
    pending = new Map<string, AbortController>()
  const owner = (
    event: IpcMainInvokeEvent,
    row: Record<string, unknown>,
    requireWriter: boolean,
  ) => {
    const bound = options.chats.current(event.sender.id),
      editor = options.describe(event.sender)
    if (
      !bound ||
      !editor ||
      row.projectId !== bound.ids.projectId ||
      row.chatId !== bound.ids.chatId ||
      row.contextToken !== bound.ids.contextToken ||
      (requireWriter && (!bound.writerId || row.writerId !== bound.writerId))
    )
      throw new Error('task_conversation_binding_changed')
    const meta = options
      .projects()
      .managedChat(bound.ids.projectId, bound.ids.chatId, editor.path, true)
    if (!meta || meta.state !== 'active' || meta.kind !== 'file')
      throw new Error('task_session_inactive')
    return { meta, bound, editor }
  }
  const view = (service: AssistantTaskService, id: string): FileTaskItem => {
    const task = service.store.get(id)
    return {
      id,
      title: task.title,
      state: task.state,
      revision: task.revision,
      reason: task.reason,
      pendingControl: task.pendingControl,
      confirmation: task.confirmation,
      ...(task.reviewTargets ? { reviewTargets: task.reviewTargets } : {}),
      attempts: task.budget.usedAttempts,
      resumeRequested: !!service.store.plan(id).resumeRequested,
    }
  }
  options.ipc.handle(FILE_TASK_CHANNEL, async (event, action: unknown, input: unknown) => {
    if (
      event.sender.isDestroyed() ||
      event.senderFrame !== event.sender.mainFrame ||
      !options.describe(event.sender)
    )
      throw new Error('task_requires_trusted_editor')
    const row = object(input),
      service = options.service()
    if (
      action === 'probe' ||
      action === 'panel' ||
      action === 'control' ||
      action === 'answer' ||
      action === 'openArtifact'
    ) {
      const { meta } = owner(event, row, false)
      if (action === 'probe') return true
      if (action === 'panel')
        return {
          tasks: service.store
            .forSession(meta.id)
            .slice(0, 40)
            .map((task) => view(service, task.id)),
          results: await options.results(meta.id),
        }
      if (action === 'openArtifact') {
        await options.openArtifact?.(meta.id, taskId(row.artifactId))
        return
      }
      if (action === 'answer') {
        const reply = parseAssistantTaskAnswer(row)
        if (service.store.get(reply.taskId).sessionId !== meta.id)
          throw new Error('task_conversation_binding_changed')
        return view(service, service.answer(reply).task.id)
      }
      const control = parseAssistantTaskControl(row)
      if (service.store.get(control.taskId).sessionId !== meta.id)
        throw new Error('task_conversation_binding_changed')
      return view(service, (await service.control(control)).id)
    }
    if (action === 'begin' || action === 'abortBegin') {
      const { meta, bound, editor } = owner(event, row, true),
        requestId = taskId(row.requestId)
      const key = `${event.sender.id}:${bound.writerId}:${requestId}`
      if (action === 'abortBegin') {
        pending.get(key)?.abort()
        return
      }
      const controller = new AbortController()
      if (pending.has(key)) throw new Error('task_already_accepted')
      pending.set(key, controller)
      const sender = event.sender,
        contextToken = bound.ids.contextToken!,
        writerId = bound.writerId!,
        sessionId = meta.id
      const unchangedBinding = () => {
        if (sender.isDestroyed()) return false
        const next = options.chats.current(sender.id),
          currentEditor = options.describe(sender)
        if (!next || !currentEditor) return false
        if (next.ids.contextToken === contextToken && next.writerId === writerId) return true
        // The first Save may move a temporary chat to its saved file identity.
        // Preserve the run only if the catalog confirms the same managed Session.
        if (!editor.path && currentEditor.path && !next.writerId) {
          const current = options
            .projects()
            .managedChat(next.ids.projectId, next.ids.chatId, currentEditor.path)
          return current?.id === sessionId
        }
        return false
      }
      const binding: TaskExecutorBinding = {
        alive: () => !controller.signal.aborted && unchangedBinding(),
        stop: () => {
          if (!sender.isDestroyed())
            sender.send('project:cancel-chat', options.chats.current(sender.id)?.ids ?? bound.ids)
        },
        classify: (call) => recoveryForCapability(editorToolCapability(call, editor.kind)),
        approveEach: options.approveEach,
        observe: async () => [await options.observe(sender, sessionId)],
        inspectTarget: (expected) => options.inspect(sender, sessionId, expected),
      }
      try {
        const value = object(row.input),
          snapshot = normalizeSessionContext(value.snapshot, 'file')
        const goal = typeof value.goal === 'string' ? value.goal : ''
        if (!goal.trim() || goal.length > 128_000) throw new Error('invalid_task_goal')
        const start: TaskStartInput = { goal, snapshot, budget: snapshot.budget }
        const observation = await options.observe(sender, sessionId)
        const request: AssistantTaskRequest = {
          requestId,
          sessionId,
          entry: 'file',
          input: { text: goal },
          document: observation.version,
          requiredResult: explicitlyRequestsSave(goal) ? 'saved_artifact' : 'answer',
        }
        let resumeId = row.resumeId === undefined ? undefined : taskId(row.resumeId)
        const question =
          !resumeId &&
          service.store
            .forSession(sessionId)
            .find(
              (task) =>
                task.state === 'waiting_user' && task.confirmation?.kind === 'clarification',
            )
        if (question && question.confirmation) {
          service.answer(
            {
              taskId: question.id,
              confirmationId: question.confirmation.id,
              requestId,
              expectedRevision: question.revision,
              answer: goal,
            },
            false,
          )
          resumeId = question.id
        }
        const result = await service.beginFile(request, binding, start, resumeId, controller.signal)
        const token = randomUUID()
        // Keep valid active owners; remove abandoned tokens after renderer loss.
        for (const [id, run] of runs) if (!run.alive()) runs.delete(id)
        runs.set(token, {
          token,
          taskId: result.taskId,
          senderId: sender.id,
          sessionId,
          contextToken,
          writerId,
          execution: result.execution,
          alive: unchangedBinding,
        })
        return {
          id: result.taskId,
          token,
          checkpoint: result.checkpoint,
          remainingMs: result.remainingMs,
          memoryContext: result.memoryContext,
        }
      } finally {
        pending.delete(key)
      }
    }
    const token = taskId(row.token),
      run = runs.get(token)
    if (
      !run ||
      run.senderId !== event.sender.id ||
      row.contextToken !== run.contextToken ||
      row.writerId !== run.writerId ||
      !run.alive()
    )
      throw new Error('stale_file_task_binding')
    const execution = run.execution,
      signal = new AbortController().signal
    const checkpoint = () => normalizeSessionContext(row.checkpoint, 'file')
    switch (action) {
      case 'reserveModel':
        return execution.reserveModel(checkpoint(), signal)
      case 'usage':
        return execution.usage(taskId(row.modelCallId), object(row.usage))
      case 'beforeApproval':
        return execution.beforeApproval?.(proposal(row.call))
      case 'afterApproval':
        return execution.afterApproval?.(
          proposal(row.call),
          row.allowed === true,
          row.round === true,
        )
      case 'prepareTool':
        return execution.prepareTool(proposal(row.call), signal)
      case 'finishTool': {
        const permit = object(row.permit),
          value = object(row.result)
        const result: ToolExecution = {
          output: typeof value.output === 'string' ? value.output.slice(0, 100_000) : '',
          summary: typeof value.summary === 'string' ? value.summary.slice(0, 500) : '',
          isError: value.isError === true,
          mutated: value.mutated === true,
          ...(value.exclusiveWordDraft === true ? { exclusiveWordDraft: true } : {}),
          denied: value.denied === true,
          ...(value.externalEffect === 'applied' ||
          value.externalEffect === 'not_applied' ||
          value.externalEffect === 'unknown'
            ? { externalEffect: value.externalEffect }
            : {}),
          ...(value.savedFile &&
          typeof value.savedFile === 'object' &&
          typeof object(value.savedFile).path === 'string' &&
          /^[a-f0-9]{64}$/.test(String(object(value.savedFile).sha256))
            ? { savedFile: value.savedFile as ToolExecution['savedFile'] }
            : {}),
          ...(value.wordPart ? { wordPart: fileWordEvidence(value.wordPart) } : {}),
          ...(value.effect === 'not_applied' ||
          value.effect === 'applied' ||
          value.effect === 'unknown'
            ? { effect: value.effect }
            : {}),
        }
        return execution.finishTool(
          { stepId: taskId(permit.stepId), ...(permit.reused ? { reused: result } : {}) },
          result,
          checkpoint(),
        )
      }
      case 'pendingWordParts':
        return execution.pendingWordParts?.() ?? []
      case 'checkDelivery':
        return execution.checkDelivery?.() ?? []
      case 'checkSaved':
        return execution.checkSaved?.()
      case 'finish': {
        const result = object(row.result)
        if (
          result.state !== 'completed' &&
          result.state !== 'cancelled' &&
          result.state !== 'failed' &&
          result.state !== 'waiting_user'
        )
          throw new Error('invalid_task_result')
        runs.delete(token)
        return execution.finish(
          {
            state: result.state,
            question:
              typeof result.question === 'string' ? result.question.slice(0, 1000) : undefined,
            text: typeof result.text === 'string' ? result.text.slice(0, 20_000) : undefined,
            error: typeof result.error === 'string' ? result.error.slice(0, 2000) : undefined,
          },
          checkpoint() as AgentSessionSnapshot,
        )
      }
      default:
        throw new Error('invalid_file_task_request')
    }
  })
}
