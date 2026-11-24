import { extname, isAbsolute } from 'node:path'
import { saveWitness } from '@genoffice/electron-utils/save-witness'
import { randomUUID } from 'node:crypto'
import { isRoundWordEdit, sessionBudget } from '@genoffice/agent-core'
import type {
  AgentSessionSnapshot,
  AgentTaskExecution,
  AgentToolCall,
  TaskExecutionEnd,
  TaskStartInput,
  TaskToolPermit,
  ToolExecution,
} from '@genoffice/agent-core'
import type { AgentRunFileChange, AgentRunReview } from '../../shared/workspace-agent-api'
import {
  TASK_TERMINAL_STATES,
  type AssistantTaskAnswer,
  type AssistantTaskControl,
  type AssistantTaskRequest,
  type AssistantTaskView,
  type TaskTargetVersion,
} from '../../shared/assistant-task-api'
import { AssistantTaskStore, taskHash, type TaskLease, type TaskStepStart } from './assistant-store'
import {
  observeTaskDocument,
  observedEffect,
  verifiedWordSave,
  nextRoundVersion,
  sameTaskTarget,
  taskFileChanges,
  mergeTaskFileChange,
  type DocumentObservation,
} from './document-observation'
import { portableResult } from '../workspace-agent/portable-results'
import { taskFailureReason } from './failure-reason'

export interface TaskExecutorBinding {
  /** Resolves targets through the original device/editor authority, never a renderer path override. */
  observe(call?: AgentToolCall): Promise<DocumentObservation[]>
  inspectTarget(target: TaskTargetVersion): Promise<DocumentObservation>
  classify(call: AgentToolCall): TaskStepStart['recovery']
  alive(): boolean
  stop(): void
  approveEach(): boolean
  review?(result: TaskExecutionEnd): Promise<AgentRunReview | null>
}
interface Active {
  lease: TaskLease
  binding: TaskExecutorBinding
  request: AssistantTaskRequest
  startedAt: string
  execution: AgentTaskExecution
  lastHeartbeat: number
}
const sleep = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms))
function stableInput(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(stableInput)
  if (value && typeof value === 'object')
    return Object.fromEntries(
      Object.entries(value)
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([key, item]) => [key, stableInput(item)]),
    )
  return value
}
const operation = (call: AgentToolCall) => taskHash([call.name, stableInput(call.input)])
const targets = (items: DocumentObservation[]) =>
  items.map((item) => item.version).sort((a, b) => a.path.localeCompare(b.path))

/** The local scheduler owns durable attempts; a panel is only an executor/client. */
export class AssistantTaskService {
  readonly bootId = randomUUID()
  private readonly active = new Map<string, Active>()
  private readonly waitingFiles = new Set<string>()
  private readonly workspacePromises = new Set<string>()
  private readonly workspaceOwners = new Set<string>()
  private timer?: ReturnType<typeof setInterval>
  private pumping = false
  private closed = false
  constructor(
    readonly store: AssistantTaskStore,
    private readonly options: {
      isActiveSession(id: string): boolean
      providerRevision(entry: AssistantTaskRequest['entry']): string
      workspace(
        request: AssistantTaskRequest,
        task: AssistantTaskView,
      ): Promise<{
        binding: TaskExecutorBinding
        run(execution: AgentTaskExecution): Promise<void>
      }>
      saveReview(sessionId: string, review: AgentRunReview): Promise<void>
      memoryContext?(request: AssistantTaskRequest): Promise<string>
      inspectReviewed?(
        request: AssistantTaskRequest,
        target: TaskTargetVersion,
      ): Promise<DocumentObservation>
    },
  ) {
    store.startBoot(this.bootId)
  }

  start(): void {
    if (this.timer || this.closed) return
    this.timer = setInterval(() => {
      this.tick()
      void this.pump().catch(() => {})
    }, 350)
    this.timer.unref?.()
    void this.pump().catch(() => {})
  }
  close(): void {
    this.closed = true
    if (this.timer) clearInterval(this.timer)
    for (const entry of this.active.values()) entry.binding.stop()
  }
  private tick(): void {
    const stamp = Date.now()
    for (const [id, entry] of this.active) {
      if (!entry.binding.alive()) {
        try {
          this.store.finish(entry.lease, { state: 'failed', reason: 'executor_lost' })
        } catch {
          /* stale attempts never regain authority */
        }
        this.active.delete(id)
        continue
      }
      if (stamp - entry.lastHeartbeat < 1_000) continue
      entry.lastHeartbeat = stamp
      try {
        const status = this.store.heartbeat(entry.lease)
        if (status.exhausted || status.pendingControl) entry.binding.stop()
      } catch {
        entry.binding.stop()
      }
    }
  }
  submit(request: AssistantTaskRequest): { task: AssistantTaskView; duplicate: boolean } {
    if (this.closed || !this.options.isActiveSession(request.sessionId))
      throw new Error('task_session_inactive')
    const receipt = this.store.accept({
      ...request,
      providerRevision: request.providerRevision ?? this.options.providerRevision(request.entry),
    })
    this.start()
    void this.pump().catch(() => {})
    return receipt
  }
  async wait(id: string, signal?: AbortSignal): Promise<AssistantTaskView> {
    while (!this.closed) {
      signal?.throwIfAborted()
      const value = this.store.get(id)
      if (
        TASK_TERMINAL_STATES.has(value.state) ||
        value.state === 'paused' ||
        value.state === 'interrupted' ||
        (value.state === 'waiting_user' && value.confirmation?.kind === 'clarification')
      )
        return value
      await sleep(75)
    }
    return this.store.get(id)
  }
  async control(input: AssistantTaskControl): Promise<AssistantTaskView> {
    const task = this.store.get(input.taskId)
    if (['resume', 'retry'].includes(input.action) && !this.options.isActiveSession(task.sessionId))
      throw new Error('task_session_inactive')
    let checked: TaskTargetVersion[] | undefined
    if (
      input.action === 'reconcile' &&
      input.reviewed &&
      this.options.inspectReviewed &&
      !this.active.has(task.id)
    ) {
      const request = this.store.request(task.id),
        before = this.store.reconciliationTargets(task.id)
      checked = []
      for (const item of before)
        checked.push((await this.options.inspectReviewed(request, item)).version)
    }
    const result = this.store.control(input, checked)
    if (result.pendingControl) this.active.get(result.id)?.binding.stop()
    void this.pump().catch(() => {})
    return result
  }
  async stopSession(sessionId: string): Promise<void> {
    for (const task of this.store.forSession(sessionId)) {
      if (TASK_TERMINAL_STATES.has(task.state)) continue
      const controlled = await this.control({
        taskId: task.id,
        requestId: randomUUID(),
        expectedRevision: task.revision,
        action: 'cancel',
      })
      const outcome = TASK_TERMINAL_STATES.has(controlled.state)
        ? controlled
        : await this.wait(task.id, AbortSignal.timeout(10_000))
      if (!TASK_TERMINAL_STATES.has(outcome.state)) throw new Error('task_effect_requires_review')
    }
  }
  answer(
    input: AssistantTaskAnswer,
    dispatch = true,
  ): { task: AssistantTaskView; duplicate: boolean } {
    if (!this.options.isActiveSession(this.store.get(input.taskId).sessionId))
      throw new Error('task_session_inactive')
    const result = this.store.answer(input)
    if (dispatch) void this.pump().catch(() => {})
    return result
  }
  evidence(id: string) {
    const task = this.store.get(id),
      entries = this.store.evidence(id) as Array<{ change?: AgentRunFileChange }>
    const changes = entries.flatMap((item) => (item.change ? [item.change] : []))
    if (!changes.length) return []
    return [
      portableResult({
        id,
        startedAt: new Date(task.createdAt).toISOString(),
        completedAt: new Date(task.updatedAt).toISOString(),
        state: 'completed',
        outcome: task.state === 'completed' ? 'completed' : 'partial',
        changes,
        undoable: false,
      }),
    ]
  }
  private async checkContext(
    request: AssistantTaskRequest,
    id: string,
    binding: TaskExecutorBinding,
  ): Promise<void> {
    if (!this.options.isActiveSession(request.sessionId) || !binding.alive())
      throw new Error('task_session_inactive')
    if (
      request.providerRevision &&
      request.providerRevision !== this.options.providerRevision(request.entry)
    )
      throw new Error('task_provider_changed')
    const known = this.store.plan(id).lastTargets
    for (const expected of known?.length ? known : request.document ? [request.document] : []) {
      const observed = await binding.inspectTarget(expected)
      if (!sameTaskTarget(expected, observed.version))
        throw new Error(
          expected.dirty && expected.editorIdentity !== observed.version.editorIdentity
            ? 'task_unsaved_context_unavailable'
            : 'task_context_changed',
        )
    }
  }
  private async pump(): Promise<void> {
    if (this.pumping || this.closed) return
    this.pumping = true
    try {
      for (const task of this.store.queued()) {
        if (this.closed) break
        if (!this.options.isActiveSession(task.sessionId)) {
          this.store.wait(task.id, 'waiting_resource', 'session_inactive')
          continue
        }
        if (task.entry === 'file') {
          if (!this.waitingFiles.has(task.id))
            this.store.wait(task.id, 'waiting_device', 'open_editor_to_continue')
          continue
        }
        if (this.workspacePromises.has(task.id) || this.workspaceOwners.has(task.sessionId))
          continue
        const lease = this.store.claim(task.id, this.bootId)
        if (!lease) continue
        this.workspacePromises.add(task.id)
        this.workspaceOwners.add(task.sessionId)
        void this.runWorkspace(task, lease).finally(() => {
          this.workspacePromises.delete(task.id)
          this.workspaceOwners.delete(task.sessionId)
          void this.pump().catch(() => {})
        })
      }
    } finally {
      this.pumping = false
    }
  }
  private async runWorkspace(task: AssistantTaskView, lease: TaskLease): Promise<void> {
    try {
      const request = this.store.request(task.id)
      const { binding, run } = await this.options.workspace(request, task)
      const execution = this.makeExecution(lease, request, binding)
      await run(execution)
      if (this.active.has(task.id))
        await execution.finish(
          { state: 'failed', error: 'task_executor_ended' },
          this.emptySnapshot(),
        )
    } catch (error) {
      this.active.delete(task.id)
      try {
        this.store.finish(lease, { state: 'failed', reason: this.reason(error) })
      } catch {
        /* terminal or stale */
      }
    }
  }
  private emptySnapshot(): AgentSessionSnapshot {
    return {
      version: 1,
      messages: [],
      phase: 'error',
      updatedAt: new Date().toISOString(),
      budget: { ...sessionBudget(true), compaction: false },
    }
  }
  private reason(error: unknown): string {
    return taskFailureReason(error)
  }

  async beginFile(
    request: AssistantTaskRequest,
    binding: TaskExecutorBinding,
    input: TaskStartInput,
    resumeId?: string,
    signal?: AbortSignal,
  ): Promise<{
    taskId: string
    execution: AgentTaskExecution
    checkpoint?: AgentSessionSnapshot
    remainingMs: number
    memoryContext?: string
  }> {
    let task: AssistantTaskView
    if (resumeId) {
      task = this.store.get(resumeId)
      if (
        task.sessionId !== request.sessionId ||
        task.entry !== 'file' ||
        !['queued', 'waiting_device', 'waiting_resource'].includes(task.state) ||
        !this.store.plan(task.id).resumeRequested
      )
        throw new Error('task_cannot_resume')
      request = this.store.request(task.id)
    } else {
      const accepted = this.store.accept(
        { ...request, providerRevision: this.options.providerRevision('file') },
        { maxActiveMs: input.budget.runTimeoutMs, maxModelTurns: input.budget.maxTurns },
      )
      task = accepted.task
      if (
        accepted.duplicate &&
        !['queued', 'waiting_device', 'waiting_resource'].includes(task.state)
      )
        throw new Error('task_already_accepted')
    }
    this.start()
    const id = task.id
    if (this.waitingFiles.has(id)) throw new Error('task_already_accepted')
    this.waitingFiles.add(id)
    try {
      while (!this.closed && binding.alive() && !signal?.aborted) {
        const lease = this.store.claim(id, this.bootId)
        if (lease) {
          const execution = this.makeExecution(lease, request, binding)
          try {
            const start = await execution.start(input, signal ?? new AbortController().signal)
            return {
              taskId: id,
              execution,
              checkpoint: start?.checkpoint,
              remainingMs: start?.remainingMs ?? 1,
              memoryContext: start?.memoryContext,
            }
          } catch (error) {
            this.active.delete(id)
            this.store.finish(lease, { state: 'failed', reason: this.reason(error) })
            throw error
          }
        }
        const current = this.store.get(id)
        if (
          TASK_TERMINAL_STATES.has(current.state) ||
          current.state === 'paused' ||
          current.state === 'interrupted'
        )
          throw new Error('task_stopping')
        await sleep(90)
      }
      const current = this.store.get(id)
      if (signal?.aborted && !TASK_TERMINAL_STATES.has(current.state))
        this.store.control({
          taskId: id,
          requestId: randomUUID(),
          expectedRevision: current.revision,
          action: 'cancel',
        })
      else this.store.wait(id, 'waiting_device', 'open_editor_to_continue')
      throw new Error('task_editor_unavailable')
    } finally {
      this.waitingFiles.delete(id)
    }
  }

  private makeExecution(
    lease: TaskLease,
    request: AssistantTaskRequest,
    binding: TaskExecutorBinding,
  ): AgentTaskExecution {
    const id = lease.taskId,
      before = new Map<
        string,
        {
          values: DocumentObservation[]
          call: AgentToolCall
          recovery: TaskStepStart['recovery']
          startedAt: number
        }
      >()
    const confirmations = new Map<
      string,
      { id: string; digest: string; head?: TaskTargetVersion }
    >()
    let roundHead: TaskTargetVersion | undefined
    const authorizedHeads = new Map<string, TaskTargetVersion>()
    const fileChanges = new Map<string, AgentRunFileChange>()
    const startedAt = new Date().toISOString()
    const checkSaved = async (): Promise<boolean | undefined> => {
      if (request.requiredResult !== 'saved_artifact') return undefined
      const receipts = this.store.saveTargets(id)
      if (!receipts.length) return false
      for (const item of receipts) {
        if (!item.saved) return false
        const current = await binding.inspectTarget(item.saved).catch(() => undefined)
        if (!current || current.version.dirty || !sameTaskTarget(item.saved, current.version))
          return false
      }
      return true
    }
    let finishing: Promise<void> | undefined
    const execution: AgentTaskExecution = {
      taskId: id,
      supportsClarification: true,
      checkSaved,
      pendingWordParts: async () =>
        Object.keys(this.store.plan(id).pendingWordParts ?? {}).map((key) =>
          (JSON.parse(key) as string[]).filter(Boolean).join(' · '),
        ),
      start: async (input, signal) => {
        signal.throwIfAborted()
        await this.checkContext(request, id, binding)
        const memoryContext =
          request.entry === 'file' ? await this.options.memoryContext?.(request) : undefined
        signal.throwIfAborted()
        return {
          id,
          ...this.store.begin(lease, input.snapshot),
          ...(memoryContext ? { memoryContext } : {}),
        }
      },
      reserveModel: async (snapshot, signal) => {
        signal.throwIfAborted()
        if (!binding.alive()) throw new Error('task_editor_unavailable')
        return this.store.reserveModel(lease, snapshot)
      },
      reserveAuxiliaryModel: async (signal) => {
        signal?.throwIfAborted()
        return this.store.reserveModel(lease)
      },
      usage: async (modelCallId, usage) => this.store.usage(lease, modelCallId, usage),
      beforeApproval: async (call, required, parentCall) => {
        if (binding.classify(call) === 'repeat_read' || (!required && !binding.approveEach()))
          return
        const digest = operation(call),
          observed = targets(await binding.observe(call))
        const head =
          request.entry === 'file' &&
          observed.length === 1 &&
          observed[0]?.type === 'docx' &&
          isRoundWordEdit(call) &&
          observed[0].complete &&
          observed[0].exists &&
          observed[0].diskDigest
            ? observed[0]
            : undefined
        if (head && roundHead && sameTaskTarget(roundHead, head)) {
          authorizedHeads.set(call.id, head)
          return { covered: true }
        }
        roundHead = undefined
        const confirmation = this.store.requestDecision(
          lease,
          digest,
          observed,
          call.name,
          'approval',
          parentCall ? operation(parentCall) : digest,
        )
        confirmations.set(call.id, { id: confirmation, digest, ...(head ? { head } : {}) })
        if (head && call.name !== 'save_document')
          return { round: { path: head.path, baseVersion: head.modelDigest ?? head.diskDigest! } }
      },
      afterApproval: async (call, allowed, round) => {
        const decision = confirmations.get(call.id)
        if (!decision) return
        confirmations.delete(call.id)
        if (allowed) {
          while (!this.store.acquireAfterDecision(lease)) {
            if (this.closed || !binding.alive()) throw new Error('task_editor_unavailable')
            await sleep(90)
          }
        } else if (this.store.get(id).state === 'waiting_user') {
          while (!this.store.acquireAfterDecision(lease)) {
            if (this.closed || !binding.alive()) throw new Error('task_editor_unavailable')
            await sleep(90)
          }
        }
        const current = allowed ? targets(await binding.observe(call)) : []
        const accepted = this.store.answerDecision(
          lease,
          decision.id,
          decision.digest,
          allowed,
          current,
        )
        if (allowed && !accepted) throw new Error('task_confirmation_stale')
        if (accepted && decision.head) authorizedHeads.set(call.id, decision.head)
        if (accepted && round && decision.head && call.name !== 'save_document')
          roundHead = decision.head
      },
      prepareTool: async (call, signal) => {
        signal.throwIfAborted()
        if (!binding.alive() || !this.options.isActiveSession(request.sessionId))
          throw new Error('task_session_inactive')
        const recovery = binding.classify(call)
        const values = recovery === 'repeat_read' ? [] : await binding.observe(call)
        // Revalidate only previously observed targets involved in this step.
        const known =
          this.store.plan(id).lastTargets ?? (request.document ? [request.document] : [])
        const approved = authorizedHeads.get(call.id)
        authorizedHeads.delete(call.id)
        for (const item of values) {
          const expected =
            approved?.path === item.version.path
              ? approved
              : known.find((value) => value.path === item.version.path)
          if (expected && !sameTaskTarget(expected, item.version))
            throw new Error('task_context_changed')
        }
        signal.throwIfAborted()
        const permit = this.store.prepareStep(lease, {
          digest: operation(call),
          label: call.name,
          recovery,
          targets: targets(values),
        })
        if (!permit.reused)
          before.set(permit.stepId, { values, call, recovery, startedAt: Date.now() })
        return permit
      },
      finishTool: async (permit: TaskToolPermit, result: ToolExecution, snapshot) => {
        if (permit.reused) {
          this.store.finishStep(
            lease,
            permit.stepId,
            { result, effect: 'not_applied', targets: [] },
            snapshot,
          )
          return
        }
        const started = before.get(permit.stepId)
        if (!started) throw new Error('task_step_unavailable')
        before.delete(permit.stepId)
        let after: DocumentObservation[] = [],
          observed = true
        if (started.recovery !== 'repeat_read') {
          try {
            after = await binding.observe(started.call)
          } catch {
            observed = false
          }
        }
        if (
          started.call.name === 'create_document' &&
          result.savedFile &&
          isAbsolute(result.savedFile.path)
        ) {
          const { path, sha256 } = result.savedFile
          const witness = saveWitness(path)
          const type = extname(path).slice(1).toLowerCase()
          if (
            witness?.complete &&
            witness.at >= started.startedAt &&
            witness.sha256 === sha256 &&
            (type === 'docx' || type === 'pdf' || type === 'html' || type === 'md')
          ) {
            const created = await observeTaskDocument({ path, type })
            if (created.version.complete && created.version.diskDigest === sha256) {
              created.version.generatedFile = true
              started.values = []
              after = [created]
              observed = true
            }
          }
        }
        if (roundHead && started.recovery !== 'repeat_read') {
          const previous = started.values[0]?.version,
            next = after[0]?.version
          roundHead =
            (!result.isError || result.effect === 'not_applied') &&
            previous &&
            isRoundWordEdit(started.call) &&
            sameTaskTarget(roundHead, previous) &&
            (result.effect !== 'not_applied' || (next && sameTaskTarget(previous, next))) &&
            started.call.name !== 'write_document'
              ? nextRoundVersion(previous, next)
              : undefined
        }
        const changed = observed && observedEffect(started.values, after)
        const complete =
          observed && after.length > 0 && after.every((item) => item.version.complete)
        const effect =
          result.effect ??
          result.externalEffect ??
          (started.recovery === 'repeat_read' || result.denied
            ? 'not_applied'
            : changed && complete
              ? 'applied'
              : started.recovery === 'verify_effect' && complete && !changed && !result.mutated
                ? 'not_applied'
                : started.recovery === 'user_reconcile' && !result.isError
                  ? result.mutated
                    ? 'applied'
                    : 'not_applied'
                  : 'unknown')
        const changes =
          changed && effect === 'applied'
            ? await taskFileChanges(started.values, after).catch(() => [])
            : []
        for (const change of changes) {
          change.tools = [started.call.name]
          mergeTaskFileChange(fileChanges, change)
        }
        this.store.finishStep(
          lease,
          permit.stepId,
          {
            result,
            effect,
            targets: targets(after),
            ...(after.some((item) => item.version.type === 'docx')
              ? {
                  wordSaveVerified: after.every(
                    (item) =>
                      item.version.type !== 'docx' ||
                      verifiedWordSave(
                        started.values.find(
                          (previous) =>
                            previous.version.path === item.version.path ||
                            (!!previous.version.editorIdentity &&
                              previous.version.editorIdentity === item.version.editorIdentity),
                        ),
                        item,
                      ),
                  ),
                }
              : {}),
            evidence: changes.flatMap((change) =>
              change.evidence ? [{ id: change.evidence.id, change }] : [],
            ),
          },
          snapshot,
        )
      },
      finish: (result, snapshot) =>
        (finishing ??= (async () => {
          try {
            let review = await binding.review?.(result)
            if (request.entry === 'file' || !review || lease.epoch > 1) {
              const recorded = this.store.evidence(id) as Array<{ change?: AgentRunFileChange }>
              const accumulated = new Map<string, AgentRunFileChange>()
              for (const item of recorded)
                if (item.change) mergeTaskFileChange(accumulated, item.change)
              if (!accumulated.size)
                for (const item of fileChanges.values()) mergeTaskFileChange(accumulated, item)
              const changes = [...accumulated.values()]
              if (changes.length)
                review = {
                  id,
                  state: 'completed',
                  startedAt,
                  completedAt: new Date().toISOString(),
                  changes,
                  undoable: false,
                  outcome:
                    result.state === 'cancelled'
                      ? 'cancelled'
                      : result.state === 'failed' || result.state === 'waiting_user'
                        ? 'partial'
                        : changes.some((change) => change.evidence?.verification !== 'passed')
                          ? 'partial'
                          : 'completed',
                  evidenceStored: true,
                }
            }
            const changes = review?.changes ?? [...fileChanges.values()]
            const saved =
              result.state === 'completed' && request.requiredResult === 'saved_artifact'
                ? await checkSaved()
                : undefined
            const edited = changes.some(
              (change) =>
                change.evidence?.effect === 'changed' && change.evidence.verification === 'passed',
            )
            const unfulfilled =
              result.state === 'completed' ? this.store.unfulfilledOperation(id) : undefined
            const resultMissing =
              result.state === 'completed' &&
              (!!unfulfilled ||
                Object.keys(this.store.plan(id).pendingWordParts ?? {}).length > 0 ||
                (request.requiredResult === 'saved_artifact' && !saved) ||
                (request.requiredResult === 'editor_change' && !edited))
            if (review) {
              review = {
                ...review,
                taskId: id,
                ...(resultMissing
                  ? { outcome: edited ? ('partial' as const) : ('failed' as const) }
                  : {}),
              }
              await this.options.saveReview(request.sessionId, review)
            }
            this.store.finish(
              lease,
              {
                state: resultMissing ? 'failed' : result.state,
                question: result.question,
                reason: resultMissing
                  ? (unfulfilled ?? 'result_not_verified')
                  : result.state === 'failed'
                    ? this.reason(new Error(result.error ?? ''))
                    : undefined,
              },
              snapshot,
            )
          } catch (error) {
            try {
              this.store.finish(lease, { state: 'failed', reason: this.reason(error) })
            } catch {
              /* boot recovery preserves uncertain effects */
            }
            throw error
          } finally {
            this.active.delete(id)
            void this.pump().catch(() => {})
          }
        })()),
    }
    this.active.set(id, {
      lease,
      request,
      binding,
      startedAt,
      execution,
      lastHeartbeat: Date.now(),
    })
    return execution
  }
}
