import { createHash, randomUUID } from 'node:crypto'
import type { DatabaseSync } from 'node:sqlite'
import type { AgentSessionSnapshot, AgentTokenUsage, ToolExecution } from '@genoffice/agent-core'
import { sessionBudget, taskMessagePrefix } from '@genoffice/agent-core'
import {
  TASK_TERMINAL_STATES,
  type AssistantTaskAnswer,
  type AssistantTaskBudget,
  type AssistantTaskControl,
  type AssistantTaskFeed,
  type AssistantTaskRequest,
  type AssistantTaskState,
  type AssistantTaskView,
  type TaskTargetVersion,
} from '../../shared/assistant-task-api'

import { updateWordChecks } from '@genoffice/agent-core'
import type { DeliveryState } from './delivery'
type Row = Record<string, unknown>
export const TASK_EXECUTOR = 'local-office'
const RESOURCE = 'local-office:mutation'
const targetKey = (path: string) =>
  process.platform === 'win32' ? path.replace(/\\/g, '/').toLowerCase() : path
function resourceDomain(tool: string): string {
  if (tool.startsWith('computer')) return 'computer'
  if (tool === 'run_cli' || tool === 'workspace.run_command') return 'command'
  return tool.split('.')[0]!
}
const LEASE_MS = 30_000
const MAX_CHECKPOINT_BYTES = 4 * 1024 * 1024
const MAX_REQUEST_BYTES = 36 * 1024 * 1024
const MAX_MEDIA_BYTES = 18 * 1024 * 1024
const MEDIA_PREFIX = 'task-media:'
const QUEUE_STATES = new Set(['queued', 'waiting_device', 'waiting_resource', 'interrupted'])
const ACTIVE_STATES = new Set(['preparing', 'running', 'verifying', 'pausing', 'cancelling'])
export const taskHash = (value: unknown) =>
  createHash('sha256').update(JSON.stringify(value)).digest('hex')
export const taskId = (value: unknown): string => {
  if (typeof value !== 'string' || !/^[a-z0-9_-]{1,128}$/i.test(value))
    throw new Error('invalid_task_id')
  return value
}
function json(value: unknown, max = MAX_CHECKPOINT_BYTES): string {
  const text = JSON.stringify(value)
  if (typeof text !== 'string' || Buffer.byteLength(text) > max)
    throw new Error('task_payload_too_large')
  return text
}
const parse = <T>(value: unknown): T => JSON.parse(String(value)) as T
const bounded = (value: unknown, max: number) =>
  typeof value === 'string' ? value.slice(0, max) : ''
export interface TaskLease {
  taskId: string
  attemptId: string
  bootId: string
  epoch: number
}
export interface TaskStepStart {
  digest: string
  label: string
  recovery: 'repeat_read' | 'verify_effect' | 'user_reconcile'
  targets: TaskTargetVersion[]
}
export interface TaskStepReceipt {
  delivery?: DeliveryState
  wordSaveVerified?: boolean
  result: ToolExecution
  targets: TaskTargetVersion[]
  effect: 'applied' | 'not_applied' | 'unknown'
  evidence?: Array<{ id: string; [key: string]: unknown }>
}
interface Plan {
  delivery?: DeliveryState
  accountedAt?: number
  resumeRequested?: boolean
  retryRequestId?: string
  lastTargets?: TaskTargetVersion[]
  pendingWordParts?: Record<string, string>
  modelCalls?: Record<string, number | null>
  reply?: { confirmationId: string; text: string }
}

/** A single SQLite writer shares the maintenance ledger's existing schema. */
export class AssistantTaskStore {
  constructor(
    private readonly db: DatabaseSync,
    private readonly now: () => number = Date.now,
  ) {}

  private tx<T>(run: () => T): T {
    this.db.exec('BEGIN IMMEDIATE')
    try {
      const value = run()
      this.db.exec('COMMIT')
      return value
    } catch (error) {
      this.db.exec('ROLLBACK')
      throw error
    }
  }
  private row(id: string): Row {
    const row = this.db
      .prepare("SELECT * FROM tasks WHERE kind='assistant' AND id=?")
      .get(taskId(id))
    if (!row) throw new Error('unknown_assistant_task')
    return row
  }
  /** Find a durable acceptance before re-evaluating current provider/document configuration. */
  receiptForRequest(
    requestId: string,
  ): { task: AssistantTaskView; request: AssistantTaskRequest } | null {
    const row = this.db
      .prepare(
        "SELECT id FROM tasks WHERE authority_id='local' AND principal_id='assistant' AND request_id=? AND kind='assistant'",
      )
      .get(requestId)
    if (row) return { task: this.get(String(row.id)), request: this.request(String(row.id)) }
    const reply = this.db
      .prepare(
        "SELECT task_id,response_json FROM task_confirmations WHERE response_request_id=? AND kind='clarification' AND response_json IS NOT NULL",
      )
      .get(requestId)
    if (!reply) return null
    const request = this.request(String(reply.task_id))
    return {
      task: this.get(String(reply.task_id)),
      request: {
        ...request,
        input: { ...request.input, text: parse<{ answer: string }>(reply.response_json).answer },
      },
    }
  }
  private event(id: string, type: string, label: string, refs: string[] = []): number {
    const row = this.row(id),
      sequence = Number(row.next_event_sequence)
    this.db
      .prepare(
        'INSERT INTO task_events(id,task_id,sequence,revision,type,summary,refs_json,created_at) VALUES(?,?,?,?,?,?,?,?)',
      )
      .run(
        randomUUID(),
        id,
        sequence,
        Number(row.revision),
        type,
        bounded(label, 240),
        json(refs),
        this.now(),
      )
    this.db.prepare('UPDATE tasks SET next_event_sequence=next_event_sequence+1 WHERE id=?').run(id)
    return sequence
  }
  private transition(id: string, state: AssistantTaskState, reason?: string): number {
    this.db
      .prepare('UPDATE tasks SET state=?,reason_code=?,revision=revision+1,updated_at=? WHERE id=?')
      .run(state, reason ?? null, this.now(), id)
    return this.event(id, 'task.' + state, reason ?? state)
  }
  private notice(
    id: string,
    sequence: number,
    kind: 'attention_required' | 'result_ready' | 'task_blocked',
  ): void {
    this.db
      .prepare(
        "INSERT OR IGNORE INTO task_outbox(id,task_id,event_sequence,destination_id,route,kind,dedupe_key,payload_json,state) VALUES(?,?,?,'shell','local',?,?,?,'pending')",
      )
      .run(
        randomUUID(),
        id,
        sequence,
        kind,
        `${id}:${sequence}:${kind}`,
        json({ taskId: id, kind }),
      )
  }
  private release(id: string, keepOwner = false): void {
    // Uncertain steps retain their target fences after releasing the active executor.
    this.db.prepare('DELETE FROM execution_resources WHERE owner_task_id=?').run(id)
    this.db.prepare('DELETE FROM task_leases WHERE task_id=?').run(id)
    if (!keepOwner) this.db.prepare('DELETE FROM session_owners WHERE task_id=?').run(id)
  }
  private invalidateConfirmations(id: string): void {
    this.db
      .prepare(
        "UPDATE task_confirmations SET state='invalidated',resolved_at=? WHERE task_id=? AND state='pending'",
      )
      .run(this.now(), id)
    this.db
      .prepare(
        "UPDATE task_outbox SET state='cancelled' WHERE task_id=? AND kind='attention_required' AND state IN ('pending','delivered','inflight')",
      )
      .run(id)
  }
  private uncertain(id: string): boolean {
    return !!this.db
      .prepare(
        "SELECT 1 FROM task_steps WHERE task_id=? AND state IN ('executing','awaiting_reconciliation') AND recovery!='repeat_read' LIMIT 1",
      )
      .get(id)
  }

  startBoot(bootId: string): void {
    this.tx(() => {
      const stamp = this.now()
      this.db
        .prepare(
          "INSERT INTO executors(id,device_id,boot_id,authority_revision,capabilities_json,availability,last_seen_at) VALUES(?,?,?,1,'{}','online',?) ON CONFLICT(id) DO UPDATE SET boot_id=excluded.boot_id,availability='online',last_seen_at=excluded.last_seen_at",
        )
        .run(TASK_EXECUTOR, TASK_EXECUTOR, taskId(bootId), stamp)
      // Older ledgers used a machine-wide quarantine. Their unfinished steps still identify the targets.
      this.db
        .prepare(
          "DELETE FROM execution_resources WHERE state='quarantined' AND owner_task_id IN (SELECT id FROM tasks WHERE kind='assistant')",
        )
        .run()
      const old = this.db
        .prepare('SELECT task_id,attempt_id FROM task_leases WHERE boot_id<>?')
        .all(bootId)
      for (const value of old) {
        const id = String(value.task_id),
          row = this.row(id),
          plan = parse<Plan>(row.plan_json)
        this.db
          .prepare("UPDATE task_attempts SET state='lost',ended_at=? WHERE id=?")
          .run(stamp, String(value.attempt_id))
        this.db
          .prepare(
            "UPDATE task_steps SET state='superseded',effect_state='not_applied',revision=revision+1 WHERE task_id=? AND (state='prepared' OR (state='executing' AND recovery='repeat_read'))",
          )
          .run(id)
        this.db
          .prepare(
            "UPDATE task_steps SET state='awaiting_reconciliation',effect_state='unknown',revision=revision+1 WHERE task_id=? AND state='executing'",
          )
          .run(id)
        const uncertain = this.uncertain(id),
          waited = row.state === 'waiting_user'
        this.invalidateConfirmations(id)
        this.release(id, uncertain || (!row.pending_control && !waited))
        const state = uncertain
          ? 'interrupted'
          : row.pending_control === 'cancel'
            ? 'cancelled'
            : row.pending_control === 'pause'
              ? 'paused'
              : 'interrupted'
        if (!uncertain) this.db.prepare('UPDATE tasks SET pending_control=NULL WHERE id=?').run(id)
        this.db.prepare('UPDATE tasks SET plan_json=? WHERE id=?').run(
          json({
            ...plan,
            accountedAt: undefined,
            resumeRequested: !uncertain && !waited && !row.pending_control,
          }),
          id,
        )
        const sequence = this.transition(
          id,
          state,
          uncertain
            ? 'outcome_unknown'
            : waited
              ? 'fresh_confirmation_required'
              : 'restart_checkpoint',
        )
        if (uncertain || waited) this.notice(id, sequence, 'task_blocked')
      }
      this.db
        .prepare(
          "UPDATE task_outbox SET state='pending',delivery_lease_until=NULL WHERE route='local' AND state='inflight'",
        )
        .run()
    })
  }
  accept(
    request: AssistantTaskRequest,
    limits?: Partial<Pick<AssistantTaskBudget, 'maxActiveMs' | 'maxModelTurns'>>,
  ): { task: AssistantTaskView; duplicate: boolean } {
    taskId(request.requestId)
    taskId(request.sessionId)
    if (!request.input.text.trim() || request.input.text.length > 128_000)
      throw new Error('invalid_task_goal')
    const digest = taskHash([
        request.requestId,
        request.sessionId,
        request.entry,
        request.input,
        request.requiredResult,
      ]),
      body = json(request, MAX_REQUEST_BYTES)
    return this.tx(() => {
      if (
        this.db
          .prepare('SELECT 1 FROM session_tombstones WHERE session_id=?')
          .get(request.sessionId) ||
        this.db
          .prepare(
            "SELECT 1 FROM task_tombstones WHERE authority_id='local' AND principal_id='assistant' AND request_id=?",
          )
          .get(request.requestId)
      )
        throw new Error('task_was_purged')
      const old = this.db
        .prepare(
          "SELECT id,request_hash FROM tasks WHERE authority_id='local' AND principal_id='assistant' AND request_id=?",
        )
        .get(request.requestId)
      if (old) {
        if (old.request_hash !== digest) throw new Error('request_conflict')
        return { task: this.get(String(old.id)), duplicate: true }
      }
      const id = randomUUID(),
        stamp = this.now()
      const budget: AssistantTaskBudget = {
        maxAttempts: 3,
        maxActiveMs: Math.max(1, Math.min(600_000, limits?.maxActiveMs ?? 600_000)),
        maxModelTurns: Math.max(
          1,
          Math.min(
            100,
            limits?.maxModelTurns ?? sessionBudget(request.entry === 'workspace').maxTurns,
          ),
        ),
        usedAttempts: 0,
        usedActiveMs: 0,
        usedModelTurns: 0,
        usedReportedTokens: 0,
        unreportedModelCalls: 0,
      }
      this.db
        .prepare(
          "INSERT INTO tasks(id,authority_id,principal_id,request_id,request_hash,kind,session_id,title,goal_json,context_json,budget_json,state,created_at,updated_at) VALUES(?,'local','assistant',?,?,'assistant',?,?,?,? ,?,'queued',?,?)",
        )
        .run(
          id,
          request.requestId,
          digest,
          request.sessionId,
          request.input.text.trim().slice(0, 100),
          body,
          json(request.document ?? {}),
          json(budget),
          stamp,
          stamp,
        )
      this.event(id, 'task.accepted', 'Task accepted')
      return { task: this.get(id), duplicate: false }
    })
  }
  request(id: string): AssistantTaskRequest {
    return parse(this.row(id).goal_json)
  }
  plan(id: string): Plan {
    return parse(this.row(id).plan_json)
  }
  queued(): AssistantTaskView[] {
    return this.db
      .prepare(
        "SELECT id FROM tasks WHERE kind='assistant' AND (state IN ('queued','waiting_resource','waiting_device') OR (state='interrupted' AND json_extract(plan_json,'$.resumeRequested')=1)) ORDER BY created_at,rowid",
      )
      .all()
      .map((row) => this.get(String(row.id)))
  }
  wait(id: string, state: 'waiting_device' | 'waiting_resource', reason: string): void {
    this.tx(() => {
      const row = this.row(id)
      if (QUEUE_STATES.has(String(row.state)) && row.state !== state && !this.uncertain(id))
        this.transition(id, state, reason)
    })
  }
  block(id: string, reason: string): void {
    this.tx(() => {
      if (!QUEUE_STATES.has(String(this.row(id).state))) return
      const plan = this.plan(id)
      this.db
        .prepare('UPDATE tasks SET plan_json=? WHERE id=?')
        .run(json({ ...plan, resumeRequested: false }), id)
      this.notice(id, this.transition(id, 'interrupted', reason), 'task_blocked')
    })
  }
  claim(id: string, bootId: string): TaskLease | null {
    return this.tx(() => {
      if (
        this.db.prepare('SELECT boot_id FROM executors WHERE id=?').get(TASK_EXECUTOR)?.boot_id !==
        bootId
      )
        throw new Error('stale_task_executor')
      const row = this.row(id),
        plan = parse<Plan>(row.plan_json)
      if (
        !QUEUE_STATES.has(String(row.state)) ||
        (row.state === 'interrupted' && !plan.resumeRequested) ||
        this.uncertain(id)
      )
        return null
      if (this.db.prepare('SELECT 1 FROM task_leases WHERE task_id=?').get(id)) return null
      const owner = this.db
        .prepare('SELECT task_id FROM session_owners WHERE session_id=?')
        .get(String(row.session_id))
      if (
        (owner && owner.task_id !== id) ||
        this.db.prepare('SELECT 1 FROM execution_resources WHERE resource_key=?').get(RESOURCE)
      ) {
        if (row.state !== 'waiting_resource')
          this.transition(id, 'waiting_resource', 'resource_busy')
        return null
      }
      const budget = parse<AssistantTaskBudget>(row.budget_json)
      if (
        budget.usedAttempts >= budget.maxAttempts ||
        budget.usedModelTurns >= budget.maxModelTurns ||
        budget.usedActiveMs >= budget.maxActiveMs
      ) {
        this.notice(id, this.transition(id, 'failed', 'budget_exhausted'), 'task_blocked')
        this.release(id)
        return null
      }
      const attemptId = randomUUID(),
        epoch = Number(row.lease_epoch) + 1,
        stamp = this.now()
      this.invalidateConfirmations(id)
      this.db
        .prepare(
          "INSERT INTO task_attempts(id,task_id,executor_id,boot_id,epoch,state,started_at) VALUES(?,?,?,?,?,'active',?)",
        )
        .run(attemptId, id, TASK_EXECUTOR, bootId, epoch, stamp)
      this.db
        .prepare(
          'INSERT INTO task_leases(task_id,attempt_id,executor_id,boot_id,epoch,expires_at) VALUES(?,?,?,?,?,?)',
        )
        .run(id, attemptId, TASK_EXECUTOR, bootId, epoch, stamp + LEASE_MS)
      this.db
        .prepare('INSERT OR IGNORE INTO session_owners(session_id,task_id) VALUES(?,?)')
        .run(String(row.session_id), id)
      this.db
        .prepare(
          "INSERT INTO execution_resources(resource_key,device_id,owner_task_id,owner_attempt_id,owner_epoch,state,proof_json,updated_at) VALUES(?,?,?,?,?,'held','{}',?)",
        )
        .run(RESOURCE, TASK_EXECUTOR, id, attemptId, epoch, stamp)
      this.db
        .prepare(
          'UPDATE tasks SET executor_id=?,lease_epoch=?,budget_json=?,plan_json=?,pending_control=NULL WHERE id=?',
        )
        .run(
          TASK_EXECUTOR,
          epoch,
          json({ ...budget, usedAttempts: budget.usedAttempts + 1 }),
          json({ ...plan, accountedAt: stamp, resumeRequested: false }),
          id,
        )
      this.transition(id, 'preparing')
      return { taskId: id, attemptId, bootId, epoch }
    })
  }
  private lease(lease: TaskLease, allowExpired = false): Row {
    const row = this.row(lease.taskId)
    const current = this.db.prepare('SELECT * FROM task_leases WHERE task_id=?').get(lease.taskId)
    if (
      !current ||
      current.attempt_id !== lease.attemptId ||
      current.boot_id !== lease.bootId ||
      current.epoch !== lease.epoch ||
      row.lease_epoch !== lease.epoch ||
      (!allowExpired && Number(current.expires_at) <= this.now())
    )
      throw new Error('stale_task_lease')
    return row
  }
  private account(row: Row): { budget: AssistantTaskBudget; plan: Plan } {
    const budget = parse<AssistantTaskBudget>(row.budget_json),
      plan = parse<Plan>(row.plan_json),
      stamp = this.now()
    if (plan.accountedAt !== undefined && ACTIVE_STATES.has(String(row.state)))
      budget.usedActiveMs += Math.max(0, stamp - plan.accountedAt)
    plan.accountedAt = stamp
    return { budget, plan }
  }
  private accountSave(id: string, budget: AssistantTaskBudget, plan: Plan): void {
    this.db
      .prepare('UPDATE tasks SET budget_json=?,plan_json=? WHERE id=?')
      .run(json(budget), json(plan), id)
  }
  heartbeat(lease: TaskLease): { pendingControl?: 'pause' | 'cancel'; exhausted: boolean } {
    return this.tx(() => {
      const row = this.lease(lease),
        { budget, plan } = this.account(row)
      this.accountSave(lease.taskId, budget, plan)
      this.db
        .prepare('UPDATE task_leases SET expires_at=? WHERE task_id=?')
        .run(this.now() + LEASE_MS, lease.taskId)
      return {
        pendingControl: (row.pending_control ?? undefined) as 'pause' | 'cancel' | undefined,
        exhausted: budget.usedActiveMs >= budget.maxActiveMs,
      }
    })
  }
  private checkpoint(lease: TaskLease, snapshot: AgentSessionSnapshot): void {
    if (snapshot.version !== 1 || !Array.isArray(snapshot.messages))
      throw new Error('invalid_task_checkpoint')
    const id = lease.taskId,
      stamp = this.now()
    const value = {
      ...snapshot,
      messages: taskMessagePrefix(snapshot.messages).map((message) => {
        if (message.role !== 'user' || !message.images?.length) return message
        return {
          ...message,
          images: message.images.map((image) => {
            if (
              typeof image.base64 !== 'string' ||
              image.base64.length > MAX_MEDIA_BYTES ||
              image.base64.startsWith(MEDIA_PREFIX)
            )
              throw new Error('invalid_task_attachment')
            const media = { mime: image.mime, base64: image.base64 },
              digest = taskHash(media)
            const existing = this.db
              .prepare(
                "SELECT id FROM task_artifacts WHERE task_id=? AND content_digest=? AND json_extract(evidence_json,'$.role')='media' LIMIT 1",
              )
              .get(id, digest)
            const artifactId = existing ? String(existing.id) : randomUUID()
            if (!existing)
              this.db
                .prepare(
                  'INSERT INTO task_artifacts(id,task_id,kind,owned_by_store,storage_ref_json,content_digest,byte_length,evidence_json,created_at) VALUES(?,?,\'checkpoint\',1,?,?,?,\'{"role":"media"}\',?)',
                )
                .run(
                  artifactId,
                  id,
                  json(media, MAX_MEDIA_BYTES + 1024),
                  digest,
                  Buffer.byteLength(image.base64),
                  stamp,
                )
            return { mime: image.mime, base64: MEDIA_PREFIX + artifactId }
          }),
        }
      }),
    }
    const content = json(value),
      digest = taskHash(value)
    const last = this.db
      .prepare(
        'SELECT a.content_digest FROM task_checkpoints c JOIN task_artifacts a ON a.id=c.conversation_artifact_id WHERE c.task_id=? ORDER BY c.rowid DESC LIMIT 1',
      )
      .get(id)
    if (last?.content_digest === digest) return
    const artifact = randomUUID(),
      context = randomUUID(),
      request = this.request(id)
    const contextValue = {
      sessionId: request.sessionId,
      document: request.document ?? null,
      providerRevision: request.providerRevision ?? null,
    }
    const insert = this.db.prepare(
      "INSERT INTO task_artifacts(id,task_id,kind,owned_by_store,storage_ref_json,content_digest,byte_length,created_at) VALUES(?,?,'checkpoint',1,?,?,?,?)",
    )
    insert.run(
      artifact,
      id,
      json({ inline: content }, MAX_CHECKPOINT_BYTES * 2),
      digest,
      Buffer.byteLength(content),
      stamp,
    )
    insert.run(
      context,
      id,
      json({ inline: json(contextValue) }),
      taskHash(contextValue),
      Buffer.byteLength(json(contextValue)),
      stamp,
    )
    const checkpointId = randomUUID()
    const sequence = this.event(id, 'checkpoint.committed', 'Checkpoint saved', [checkpointId])
    this.db
      .prepare(
        'INSERT INTO task_checkpoints(id,task_id,attempt_id,epoch,through_sequence,conversation_artifact_id,context_artifact_id,progress_json,created_at) VALUES(?,?,?,?,?,?,?,?,?)',
      )
      .run(
        checkpointId,
        id,
        lease.attemptId,
        lease.epoch,
        sequence,
        artifact,
        context,
        json(this.get(id).budget),
        stamp,
      )
  }
  restore(id: string): AgentSessionSnapshot | undefined {
    const row = this.db
      .prepare(
        'SELECT a.storage_ref_json,a.content_digest,b.storage_ref_json AS context,b.content_digest AS context_digest FROM task_checkpoints c JOIN task_artifacts a ON a.id=c.conversation_artifact_id JOIN task_artifacts b ON b.id=c.context_artifact_id WHERE c.task_id=? ORDER BY c.rowid DESC LIMIT 1',
      )
      .get(id)
    if (!row) return
    const snapshot = parse<AgentSessionSnapshot>(
      parse<{ inline: string }>(row.storage_ref_json).inline,
    )
    const context = parse<{ sessionId: string }>(parse<{ inline: string }>(row.context).inline)
    if (
      taskHash(snapshot) !== row.content_digest ||
      taskHash(context) !== row.context_digest ||
      context.sessionId !== this.request(id).sessionId
    )
      throw new Error('task_checkpoint_corrupt')
    return {
      ...snapshot,
      messages: snapshot.messages.map((message) => {
        if (message.role !== 'user' || !message.images?.length) return message
        return {
          ...message,
          images: message.images.map((image) => {
            if (!image.base64.startsWith(MEDIA_PREFIX)) return image
            const artifactId = taskId(image.base64.slice(MEDIA_PREFIX.length))
            const item = this.db
              .prepare(
                "SELECT storage_ref_json,content_digest FROM task_artifacts WHERE task_id=? AND id=? AND json_extract(evidence_json,'$.role')='media'",
              )
              .get(id, artifactId)
            if (!item) throw new Error('task_checkpoint_corrupt')
            const media = parse<{ mime: string; base64: string }>(item.storage_ref_json)
            if (taskHash(media) !== item.content_digest || media.mime !== image.mime)
              throw new Error('task_checkpoint_corrupt')
            return media
          }),
        }
      }),
    }
  }
  begin(
    lease: TaskLease,
    snapshot: AgentSessionSnapshot,
  ): { checkpoint?: AgentSessionSnapshot; remainingMs: number } {
    return this.tx(() => {
      const row = this.lease(lease),
        { budget, plan } = this.account(row)
      if (row.state !== 'preparing' || row.pending_control) throw new Error('task_stopping')
      let checkpoint = this.restore(lease.taskId)
      if (plan.reply) {
        if (!checkpoint) throw new Error('task_checkpoint_corrupt')
        checkpoint = {
          ...checkpoint,
          messages: [...checkpoint.messages, { role: 'user', text: plan.reply.text }],
        }
        delete plan.reply
      }
      if (plan.retryRequestId) {
        if (checkpoint) {
          const goal = this.request(lease.taskId).input.text
          checkpoint = {
            ...checkpoint,
            phase: 'running',
            pendingInput: goal,
            messages: [
              ...checkpoint.messages,
              {
                role: 'user',
                text:
                  'The user explicitly requested Retry for this same task. Continue toward the original goal below. Previous denials or expired approvals are not permission grants: request fresh approval before any protected action. Recheck recorded effects and do not repeat verified or uncertain operations.\nOriginal goal: ' +
                  bounded(goal, 8000),
              },
            ],
          }
        }
        delete plan.retryRequestId
      }
      this.accountSave(lease.taskId, budget, plan)
      this.checkpoint(lease, checkpoint ?? snapshot)
      this.transition(lease.taskId, 'running')
      return { checkpoint, remainingMs: Math.max(1, budget.maxActiveMs - budget.usedActiveMs) }
    })
  }
  reserveModel(lease: TaskLease, snapshot?: AgentSessionSnapshot): string {
    return this.tx(() => {
      const row = this.lease(lease),
        { budget, plan } = this.account(row)
      if (row.state !== 'running' || row.pending_control) throw new Error('task_stopping')
      if (snapshot && this.uncertain(lease.taskId)) throw new Error('task_effect_requires_review')
      if (
        !snapshot &&
        !this.db
          .prepare(
            "SELECT 1 FROM task_steps WHERE task_id=? AND state='executing' AND plan_revision=?",
          )
          .get(lease.taskId, lease.epoch)
      )
        throw new Error('task_step_unavailable')
      if (
        budget.usedModelTurns >= budget.maxModelTurns ||
        budget.usedActiveMs >= budget.maxActiveMs
      )
        throw new Error('task_budget_exhausted')
      const id = randomUUID()
      budget.usedModelTurns++
      budget.unreportedModelCalls++
      plan.modelCalls = { ...plan.modelCalls, [id]: null }
      this.accountSave(lease.taskId, budget, plan)
      if (snapshot) this.checkpoint(lease, snapshot)
      this.event(lease.taskId, 'model.reserved', 'Model call reserved', [id])
      return id
    })
  }
  usage(lease: TaskLease, id: string, usage: AgentTokenUsage): void {
    this.tx(() => {
      const row = this.lease(lease, true),
        { budget, plan } = this.account(row)
      if (!plan.modelCalls || !(id in plan.modelCalls)) return
      const number = (value: unknown) =>
        typeof value === 'number' && Number.isFinite(value)
          ? Math.min(1_000_000_000, Math.max(0, Math.trunc(value)))
          : 0
      const total = number(usage.inputTokens) + number(usage.outputTokens),
        previous = plan.modelCalls[id]
      if (previous === null) budget.unreportedModelCalls--
      budget.usedReportedTokens += total - (previous ?? 0)
      plan.modelCalls[id] = total
      this.accountSave(lease.taskId, budget, plan)
    })
  }
  prepareStep(lease: TaskLease, input: TaskStepStart): { stepId: string; reused?: ToolExecution } {
    return this.tx(() => {
      const row = this.lease(lease),
        { budget, plan } = this.account(row)
      if (
        row.state !== 'running' ||
        row.pending_control ||
        budget.usedActiveMs >= budget.maxActiveMs
      )
        throw new Error('task_stopping')
      if (this.uncertain(lease.taskId)) throw new Error('task_effect_requires_review')
      if (input.recovery !== 'repeat_read') {
        const unresolved = this.db
          .prepare(
            "SELECT label,targets_json,recovery FROM task_steps WHERE state='awaiting_reconciliation' AND task_id<>?",
          )
          .all(lease.taskId)
        for (const row of unresolved) {
          if (row.recovery === 'repeat_read') continue
          const previous = parse<TaskTargetVersion[]>(row.targets_json)
          const overlaps = previous.some((a) =>
            input.targets.some(
              (b) =>
                targetKey(a.path) === targetKey(b.path) ||
                (!!a.fileIdentity && a.fileIdentity === b.fileIdentity) ||
                (!!a.editorIdentity && a.editorIdentity === b.editorIdentity),
            ),
          )
          if (
            overlaps ||
            ((row.recovery === 'user_reconcile' || !previous.length) &&
              resourceDomain(String(row.label)) === resourceDomain(input.label))
          ) {
            throw new Error('task_resource_requires_review')
          }
        }
      }
      if (lease.epoch > 1 && input.recovery !== 'repeat_read') {
        const reviewed = this.db
          .prepare(
            "SELECT 1 FROM task_steps WHERE task_id=? AND operation_digest=? AND state='superseded' AND effect_state='unknown' LIMIT 1",
          )
          .get(lease.taskId, input.digest)
        if (reviewed)
          throw new Error(
            'This exact operation was already reviewed. Read its target and continue from the current state without replaying it.',
          )
        const prior = this.db
          .prepare(
            "SELECT id,receipt_json FROM task_steps WHERE task_id=? AND operation_digest=? AND plan_revision<? AND state IN ('verified','failed') AND effect_state='applied' ORDER BY rowid DESC LIMIT 1",
          )
          .get(lease.taskId, input.digest, lease.epoch)
        if (prior) {
          const receipt = parse<TaskStepReceipt>(prior.receipt_json)
          return {
            stepId: String(prior.id),
            reused: {
              ...receipt.result,
              mutated: false,
              summary: 'Previous effect verified; no write repeated.',
              output:
                'Already applied in a previous attempt. ' + bounded(receipt.result.output, 12_000),
            },
          }
        }
      }
      const id = randomUUID(),
        count = Number(
          this.db.prepare('SELECT count(*) AS n FROM task_steps WHERE task_id=?').get(lease.taskId)!
            .n,
        )
      if (count >= 512) throw new Error('task_step_limit')
      this.db
        .prepare(
          "INSERT INTO task_steps(id,task_id,label,plan_revision,display_order,operation_digest,arguments_ref_json,targets_json,recovery,state,effect_state) VALUES(?,?,?,?,?,?,'{}',?,?,'executing','started')",
        )
        .run(
          id,
          lease.taskId,
          bounded(input.label, 100),
          lease.epoch,
          count,
          input.digest,
          json(input.targets),
          input.recovery,
        )
      this.accountSave(lease.taskId, budget, plan)
      this.event(lease.taskId, 'step.started', input.label, [id])
      return { stepId: id }
    })
  }
  finishStep(
    lease: TaskLease,
    stepId: string,
    receipt: TaskStepReceipt,
    snapshot: AgentSessionSnapshot,
  ): void {
    this.tx(() => {
      const row = this.lease(lease, true),
        { budget, plan } = this.account(row)
      const step = this.db
        .prepare('SELECT state,plan_revision FROM task_steps WHERE id=? AND task_id=?')
        .get(taskId(stepId), lease.taskId)
      if (!step) throw new Error('unknown_task_step')
      if (step.plan_revision === lease.epoch) {
        if (step.state !== 'executing') throw new Error('stale_task_step')
        this.db
          .prepare(
            'UPDATE task_steps SET state=?,effect_state=?,receipt_json=?,revision=revision+1 WHERE id=?',
          )
          .run(
            receipt.effect === 'unknown'
              ? 'awaiting_reconciliation'
              : receipt.result.isError
                ? 'failed'
                : 'verified',
            receipt.effect,
            json({
              ...receipt,
              evidence: undefined,
              delivery: undefined, // The baseline is stored once in the plan, atomically with this receipt.
              result: {
                output: bounded(receipt.result.output, 20_000),
                summary: bounded(receipt.result.summary, 240),
                isError: receipt.result.isError,
                mutated: receipt.result.mutated,
                effect: receipt.effect,
                denied: receipt.result.denied,
              },
            }),
            stepId,
          )
        for (const evidence of receipt.evidence ?? [])
          this.db
            .prepare(
              'INSERT INTO task_evidence(id,task_id,step_id,evidence_json,created_at) VALUES(?,?,?,?,?)',
            )
            .run(taskId(evidence.id), lease.taskId, stepId, json(evidence), this.now())
        const targets = new Map((plan.lastTargets ?? []).map((target) => [target.path, target]))
        for (const target of receipt.targets) {
          if (target.editorIdentity && target.exists)
            for (const [path, previous] of targets) {
              if (
                !previous.exists &&
                previous.editorIdentity === target.editorIdentity &&
                path !== target.path
              )
                targets.delete(path)
            }
          targets.set(target.path, target)
        }
        plan.lastTargets = [...targets.values()]
        if (receipt.delivery) plan.delivery = receipt.delivery
        if (receipt.result.wordPart) {
          const missing = new Map(Object.entries(plan.pendingWordParts ?? {}))
          updateWordChecks(missing, receipt.result.wordPart)
          plan.pendingWordParts = Object.fromEntries(missing)
        }
        this.event(
          lease.taskId,
          'step.' + (receipt.effect === 'unknown' ? 'uncertain' : 'observed'),
          String(stepId),
          [stepId],
        )
      } else if (!['verified', 'failed'].includes(String(step.state)))
        throw new Error('stale_task_step')
      this.accountSave(lease.taskId, budget, plan)
      this.checkpoint(lease, snapshot)
    })
  }
  requestDecision(
    lease: TaskLease,
    digest: string,
    targets: TaskTargetVersion[],
    label: string,
    kind: 'approval' | 'clarification' = 'approval',
    stepDigest = digest,
  ): string {
    return this.tx(() => {
      const row = this.lease(lease),
        { budget, plan } = this.account(row)
      if (row.state !== 'running' || row.pending_control) throw new Error('task_stopping')
      const step = this.db
        .prepare(
          "SELECT id FROM task_steps WHERE task_id=? AND operation_digest=? AND state='executing' AND plan_revision=? ORDER BY rowid DESC LIMIT 1",
        )
        .get(lease.taskId, stepDigest, lease.epoch)
      if (step)
        this.db
          .prepare(
            "UPDATE task_steps SET state='prepared',effect_state='not_started',revision=revision+1 WHERE id=?",
          )
          .run(String(step.id))
      const id = randomUUID(),
        revision = Number(row.revision) + 1
      this.db
        .prepare(
          "INSERT INTO task_confirmations(id,task_id,step_id,kind,state,revision,subject_digest,authority_revision,context_revisions_json,question_json,expires_at) VALUES(?,?,?,?,'pending',?,?,1,?,?,?)",
        )
        .run(
          id,
          lease.taskId,
          step ? String(step.id) : null,
          kind,
          revision,
          digest,
          json(targets),
          json({ text: bounded(label, 1000) }),
          this.now() + 15 * 60_000,
        )
      this.accountSave(lease.taskId, budget, plan)
      this.db
        .prepare('DELETE FROM execution_resources WHERE owner_task_id=? AND owner_attempt_id=?')
        .run(lease.taskId, lease.attemptId)
      this.notice(lease.taskId, this.transition(lease.taskId, 'waiting_user'), 'attention_required')
      return id
    })
  }
  acquireAfterDecision(lease: TaskLease): boolean {
    return this.tx(() => {
      const row = this.lease(lease)
      if (row.state !== 'waiting_user' || row.pending_control) throw new Error('task_stopping')
      const current = this.db
        .prepare('SELECT owner_attempt_id FROM execution_resources WHERE resource_key=?')
        .get(RESOURCE)
      if (current) return current.owner_attempt_id === lease.attemptId
      this.db
        .prepare(
          "INSERT INTO execution_resources(resource_key,device_id,owner_task_id,owner_attempt_id,owner_epoch,state,proof_json,updated_at) VALUES(?,?,?,?,?,'held','{}',?)",
        )
        .run(RESOURCE, TASK_EXECUTOR, lease.taskId, lease.attemptId, lease.epoch, this.now())
      return true
    })
  }
  answerDecision(
    lease: TaskLease,
    id: string,
    digest: string,
    allowed: boolean,
    targets: TaskTargetVersion[],
  ): boolean {
    return this.tx(() => {
      const row = this.lease(lease, !allowed),
        { budget, plan } = this.account(row)
      const current = this.db
        .prepare('SELECT * FROM task_confirmations WHERE id=? AND task_id=?')
        .get(taskId(id), lease.taskId)
      if (!current || current.state !== 'pending' || current.subject_digest !== digest)
        throw new Error('stale_task_confirmation')
      const resource = this.db
        .prepare('SELECT owner_attempt_id FROM execution_resources WHERE resource_key=?')
        .get(RESOURCE)
      const valid =
        allowed &&
        resource?.owner_attempt_id === lease.attemptId &&
        row.state === 'waiting_user' &&
        !row.pending_control &&
        Number(current.expires_at) > this.now() &&
        current.context_revisions_json === json(targets)
      this.db
        .prepare(
          'UPDATE task_confirmations SET state=?,response_request_id=?,response_digest=?,response_json=?,resolved_at=? WHERE id=?',
        )
        .run(
          valid ? 'accepted' : allowed ? 'invalidated' : 'rejected',
          randomUUID(),
          taskHash({ allowed, digest, targets }),
          json({ allowed: valid }),
          this.now(),
          id,
        )
      if (current.step_id)
        this.db
          .prepare(
            "UPDATE task_steps SET state='executing',effect_state='started',revision=revision+1 WHERE id=? AND state='prepared'",
          )
          .run(String(current.step_id))
      this.db
        .prepare(
          "UPDATE task_outbox SET state='cancelled' WHERE task_id=? AND kind='attention_required' AND state IN ('pending','delivered','inflight')",
        )
        .run(lease.taskId)
      this.accountSave(lease.taskId, budget, plan)
      if (row.state === 'waiting_user')
        this.transition(
          lease.taskId,
          'running',
          allowed && !valid ? 'confirmation_stale' : undefined,
        )
      this.event(lease.taskId, 'confirmation.resolved', valid ? 'Allowed' : 'Not allowed', [id])
      return valid
    })
  }
  finish(
    lease: TaskLease,
    result: {
      state: 'completed' | 'failed' | 'cancelled' | 'waiting_user'
      reason?: string
      question?: string
    },
    snapshot?: AgentSessionSnapshot,
  ): AssistantTaskView {
    return this.tx(() => {
      const row = this.lease(lease, true),
        { budget, plan } = this.account(row),
        id = lease.taskId
      this.db
        .prepare(
          "UPDATE task_steps SET state='superseded',effect_state='not_applied' WHERE task_id=? AND (state='prepared' OR (state='executing' AND recovery='repeat_read'))",
        )
        .run(id)
      this.db
        .prepare(
          "UPDATE task_steps SET state='awaiting_reconciliation',effect_state='unknown' WHERE task_id=? AND state='executing'",
        )
        .run(id)
      const uncertain = this.uncertain(id)
      const state: AssistantTaskState = uncertain
        ? 'interrupted'
        : row.pending_control === 'pause'
          ? 'paused'
          : row.pending_control === 'cancel'
            ? 'cancelled'
            : result.reason === 'executor_lost'
              ? 'interrupted'
              : result.state
      if (snapshot && !uncertain && ['completed', 'waiting_user'].includes(state))
        this.checkpoint(lease, snapshot)
      this.accountSave(id, budget, {
        ...plan,
        accountedAt: undefined,
        resumeRequested: state === 'interrupted' && !uncertain && result.reason === 'executor_lost',
      })
      this.invalidateConfirmations(id)
      this.db
        .prepare('UPDATE task_attempts SET state=?,ended_at=? WHERE id=?')
        .run(
          state === 'paused' || state === 'waiting_user'
            ? 'suspended'
            : state === 'interrupted'
              ? 'lost'
              : state,
          this.now(),
          lease.attemptId,
        )
      this.release(id, state === 'interrupted' || state === 'waiting_user')
      if (!uncertain) this.db.prepare('UPDATE tasks SET pending_control=NULL WHERE id=?').run(id)
      const sequence = this.transition(id, state, uncertain ? 'outcome_unknown' : result.reason)
      if (state === 'waiting_user') {
        const question = bounded(result.question, 1000).trim()
        if (!question || !snapshot) throw new Error('invalid_task_question')
        const confirmation = randomUUID()
        this.db
          .prepare(
            "INSERT INTO task_confirmations(id,task_id,kind,state,revision,subject_digest,authority_revision,context_revisions_json,question_json) VALUES(?,?,'clarification','pending',?,?,1,?,?)",
          )
          .run(
            confirmation,
            id,
            Number(this.row(id).revision),
            taskHash([id, lease.epoch, question]),
            json(plan.lastTargets ?? []),
            json({ text: question }),
          )
        this.event(id, 'confirmation.requested', 'Question awaiting a reply', [confirmation])
        this.notice(id, sequence, 'attention_required')
      }
      if (state === 'completed') this.notice(id, sequence, 'result_ready')
      else if (state === 'failed' || state === 'interrupted')
        this.notice(id, sequence, 'task_blocked')
      return this.get(id)
    })
  }
  answer(input: AssistantTaskAnswer): { task: AssistantTaskView; duplicate: boolean } {
    const id = taskId(input.taskId),
      confirmation = taskId(input.confirmationId),
      requestId = taskId(input.requestId)
    const text = bounded(input.answer, 8_000).trim(),
      digest = taskHash(input)
    if (!text || text.length !== input.answer.trim().length) throw new Error('invalid_task_answer')
    return this.tx(() => {
      const prior = this.db
        .prepare(
          "SELECT request_hash,result_json FROM task_commands WHERE task_id=? AND principal_id='local-ui' AND request_id=?",
        )
        .get(id, requestId)
      if (prior) {
        if (prior.request_hash !== digest) throw new Error('request_conflict')
        return { task: parse(prior.result_json), duplicate: true }
      }
      const row = this.row(id),
        current = this.db
          .prepare(
            "SELECT * FROM task_confirmations WHERE id=? AND task_id=? AND kind='clarification'",
          )
          .get(confirmation, id)
      if (
        row.state !== 'waiting_user' ||
        row.revision !== input.expectedRevision ||
        !current ||
        current.state !== 'pending' ||
        this.db.prepare('SELECT 1 FROM task_leases WHERE task_id=?').get(id)
      )
        throw new Error('stale_task_confirmation')
      this.db
        .prepare(
          "UPDATE task_confirmations SET state='accepted',response_request_id=?,response_digest=?,response_json=?,resolved_at=? WHERE id=?",
        )
        .run(requestId, digest, json({ answer: text }), this.now(), confirmation)
      this.db.prepare('UPDATE tasks SET plan_json=? WHERE id=?').run(
        json({
          ...this.plan(id),
          resumeRequested: true,
          reply: { confirmationId: confirmation, text },
        }),
        id,
      )
      this.invalidateConfirmations(id)
      this.transition(id, 'queued')
      this.event(id, 'confirmation.resolved', 'Reply accepted', [confirmation])
      const task = this.get(id)
      this.db
        .prepare(
          "INSERT INTO task_commands(task_id,principal_id,request_id,request_hash,result_json,created_at) VALUES(?,'local-ui',?,?,?,?)",
        )
        .run(id, requestId, digest, json(task), this.now())
      return { task, duplicate: false }
    })
  }
  reconciliationTargets(id: string): TaskTargetVersion[] {
    const uncertain = this.db
      .prepare(
        "SELECT targets_json FROM task_steps WHERE task_id=? AND state='awaiting_reconciliation'",
      )
      .all(id)
    if (!uncertain.length) {
      const document = this.request(id).document
      return this.plan(id).lastTargets ?? (document ? [document] : [])
    }
    const targets = new Map<string, TaskTargetVersion>()
    for (const row of uncertain) {
      for (const target of parse<TaskTargetVersion[]>(row.targets_json))
        targets.set(target.path, target)
    }
    return [...targets.values()]
  }
  private resumeBlocker(row: Row): string | undefined {
    const budget = parse<AssistantTaskBudget>(row.budget_json)
    if (
      budget.usedAttempts >= budget.maxAttempts ||
      budget.usedModelTurns >= budget.maxModelTurns ||
      budget.usedActiveMs >= budget.maxActiveMs
    )
      return 'task_budget_exhausted'
    const later = this.db
      .prepare(
        "SELECT 1 FROM tasks WHERE session_id=? AND id<>? AND created_at>? AND json_extract(budget_json,'$.usedModelTurns')>0 LIMIT 1",
      )
      .get(String(row.session_id), String(row.id), Number(row.created_at))
    return later ? 'task_conversation_changed' : undefined
  }
  control(input: AssistantTaskControl, reviewedTargets?: TaskTargetVersion[]): AssistantTaskView {
    return this.tx(() => {
      const id = taskId(input.taskId),
        requestId = taskId(input.requestId),
        row = this.row(id),
        digest = taskHash(input)
      const old = this.db
        .prepare(
          "SELECT request_hash,result_json FROM task_commands WHERE task_id=? AND principal_id='local-ui' AND request_id=?",
        )
        .get(id, requestId)
      if (old) {
        if (old.request_hash !== digest) throw new Error('request_conflict')
        return parse(old.result_json)
      }
      if (row.revision !== input.expectedRevision) throw new Error('stale_task_revision')
      const state = row.state as AssistantTaskState,
        active = !!this.db.prepare('SELECT 1 FROM task_leases WHERE task_id=?').get(id),
        uncertain = this.uncertain(id)
      if (input.action === 'reconcile') {
        if (
          active ||
          (!uncertain &&
            !['task_context_changed', 'task_unsaved_context_unavailable'].includes(
              String(row.reason_code),
            )) ||
          input.reviewed !== true
        )
          throw new Error('task_review_required')
        const expected = this.reconciliationTargets(id)
        if (
          expected.length &&
          (!reviewedTargets ||
            reviewedTargets.length !== expected.length ||
            reviewedTargets.some(
              (target) =>
                !target.complete || !expected.some((previous) => previous.path === target.path),
            ))
        )
          throw new Error('task_review_required')
        // The operator reviewed uncertain effects. The original call remains
        // recorded as superseded, and its digest is never replayed automatically.
        this.db
          .prepare(
            "UPDATE task_steps SET state='superseded',effect_state='unknown',revision=revision+1 WHERE task_id=? AND state='awaiting_reconciliation'",
          )
          .run(id)
        const control = row.pending_control
        const blocker = input.continueAfterReview ? this.resumeBlocker(row) : undefined
        const resume = input.continueAfterReview === true && control !== 'cancel' && !blocker
        this.release(id, resume)
        this.db.prepare('UPDATE tasks SET pending_control=NULL,plan_json=? WHERE id=?').run(
          json({
            ...parse<Plan>(row.plan_json),
            lastTargets: [
              ...new Map(
                [...(this.plan(id).lastTargets ?? []), ...(reviewedTargets ?? [])].map((target) => [
                  target.path,
                  target,
                ]),
              ).values(),
            ],
            resumeRequested: resume,
          }),
          id,
        )
        this.transition(
          id,
          control === 'cancel' ? 'cancelled' : resume ? 'queued' : 'paused',
          blocker ?? 'effects_reviewed',
        )
      } else if (input.action === 'pause' || input.action === 'cancel') {
        if (!TASK_TERMINAL_STATES.has(state)) {
          if (active || uncertain) {
            const { budget, plan } = this.account(row)
            this.accountSave(id, budget, plan)
            this.db.prepare('UPDATE tasks SET pending_control=? WHERE id=?').run(input.action, id)
            this.transition(
              id,
              active ? (input.action === 'pause' ? 'pausing' : 'cancelling') : 'interrupted',
              uncertain ? 'outcome_unknown' : undefined,
            )
          } else {
            this.release(id)
            this.invalidateConfirmations(id)
            this.transition(id, input.action === 'pause' ? 'paused' : 'cancelled')
          }
        }
      } else if (input.action === 'resume' || input.action === 'retry') {
        if (
          active ||
          uncertain ||
          ![
            input.action === 'retry' ? 'failed' : 'paused',
            'interrupted',
            'waiting_device',
            'waiting_resource',
          ].includes(state)
        )
          throw new Error('task_cannot_resume')
        const blocker = this.resumeBlocker(row)
        if (blocker) throw new Error(blocker)
        const plan = parse<Plan>(row.plan_json)
        this.db.prepare('UPDATE tasks SET plan_json=?,pending_control=NULL WHERE id=?').run(
          json({
            ...plan,
            resumeRequested: true,
            ...(input.action === 'retry' ? { retryRequestId: input.requestId } : {}),
          }),
          id,
        )
        this.transition(id, 'queued')
      } else throw new Error('invalid_task_control')
      const result = this.get(id)
      this.db
        .prepare(
          "INSERT INTO task_commands(task_id,principal_id,request_id,request_hash,result_json,created_at) VALUES(?,'local-ui',?,?,?,?)",
        )
        .run(id, requestId, digest, json(result), this.now())
      return result
    })
  }
  get(id: string): AssistantTaskView {
    const row = this.row(id),
      request = parse<AssistantTaskRequest>(row.goal_json)
    const pending = this.db
      .prepare(
        "SELECT id,kind,revision,question_json FROM task_confirmations WHERE task_id=? AND state='pending' ORDER BY rowid DESC LIMIT 1",
      )
      .get(id)
    return {
      id,
      sessionId: String(row.session_id),
      entry: request.entry,
      title: String(row.title),
      state: row.state as AssistantTaskState,
      revision: Number(row.revision),
      createdAt: Number(row.created_at),
      updatedAt: Number(row.updated_at),
      budget: parse(row.budget_json),
      ...(row.reason_code ? { reason: String(row.reason_code) } : {}),
      ...(row.pending_control ? { pendingControl: row.pending_control as 'pause' | 'cancel' } : {}),
      ...(['outcome_unknown', 'task_context_changed', 'task_unsaved_context_unavailable'].includes(
        String(row.reason_code),
      )
        ? {
            reviewTargets: this.reconciliationTargets(id).map((item) => item.path),
          }
        : {}),
      steps: this.db
        .prepare(
          'SELECT id,label,state,effect_state FROM task_steps WHERE task_id=? ORDER BY display_order',
        )
        .all(id)
        .map((step) => ({
          id: String(step.id),
          label: String(step.label),
          state: String(step.state),
          effect: String(step.effect_state),
        })),
      evidenceIds: this.db
        .prepare('SELECT id FROM task_evidence WHERE task_id=? ORDER BY rowid')
        .all(id)
        .map((item) => String(item.id)),
      ...(pending
        ? {
            confirmation: {
              id: String(pending.id),
              kind: pending.kind as 'approval' | 'clarification' | 'conflict',
              revision: Number(pending.revision),
              text: parse<{ text: string }>(pending.question_json).text,
            },
          }
        : {}),
    }
  }
  evidence(id: string): unknown[] {
    this.row(id)
    return this.db
      .prepare('SELECT evidence_json FROM task_evidence WHERE task_id=? ORDER BY rowid')
      .all(id)
      .map((row) => parse(row.evidence_json))
  }
  /** Latest host-observed writer receipt for each deliverable, invalidated by subsequent changes. */
  saveTargets(id: string): Array<{ path: string; saved?: TaskTargetVersion }> {
    const targets = new Map<
      string,
      { path: string; saved?: TaskTargetVersion; identity?: string }
    >()
    const steps = this.db
      .prepare(
        "SELECT label,state,effect_state,receipt_json FROM task_steps WHERE task_id=? AND receipt_json IS NOT NULL AND recovery<>'repeat_read' ORDER BY display_order",
      )
      .all(id)
    for (const step of steps) {
      const receipt = parse<TaskStepReceipt>(step.receipt_json)
      const writer =
        /^(?:save_document|save_workbook|save|(?:office|docx|xlsx|pptx|md)\.save|(?:docx|xlsx|pptx|md|html)\.create|office\.export|pdf\.export_pages)$/.test(
          String(step.label),
        )
      if (!writer && step.effect_state !== 'applied') continue
      for (const target of receipt.targets) {
        if (!target.type || target.type === 'unknown') continue
        if (target.editorIdentity && target.exists)
          for (const [key, prior] of targets)
            if (prior.identity === target.editorIdentity && key !== target.path && !prior.saved)
              targets.delete(key)
        targets.set(target.path, {
          path: target.path,
          identity: target.editorIdentity,
          ...((writer || (step.label === 'create_document' && target.generatedFile)) &&
          receipt.wordSaveVerified !== false &&
          step.state === 'verified' &&
          target.complete &&
          target.exists &&
          target.diskDigest &&
          !target.dirty
            ? { saved: target }
            : {}),
        })
      }
    }
    return [...targets.values()]
  }

  unfulfilledOperation(id: string): 'approval_expired' | 'operation_not_completed' | undefined {
    this.row(id)
    const operations = this.db
      .prepare(
        "SELECT state,receipt_json FROM task_steps WHERE task_id=? AND recovery<>'repeat_read' ORDER BY display_order",
      )
      .all(id)
    if (
      !operations.some((row) => row.state === 'failed') ||
      operations.some((row) => row.state !== 'failed')
    )
      return undefined
    const receipt = parse<{ result?: { output?: string } }>(operations.at(-1)!.receipt_json)
    return receipt?.result?.output?.includes('Approval expired before authorization.')
      ? 'approval_expired'
      : 'operation_not_completed'
  }
  feed(sessionId?: string, afterCursor?: number): AssistantTaskFeed {
    return this.tx(() => {
      const filter = sessionId ? ' AND t.session_id=?' : '',
        args = sessionId ? [taskId(sessionId)] : []
      const tasks = this.db
        .prepare(
          "SELECT t.id FROM tasks t WHERE t.kind='assistant'" +
            filter +
            " ORDER BY CASE WHEN t.state IN ('completed','failed','cancelled') THEN 1 ELSE 0 END,t.updated_at DESC,t.rowid DESC LIMIT 100",
        )
        .all(...args)
        .map((row) => this.get(String(row.id)))
      const notices = this.db
        .prepare(
          "SELECT o.* FROM task_outbox o JOIN tasks t ON t.id=o.task_id WHERE t.kind='assistant' AND o.route='local' AND o.state IN ('pending','delivered')" +
            filter +
            ' ORDER BY o.rowid DESC LIMIT 100',
        )
        .all(...args)
      for (const notice of notices)
        if (notice.state === 'pending') {
          const stamp = this.now()
          this.db
            .prepare(
              "UPDATE task_outbox SET state='delivered',attempts=attempts+1,delivery_epoch=delivery_epoch+1,receipt_id=?,delivered_at=? WHERE id=?",
            )
            .run(String(notice.id), stamp, String(notice.id))
          notice.delivered_at = stamp
        }
      const throughCursor = Number(
        this.db
          .prepare(
            "SELECT COALESCE((SELECT seq FROM sqlite_sequence WHERE name='task_events'),0) AS n",
          )
          .get()!.n,
      )
      const invalid =
        afterCursor !== undefined &&
        (!Number.isSafeInteger(afterCursor) || afterCursor < 0 || afterCursor > throughCursor)
      const events =
        afterCursor !== undefined && !invalid
          ? this.db
              .prepare(
                "SELECT e.* FROM task_events e JOIN tasks t ON t.id=e.task_id WHERE t.kind='assistant' AND e.feed_cursor>?" +
                  filter +
                  ' ORDER BY e.feed_cursor LIMIT 200',
              )
              .all(afterCursor, ...args)
              .map((event) => ({
                cursor: Number(event.feed_cursor),
                taskId: String(event.task_id),
                sequence: Number(event.sequence),
                revision: Number(event.revision),
                type: String(event.type),
                summary: String(event.summary),
                at: Number(event.created_at),
              }))
          : undefined
      return {
        tasks,
        notices: notices.map((notice) => ({
          id: String(notice.id),
          taskId: String(notice.task_id),
          kind: notice.kind as AssistantTaskFeed['notices'][number]['kind'],
          deliveredAt: Number(notice.delivered_at ?? this.now()),
        })),
        throughCursor,
        ...(events
          ? { events, eventCursor: events.length === 200 ? events.at(-1)!.cursor : throughCursor }
          : {}),
        ...(invalid ? { resyncRequired: true } : {}),
      }
    })
  }
  acknowledge(id: string): void {
    this.db
      .prepare(
        "UPDATE task_outbox SET state='acked',acknowledged_at=?,read_at=? WHERE id=? AND route='local' AND state='delivered'",
      )
      .run(this.now(), this.now(), taskId(id))
  }
  forSession(sessionId: string): AssistantTaskView[] {
    return this.db
      .prepare("SELECT id FROM tasks WHERE kind='assistant' AND session_id=? ORDER BY rowid DESC")
      .all(taskId(sessionId))
      .map((row) => this.get(String(row.id)))
  }
  sessionPurged(id: string): boolean {
    return !!this.db.prepare('SELECT 1 FROM session_tombstones WHERE session_id=?').get(taskId(id))
  }
  pendingPurges(): Array<{ id: string; sessions: string[] }> {
    return this.db
      .prepare(
        "SELECT id,scope_json FROM lifecycle_jobs WHERE kind='manual_session_purge' AND state IN ('applying','failed') ORDER BY rowid",
      )
      .all()
      .map((row) => ({ id: String(row.id), sessions: parse<string[]>(row.scope_json) }))
  }
  completePurge(id: string): void {
    this.db
      .prepare(
        "UPDATE lifecycle_jobs SET state='completed',phase='done',updated_at=? WHERE id=? AND kind='manual_session_purge'",
      )
      .run(this.now(), taskId(id))
  }
  purgeSessions(sessionIds: string[], confirmationRef: string): string {
    return this.tx(() => {
      const ids = [...new Set(sessionIds.map(taskId))],
        scope = json(ids),
        stamp = this.now()
      const existing = this.db
        .prepare(
          "SELECT id FROM lifecycle_jobs WHERE scope_json=? AND state IN ('applying','failed')",
        )
        .get(scope)
      const job = existing ? String(existing.id) : randomUUID()
      if (!existing)
        this.db
          .prepare(
            "INSERT INTO lifecycle_jobs(id,kind,scope_json,confirmation_ref,state,phase,created_at,updated_at) VALUES(?,'manual_session_purge',?,?,'applying','task_records',?,?)",
          )
          .run(job, scope, taskId(confirmationRef), stamp, stamp)
      for (const session of ids) {
        const tasks = this.forSession(session)
        if (
          tasks.some(
            (task) =>
              this.db.prepare('SELECT 1 FROM task_leases WHERE task_id=?').get(task.id) ||
              this.uncertain(task.id),
          )
        )
          throw new Error('task_review_required')
        this.db
          .prepare(
            'INSERT OR IGNORE INTO session_tombstones(session_id,purge_job_id,purged_at) VALUES(?,?,?)',
          )
          .run(session, job, stamp)
        for (const task of tasks) {
          const row = this.row(task.id)
          this.db
            .prepare(
              'INSERT INTO task_tombstones(task_id,authority_id,principal_id,request_id,request_hash,purge_job_id,purged_at) VALUES(?,?,?,?,?,?,?)',
            )
            .run(
              task.id,
              String(row.authority_id),
              String(row.principal_id),
              String(row.request_id),
              String(row.request_hash),
              job,
              stamp,
            )
          this.release(task.id)
          for (const table of [
            'task_outbox',
            'task_checkpoints',
            'task_evidence',
            'task_confirmations',
            'task_artifacts',
          ])
            this.db.prepare(`DELETE FROM ${table} WHERE task_id=?`).run(task.id)
          this.db.prepare('DELETE FROM tasks WHERE id=?').run(task.id)
        }
      }
      this.db
        .prepare(
          "UPDATE lifecycle_jobs SET state='applying',phase='task_records_removed',updated_at=? WHERE id=?",
        )
        .run(this.now(), job)
      return job
    })
  }
}
