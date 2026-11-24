import type { AgentInput, WorkspaceFileType } from './workspace-agent-api'

export type AssistantTaskState =
  | 'queued'
  | 'waiting_device'
  | 'waiting_resource'
  | 'preparing'
  | 'running'
  | 'waiting_user'
  | 'verifying'
  | 'pausing'
  | 'paused'
  | 'cancelling'
  | 'interrupted'
  | 'completed'
  | 'failed'
  | 'cancelled'
export interface TaskTargetVersion {
  /** A separate file attested by the app writer during this task. */
  generatedFile?: boolean
  path: string
  type?: WorkspaceFileType
  exists: boolean
  /** Filesystem identity and byte revision; an editor identity is valid only for this boot. */
  fileIdentity?: string
  diskDigest?: string
  editorIdentity?: string
  modelDigest?: string
  dirty?: boolean
  /** Only a document with a complete observation may certify a file effect. */
  complete?: boolean
}
export interface AssistantTaskRequest {
  requestId: string
  sessionId: string
  entry: 'workspace' | 'file'
  input: AgentInput
  document?: TaskTargetVersion
  providerRevision?: string
  requiredResult: 'answer' | 'editor_change' | 'saved_artifact'
}
export interface AssistantTaskBudget {
  maxAttempts: number
  maxActiveMs: number
  maxModelTurns: number
  usedAttempts: number
  usedActiveMs: number
  usedModelTurns: number
  usedReportedTokens: number
  unreportedModelCalls: number
}
export interface AssistantTaskView {
  id: string
  sessionId: string
  entry: AssistantTaskRequest['entry']
  title: string
  state: AssistantTaskState
  revision: number
  reason?: string
  pendingControl?: 'pause' | 'cancel'
  createdAt: number
  updatedAt: number
  budget: AssistantTaskBudget
  reviewTargets?: string[]
  steps: Array<{ id: string; label: string; state: string; effect: string }>
  confirmation?: {
    id: string
    kind: 'approval' | 'clarification' | 'conflict'
    revision: number
    text: string
  }
  evidenceIds: string[]
}
export interface AssistantTaskControl {
  taskId: string
  requestId: string
  expectedRevision: number
  action: 'pause' | 'resume' | 'cancel' | 'retry' | 'reconcile'
  /** An explicit decision after reviewing uncertain effects. No old grant is restored. */
  reviewed?: boolean
  continueAfterReview?: boolean
}
export interface AssistantTaskAnswer {
  taskId: string
  confirmationId: string
  requestId: string
  expectedRevision: number
  answer: string
}
export interface AssistantTaskNotice {
  id: string
  taskId: string
  kind: 'attention_required' | 'result_ready' | 'task_blocked'
  deliveredAt: number
}
export interface AssistantTaskFeed {
  tasks: AssistantTaskView[]
  notices: AssistantTaskNotice[]
  throughCursor: number
  events?: Array<{
    cursor: number
    taskId: string
    sequence: number
    revision: number
    type: string
    summary: string
    at: number
  }>
  eventCursor?: number
  resyncRequired?: boolean
}
export interface AssistantTasksApi {
  feed(sessionId?: string, afterCursor?: number): Promise<AssistantTaskFeed>
  evidence(taskId: string): Promise<import('@genoffice/project-store').PortableResultRecord[]>
  control(input: AssistantTaskControl): Promise<AssistantTaskView>
  acknowledge(id: string): Promise<void>
  answer(input: AssistantTaskAnswer): Promise<AssistantTaskView>
  openSession(taskId: string): Promise<void>
  openArtifact(sessionId: string, artifactId: string): Promise<void>
}
export const ASSISTANT_TASK_CHANNELS = {
  feed: 'assistant-tasks:feed',
  evidence: 'assistant-tasks:evidence',
  control: 'assistant-tasks:control',
  answer: 'assistant-tasks:answer',
  acknowledge: 'assistant-tasks:acknowledge',
  openSession: 'assistant-tasks:open-session',
  openArtifact: 'assistant-tasks:open-artifact',
} as const
export const TASK_TERMINAL_STATES: ReadonlySet<AssistantTaskState> = new Set([
  'completed',
  'failed',
  'cancelled',
])
