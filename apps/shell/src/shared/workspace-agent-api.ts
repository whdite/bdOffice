import type { AgentImage, AgentMessage } from '@genoffice/agent-core'

export type WorkspaceFileType =
  | 'docx'
  | 'pptx'
  | 'xlsx'
  | 'xlsm'
  | 'pdf'
  | 'txt'
  | 'md'
  | 'html'
  | 'csv'
  | 'json'
  | 'png'
  | 'jpg'
  | 'jpeg'
  | 'webp'
  | 'directory'
  | 'unknown'

export interface WorkspaceDescriptor {
  id: string
  root: string
  name: string
  createdAt: string
  updatedAt: string
  /** Where the Agent workspace came from. Legacy descriptors omit this and
   * are treated as custom folder workspaces. */
  source?: 'folder' | 'project'
  managedProjectId?: string
  /** ProjectStore id when `source` is `project`. */
  projectId?: string
  /** Physical files currently associated with a ProjectStore workspace. This
   * is kept in the local descriptor so a session can rehydrate without
   * copying the user's files into a staging directory. */
  projectFiles?: string[]
  /** Stable binding to the Harness workspace represented by this directory. */
  harness?: {
    provider: string
    workspaceId: string
    root: string
  }
}

/** A project-space option surfaced by the existing GenOffice ProjectStore. */
export interface WorkspaceProjectOption {
  id: string
  name: string
  fileCount: number
  isDefault: boolean
}

export interface WorkspaceFileEntry {
  path: string
  name: string
  type: WorkspaceFileType
  sizeBytes: number
  mtimeMs: number
  directory: boolean
  /** A physical path registered for global discovery rather than a child of this working folder. */
  registered?: boolean
}

export interface WorkspaceFileContent {
  path: string
  type: WorkspaceFileType
  text?: string
  base64?: string
  mime?: string
  truncated?: boolean
}

export interface WorkspaceFileMetadata extends WorkspaceFileEntry {
  opened: boolean
  dirty: boolean
  office?: Record<string, unknown>
}

export interface WorkspaceSearchResult {
  path: string
  file: string
  type: WorkspaceFileType
  preview: string
  line?: number
  location?: Record<string, unknown>
  score?: number
}

export interface WorkspaceEditorSelection {
  /** Selected text is intentionally bounded by the main-process validator. */
  text?: string
  from?: number
  to?: number
  elementIds?: string[]
}

/** Small live-editor state used by mentions. It is a reference, not a document dump. */
export interface WorkspaceEditorContext {
  selection?: WorkspaceEditorSelection
  currentPage?: number
  currentSlide?: number
  currentSheet?: string
  currentRange?: string
}

export interface WorkspaceAgentContext {
  documentContext?: import('./document-context-api').DocumentContextView
  documentContextError?: string
  workspaceId: string
  workspaceName: string
  activeFile?: {
    path: string
    name: string
    type: WorkspaceFileType
    opened: boolean
    dirty: boolean
    editor?: WorkspaceEditorContext
  }
}

export type WorkspaceAgentReferenceKind =
  | 'workspace'
  | 'file'
  | 'current-document'
  | 'current-selection'
  | 'current-page'
  | 'current-sheet-range'

export interface WorkspaceAgentReference {
  kind: WorkspaceAgentReferenceKind
  label: string
  path?: string
  type?: WorkspaceFileType
  selection?: WorkspaceEditorSelection
  page?: number
  slide?: number
  sheet?: string
  range?: string
}

export type AgentPlanStepStatus = 'pending' | 'running' | 'completed' | 'failed'

export interface AgentPlanStep {
  id: string
  title: string
  status: AgentPlanStepStatus
}

export type ToolPermissionLevel = 'write' | 'destructive' | 'shell'
export type ToolPermissionDecision = 'allow-once' | 'allow-round' | 'deny'
/** Auto executes all registry-approved tools; per-action asks before every mutation or command. */
export type AgentApprovalMode = 'auto' | 'per-action'

export function isAgentApprovalMode(value: unknown): value is AgentApprovalMode {
  return value === 'auto' || value === 'per-action'
}

export type AgentRunChangeOperation = 'modified' | 'created' | 'deleted'
export type AgentRunReviewState = 'completed' | 'undone' | 'undo-failed'

export interface AgentRunFileChange {
  path: string
  /** Resolved by the host for historical document association, not file authority. */
  documentPath?: string
  /** Host-observed identity and first-save lineage, never an imported authority. */
  documentId?: string
  previousPath?: string
  type: WorkspaceFileType
  operation: AgentRunChangeOperation
  tools: string[]
  summaries: string[]
  undoable: boolean
  before?: Record<string, unknown>
  after?: Record<string, unknown>
  evidence?: import('./change-evidence').FileChangeEvidence
  artifactId?: string
  undoState?: 'restored' | 'removed' | 'conflict' | 'not_attempted' | 'unknown'
}

/** Durable, UI-safe review record. Binary/model checkpoints stay in userData. */
export interface AgentRunReview {
  id: string
  taskId?: string
  state: AgentRunReviewState
  startedAt: string
  completedAt: string
  undoneAt?: string
  undoable: boolean
  changes: AgentRunFileChange[]
  error?: string
  outcome?: 'completed' | 'partial' | 'failed' | 'cancelled'
  undoConflict?: boolean
  evidenceStored?: boolean
  historical?: boolean
}

export interface RunReviewPage {
  reviews: AgentRunReview[]
  nextCursor?: string
}

export interface AgentSessionSummary {
  id: string
  workspaceId: string
  title: string
  createdAt: string
  updatedAt: string
  status: 'idle' | 'running' | 'error'
  messageCount: number
  /** Stable Harness session identity (equal to id for the local adapter). */
  harness?: {
    provider: string
    sessionId: string
    workspaceId: string
    eventLogPath?: string
  }
}

export type WorkspaceAgentEventType =
  | 'usage.updated'
  | 'assistant.message'
  | 'assistant.reasoning_summary'
  | 'tool.started'
  | 'tool.completed'
  | 'tool.failed'
  | 'file.read'
  | 'file.modified'
  | 'document.created'
  | 'agent.plan'
  | 'agent.run.review'
  | 'permission.requested'
  | 'permission.resolved'
  | 'agent.status'
  | 'error'

export interface WorkspaceAgentEvent {
  usage?: import('@genoffice/agent-core').AgentTokenUsage
  id: string
  sessionId: string
  /** Monotonic per-session order for consumers that need deterministic replay. */
  sequence?: number
  /** Harness lifecycle identities for durable Turn/Step replay. */
  runId?: string
  turnId?: string
  stepId?: string
  type: WorkspaceAgentEventType
  timestamp: string
  /** Replace provisional assistant text when its proposed completion is corrected. */
  replaceText?: boolean
  text?: string
  role?: 'user' | 'assistant'
  tool?: {
    /** Stable provider call id used to match started/completed events. */
    callId?: string
    name: string
    input?: Record<string, unknown>
    summary?: string
    output?: string
    error?: boolean
    mutated?: boolean
    /** Optional UI-only display payload (for example a rendered slide image). */
    display?: {
      kind: 'images' | 'links' | 'text' | 'mindmap'
      tree?: import('@genoffice/agent-core').MindMapTreeNode
      items?: Array<{ url: string; title?: string; thumb?: string }>
      text?: string
    }
  }
  file?: {
    path: string
    type?: WorkspaceFileType
    operation?: 'read' | 'modified' | 'created'
  }
  plan?: {
    steps: AgentPlanStep[]
  }
  run?: AgentRunReview
  permission?: {
    id: string
    level: ToolPermissionLevel
    toolName: string
    title: string
    description: string
    state: 'pending' | 'allowed' | 'denied'
    expiresAt?: number
    reason?: 'expired' | 'cancelled'
    round?: { path: string; baseVersion: string; approved?: boolean }
  }
  status?: 'idle' | 'running' | 'cancelled' | 'error' | 'completed' | 'waiting_user'
  error?: string
}

export interface WorkspaceFilesChangedEvent {
  workspaceId: string
  timestamp: string
  eventType: 'rename' | 'change'
  path?: string
}

export interface AgentInput {
  requestId?: string
  requiredResult?: 'answer' | 'editor_change' | 'saved_artifact'
  contextId?: string
  instructions?: string
  text: string
  images?: AgentImage[]
  references?: WorkspaceAgentReference[]
}

export interface WorkspaceFileListRequest {
  path?: string
  recursive?: boolean
  types?: string[]
}

export interface WorkspaceSearchRequest {
  query: string
  path?: string
  fileTypes?: string[]
  limit?: number
}

export interface WorkspaceAgentApi {
  prepareContext(): Promise<WorkspaceAgentContext | null>
  clearDocumentContext(): Promise<void>
  getContextSelection(
    sessionId: string,
  ): Promise<import('@genoffice/project-store').SessionContextSelection>
  setContextSelection(
    sessionId: string,
    selection: import('@genoffice/project-store').SessionContextSelection,
  ): Promise<void>
  pickContextFiles(
    sessionId: string,
  ): Promise<import('@genoffice/project-store').SessionContextSelection>
  attachContextFiles(
    sessionId: string,
    paths: string[],
  ): Promise<
    import('@genoffice/project-store').SessionContextSelection & {
      addedPaths?: WorkspaceFileEntry[]
    }
  >
  previewContext(sessionId: string): Promise<AgentContextPreview>
  openHome(sessionId?: string): Promise<void>
  getWorkspace(): Promise<WorkspaceDescriptor | null>
  getContext(): Promise<WorkspaceAgentContext | null>
  openWorkspace(): Promise<WorkspaceDescriptor | null>
  setWorkspace(path: string): Promise<WorkspaceDescriptor>
  /** List the existing Office project spaces that can be bound to Agent. */
  listProjectWorkspaces(): Promise<WorkspaceProjectOption[]>
  /** Bind Agent to one existing ProjectStore space. */
  setProjectWorkspace(projectId: string): Promise<WorkspaceDescriptor>
  listFiles(request?: WorkspaceFileListRequest): Promise<WorkspaceFileEntry[]>
  getFile(path: string, maxChars?: number): Promise<WorkspaceFileContent>
  searchFiles(request: WorkspaceSearchRequest): Promise<WorkspaceSearchResult[]>
  getFileMetadata(path: string): Promise<WorkspaceFileMetadata>
  openFile(path: string): Promise<void>
  createSession(workspaceId?: string): Promise<AgentSessionSummary>
  /** Load an existing persisted session and return its current summary. */
  resumeSession(sessionId: string): Promise<AgentSessionSummary>
  listSessions(workspaceId?: string): Promise<AgentSessionSummary[]>
  getSessionMessages(sessionId: string): Promise<PersistedAgentMessage[]>
  getSessionHistoryPrefix?(sessionId: string): Promise<PersistedAgentMessage[]>
  /** Read the durable Harness event log, optionally after a sequence number. */
  getSessionEvents(sessionId: string, afterSequence?: number): Promise<WorkspaceAgentEvent[]>
  renameSession(sessionId: string, title: string): Promise<AgentSessionSummary>
  deleteSession(sessionId: string): Promise<void>
  send(sessionId: string, message: AgentInput): Promise<void>
  steer(sessionId: string, message: AgentInput): Promise<void>
  cancel(sessionId: string): Promise<void>
  /** Current global Agent approval policy, persisted in app settings. */
  getApprovalMode(): Promise<AgentApprovalMode>
  /** Change approval policy for subsequent tool calls, including the active run. */
  setApprovalMode(mode: AgentApprovalMode): Promise<AgentApprovalMode>
  onApprovalModeChanged?(handler: (mode: AgentApprovalMode) => void): () => void
  resolvePermission(
    sessionId: string,
    requestId: string,
    decision: ToolPermissionDecision,
  ): Promise<void>
  undoRun(sessionId: string, runId: string): Promise<AgentRunReview>
  openWordReview?(sessionId: string, runId: string, path: string): Promise<void>
  listRunReviews(sessionId: string, cursor?: string): Promise<RunReviewPage>
  subscribe(sessionId: string, handler: (event: WorkspaceAgentEvent) => void): () => void
  subscribeFiles?(
    workspaceId: string,
    handler: (event: WorkspaceFilesChangedEvent) => void,
  ): () => void
}

export const WORKSPACE_AGENT_CHANNELS = {
  runReviews: 'workspace-agent:run-reviews',
  openWordReview: 'workspace-agent:open-word-review',
  prepareContext: 'workspace-agent:prepare-context',
  clearDocumentContext: 'workspace-agent:clear-document-context',
  contextSelection: 'workspace-agent:context-selection',
  setContextSelection: 'workspace-agent:set-context-selection',
  pickContextFiles: 'workspace-agent:pick-context-files',
  attachContextFiles: 'workspace-agent:attach-context-files',
  previewContext: 'workspace-agent:preview-context',
  openHome: 'workspace-agent:open-home',
  getWorkspace: 'workspace-agent:get-workspace',
  getContext: 'workspace-agent:get-context',
  openWorkspace: 'workspace-agent:open-workspace',
  setWorkspace: 'workspace-agent:set-workspace',
  listProjectWorkspaces: 'workspace-agent:list-project-workspaces',
  setProjectWorkspace: 'workspace-agent:set-project-workspace',
  listFiles: 'workspace-agent:list-files',
  getFile: 'workspace-agent:get-file',
  searchFiles: 'workspace-agent:search-files',
  fileMetadata: 'workspace-agent:file-metadata',
  openFile: 'workspace-agent:open-file',
  createSession: 'workspace-agent:create-session',
  resumeSession: 'workspace-agent:resume-session',
  listSessions: 'workspace-agent:list-sessions',
  sessionMessages: 'workspace-agent:session-messages',
  sessionHistoryPrefix: 'workspace-agent:session-history-prefix',
  sessionEvents: 'workspace-agent:session-events',
  renameSession: 'workspace-agent:rename-session',
  deleteSession: 'workspace-agent:delete-session',
  send: 'workspace-agent:send',
  steer: 'workspace-agent:steer',
  cancel: 'workspace-agent:cancel',
  getApprovalMode: 'workspace-agent:get-approval-mode',
  setApprovalMode: 'workspace-agent:set-approval-mode',
  resolvePermission: 'workspace-agent:resolve-permission',
  undoRun: 'workspace-agent:undo-run',
  subscribe: 'workspace-agent:subscribe',
  unsubscribe: 'workspace-agent:unsubscribe',
  event: 'workspace-agent:event',
  subscribeFiles: 'workspace-agent:subscribe-files',
  unsubscribeFiles: 'workspace-agent:unsubscribe-files',
  filesChanged: 'workspace-agent:files-changed',
} as const
export interface AgentContextPreview {
  selection: import('@genoffice/project-store').SessionContextSelection
  memory: {
    sources: Array<{ id: string; title: string; characters: number; truncated: boolean }>
    omitted: number
    limit: number
  }
  files: Array<{ path: string; name: string; characters: number; error?: string }>
  characters: number
}

/** Persisted transcript uses the same AgentMessage shape as agent-core. */
export type PersistedAgentMessage = AgentMessage
