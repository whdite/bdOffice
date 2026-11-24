import type {
  AgentSessionSnapshot,
  AgentTokenUsage,
  AgentToolCall,
  TaskExecutionEnd,
  TaskStartInput,
  TaskToolPermit,
  ToolExecution,
} from '@genoffice/agent-core'
import type { ResolveChatResult } from './ipc.js'
import type { PortableResultRecord } from './result-evidence.js'

export const FILE_TASK_CHANNEL = 'project:file-task'
export interface FileTaskItem {
  id: string
  title: string
  state: string
  revision: number
  reason?: string
  pendingControl?: 'pause' | 'cancel'
  reviewTargets?: string[]
  attempts: number
  resumeRequested: boolean
  confirmation?: { id: string; kind: string; text: string } | undefined
}
export interface FileTaskPanelState {
  tasks: FileTaskItem[]
  results: PortableResultRecord[]
}
export interface FileTaskRequests {
  probe: ResolveChatResult
  panel: ResolveChatResult
  openArtifact: ResolveChatResult & { artifactId: string }
  control: ResolveChatResult & {
    taskId: string
    requestId: string
    expectedRevision: number
    action: 'pause' | 'resume' | 'cancel' | 'retry' | 'reconcile'
    reviewed?: boolean
    continueAfterReview?: boolean
  }
  answer: ResolveChatResult & {
    taskId: string
    confirmationId: string
    requestId: string
    expectedRevision: number
    answer: string
  }
  begin: ResolveChatResult & {
    writerId: string
    requestId: string
    input: TaskStartInput
    resumeId?: string
  }
  abortBegin: ResolveChatResult & { writerId: string; requestId: string; resumeId?: string }
  reserveModel: RunBound & { checkpoint: AgentSessionSnapshot }
  usage: RunBound & { modelCallId: string; usage: AgentTokenUsage }
  beforeApproval: RunBound & { call: AgentToolCall }
  afterApproval: RunBound & { call: AgentToolCall; allowed: boolean; round?: boolean }
  prepareTool: RunBound & { call: AgentToolCall }
  finishTool: RunBound & {
    permit: TaskToolPermit
    result: ToolExecution
    checkpoint: AgentSessionSnapshot
  }
  pendingWordParts: RunBound
  checkSaved: RunBound
  finish: RunBound & { result: TaskExecutionEnd; checkpoint: AgentSessionSnapshot }
}
type RunBound = ResolveChatResult & { writerId: string; token: string }
export interface FileTaskReplies {
  probe: true
  panel: FileTaskPanelState
  openArtifact: void
  control: FileTaskItem
  answer: FileTaskItem
  begin: {
    id: string
    token: string
    checkpoint?: AgentSessionSnapshot
    remainingMs: number
    memoryContext?: string
  }
  abortBegin: void
  reserveModel: string
  usage: void
  beforeApproval: void | import('@genoffice/agent-core').TaskApprovalScope
  afterApproval: void
  prepareTool: TaskToolPermit
  finishTool: void
  pendingWordParts: string[]
  checkSaved: boolean | undefined
  finish: void
}
export type FileTaskCall = <K extends keyof FileTaskRequests>(
  action: K,
  input: FileTaskRequests[K],
) => Promise<FileTaskReplies[K]>
export function missingFileTaskChannel(error: unknown): boolean {
  return (
    error instanceof Error &&
    /No handler registered for ['"]project:file-task['"]/i.test(error.message)
  )
}
