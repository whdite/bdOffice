/** JSON-Schema described tool exposed to the model */
export interface AgentToolDef {
  name: string
  description: string
  /** JSON Schema (object) describing the tool input */
  inputSchema: Record<string, unknown>
}

export interface AgentToolCall {
  id: string
  name: string
  input: Record<string, unknown>
  /** Parse error when the model emitted invalid input JSON; the loop feeds back an is_error result for retry instead of aborting the run */
  inputError?: string | undefined
  /** The argument stream was cut off by the token limit (stop_reason max_tokens); the loop asks the model to split the call instead of "fixing JSON" */
  truncated?: boolean | undefined
  /** Opaque provider continuation preserved in JSON history (e.g. thought signatures or namespaced Responses reasoning). */
  signature?: string | undefined
}

export interface AgentToolResult {
  id: string
  /** tool name (Gemini addresses function responses by name, not id) */
  name: string
  output: string
  /** Ephemeral observations for the current model run; excluded from saved history. */
  images?: AgentImage[]
  isError?: boolean | undefined
}

/** inline image attached to a user turn, fed to vision-capable providers as multimodal input */
export interface AgentImage {
  /** raw base64 (no data: URL prefix) */
  base64: string
  /** e.g. "image/png" */
  mime: string
}

/** Token usage one model turn was billed (best-effort: providers that do not
 * report usage simply never fire the callback — absent is "unknown", not zero). */
export interface AgentTokenUsage {
  inputTokens?: number
  outputTokens?: number
}

export type AgentMessage =
  | { role: 'user'; text: string; images?: AgentImage[] | undefined }
  /** reasoning: opaque vendor thinking captured during the turn; interleaved-thinking
   * models (MiniMax M3, DeepSeek V4) degrade in tool loops unless it is echoed back */
  | {
      role: 'assistant'
      text: string
      toolCalls?: AgentToolCall[] | undefined
      reasoning?: string | undefined
    }
  | { role: 'tool'; results: AgentToolResult[] }

/**
 * One node of a mind map tree (NotebookLM-style interactive diagram). `block`
 * optionally cites a document block index the node summarizes; the UI turns it
 * into a jump link. Labels are short phrases, not sentences.
 */
export interface MindMapTreeNode {
  label: string
  /** 0-based document block index this node draws on (optional) */
  block?: number
  children?: MindMapTreeNode[]
}

/**
 * Side-channel display data: UI-only, never merged into messages sent to the LLM.
 * kind='images' → image grid; kind='links' → link list; kind='text' → extra text;
 * kind='mindmap' → interactive mind map rendered from `tree`.
 */
export interface ToolDisplay {
  kind: 'images' | 'links' | 'text' | 'mindmap'
  /** entry list for images / links modes */
  items?: Array<{ url: string; title?: string; thumb?: string }>
  /** extra text for text mode */
  text?: string
  /** mind map root for the mindmap kind */
  tree?: MindMapTreeNode
}

export interface WordPartEvidence {
  path?: string
  expected?: Record<string, string>
  observed?: Record<string, string>
}

/** outcome of one tool execution */
export interface ToolExecution {
  /** result text fed back to the model */
  output: string
  /** Bounded tool observations, sent to vision-capable model adapters, never persisted. */
  images?: AgentImage[]
  isError?: boolean
  /** true when the tool changed the underlying artifact (document / sheet / deck) */
  mutated?: boolean
  savedFile?: { path: string; sha256: string }
  /** The Word writer observed no intervening user transaction while drafting. */
  exclusiveWordDraft?: boolean
  wordPart?: WordPartEvidence
  /** Host-reported effect on the operation's target, including Word parts outside the body editor. */
  effect?: import('./tool-capability').ToolEffect
  /** External action delivery, not proof of UI postconditions or an Office undo point. */
  externalEffect?: 'applied' | 'not_applied' | 'unknown'
  /** Host approval denied/cancelled the call; stop tools for this run. */
  denied?: boolean
  /** short human-readable label for activity UI */
  summary: string
  /**
   * Side-channel display: for UI only, never enters the LLM context.
   * Ignored when tool results are assembled into an AgentMessage.
   */
  display?: ToolDisplay
}

// ---- run phase (drives the in-progress status line in chat UIs) ----

export type AgentPhaseKind =
  /** request sent, waiting for the model's first content block */
  | 'requesting'
  | 'thinking'
  | 'responding'
  /** the model is streaming tool arguments (e.g. a full outline) with no visible text */
  | 'tool-input'
  | 'tool-running'

export interface AgentPhase {
  kind: AgentPhaseKind
  toolName?: string | undefined
}

// ---- LLM transport (how one model turn is streamed; app supplies the impl) ----

export interface AgentStreamRequest {
  system: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
}

export interface AgentStreamCallbacks {
  /** Vendor-managed turns request a tool and wait for its REAL result before
   * continuing the same stream. Must not also emit onToolCall for this call. */
  onToolRequest?(call: AgentToolCall): Promise<ToolExecution>
  onDelta(text: string): void
  /** raw model reasoning delta; the loop stores it on the assistant message for interleaved-thinking echo */
  onReasoning?(text: string): void
  /** complete parsed tool call (arguments finished streaming) */
  onToolCall(call: AgentToolCall): void
  /** Phase changes within the model stream (thinking / responding / tool-input); older transports may omit this */
  onPhase?(phase: AgentPhase): void
  /** normalized stop reason of the turn ('max_tokens' = cut off by the token limit); transports may omit this */
  onStopReason?(reason: string): void
  /** tokens the provider billed for this turn; transports whose provider does not report usage omit this */
  onUsage?(usage: AgentTokenUsage): void
  onDone(): void
  onError(error: string): void
}

export interface AgentStreamHandle {
  /** abort the in-flight turn; the transport must still emit onDone afterwards */
  cancel(): void
}

export interface AgentTransport {
  /** Optional host approval gate, checked immediately before skill execution. */
  approveTool?(
    call: AgentToolCall,
    signal?: AbortSignal,
    context?: { editor?: string; round?: { path: string; baseVersion: string } },
  ): Promise<boolean | 'allow-round'>
  stream(request: AgentStreamRequest, callbacks: AgentStreamCallbacks): AgentStreamHandle
}
