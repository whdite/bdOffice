import { withAttachedSkills } from '@genoffice/agent-core'
import type { SessionContextSelection } from '@genoffice/project-store'
import { ConversationContextPanel } from './ConversationContextPanel'
import { RunReviewCard } from './RunReviewCard'
import { AssistantTaskShelf } from './AssistantTaskShelf'
import { ConversationRecords } from './ConversationRecords'
import { WorkspaceToolSteps } from './WorkspaceToolSteps'
import {
  AiModelPicker,
  AiUsageChip,
  AiSkillsMenu,
  AiSkillChips,
  useDismissablePopover,
  useEscapeToClose,
  type AiModelPickerBridge,
} from '@genoffice/ui'
import type { AgentContextPreview } from '../../shared/workspace-agent-api'
import { ManagedSessionsPanel } from './ManagedSessionsPanel'
import { MemoryPanel } from './MemoryPanel'
import { SessionDrafts } from './session-drafts'
import {
  appendFileMentions,
  matchedFileMentions,
  mergeReferenceFiles,
  referencePathKey,
  removeFileMention,
  sessionReferenceFiles,
  type ChatReferenceFile,
} from './workspace-chat-files'
import { useCallback, useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent } from 'react'
import type {
  AgentApprovalMode,
  AgentSessionSummary,
  ToolPermissionDecision,
  WorkspaceAgentEvent,
  WorkspaceAgentContext,
  WorkspaceAgentReference,
  WorkspaceDescriptor,
  WorkspaceFileEntry,
  WorkspaceProjectOption,
} from '../../shared/workspace-agent-api'
import type { Lang } from '@genoffice/i18n'
import {
  emptyTimelineState,
  replayTimelineEvents,
  timelineForDisplay,
  reduceTimelineEvent,
  timelineStateFromHistory,
} from './workspace-chat-timeline'
import type { WorkspaceChatTimelineState } from './workspace-chat-timeline'
import { createChatEventBuffer } from './workspace-chat-event-buffer'
import chatIcon from './assets/chat-icon.svg'
import workspaceFolderIcon from './assets/workspace-folder.svg'
import { DogMascot } from './DogMascot'
import { DogIllustration } from './DogIllustration'
import {
  dogEventForAgentEvent,
  dogEventLabel,
  dogVisualForEvent,
  type DogEvent,
  type DogVisual,
} from './dog-assets'
import { useI18n } from './locale'
import {
  isChineseWorkspaceLanguage,
  workspaceChatCopy,
  workspaceErrorText,
  workspacePermissionText,
  workspaceSessionTitle,
  workspaceToolLabel,
} from './workspace-chat-strings'
interface WorkspaceChatSize {
  width: number
  height: number
}
import { WorkspaceMarkdown } from './WorkspaceMarkdown'
import { AgentApprovalControl } from './AgentApprovalControl'
import { openOnboarding } from './onboarding-copy'

declare global {
  interface Window {
    aiOfficeWorkspaceAgent?: import('../../shared/workspace-agent-api').WorkspaceAgentApi
    workspaceAgent?: import('../../shared/workspace-agent-api').WorkspaceAgentApi
  }
}

interface WorkspaceChatProps {
  onClose: () => void
  presentation?: 'docked' | 'window' | 'home'
  requestedSession?: { id: string; nonce: number }
  onOpenModelSettings?: () => void
  onPresentationChange?: () => void
  size?: WorkspaceChatSize
  onResize?: (size: WorkspaceChatSize) => void
}

type ResizeAxis = 'chat-width' | 'sidebar'

const DEFAULT_CHAT_SIZE: WorkspaceChatSize = { width: 560, height: 720 }
const SIDEBAR_SIZE_KEY = 'workspace-chat-sidebar-width'

function readStoredNumber(key: string, fallback: number, min: number, max: number): number {
  try {
    const raw = window.localStorage.getItem(key)
    if (raw === null) return fallback
    const value = Number(raw)
    if (Number.isFinite(value)) return Math.min(max, Math.max(min, Math.round(value)))
  } catch {
    // The panel remains usable when localStorage is unavailable.
  }
  return fallback
}

function persistNumber(key: string, value: number): void {
  try {
    window.localStorage.setItem(key, String(value))
  } catch {
    // Ignore storage failures; the current layout remains in memory.
  }
}

const FILE_ICON: Record<string, string> = {
  docx: 'DOC',
  pptx: 'PPT',
  xlsx: 'XLS',
  pdf: 'PDF',
  md: 'MD',
  txt: 'TXT',
  csv: 'CSV',
  json: 'JSON',
}

export function workspaceFileLabel(entry: ChatReferenceFile): string {
  if (entry.type === 'unknown') {
    const extension = /\.([^.]+)$/.exec(entry.name || fileName(entry.path))?.[1]
    return extension && /^[a-z0-9]{1,4}$/i.test(extension) ? extension.toUpperCase() : 'FILE'
  }
  return FILE_ICON[entry.type] ?? entry.type.toUpperCase()
}

function fileName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path
}

interface MentionCandidate {
  label: string
  detail: string
  kind: WorkspaceAgentReference['kind']
  path?: string
}

interface MentionRange {
  start: number
  end: number
  query: string
}

export function mentionRangeAtCursor(value: string, cursor: number): MentionRange | null {
  const before = value.slice(0, cursor)
  const match = /(?:^|\s)@([^\s@]*)$/u.exec(before)
  if (!match) return null
  const at = before.lastIndexOf('@')
  return at < 0 ? null : { start: at, end: cursor, query: match[1] ?? '' }
}

function mentionCandidates(
  files: ChatReferenceFile[],
  context: WorkspaceAgentContext | null,
  lang: Lang | string = 'zh',
): MentionCandidate[] {
  const active = context?.activeFile
  const editor = active?.editor
  const chinese = isChineseWorkspaceLanguage(lang)
  const labels = chinese
    ? {
        workspace: '项目',
        document: '当前文档',
        selection: '当前选区',
        page: '当前页',
        sheet: '当前工作表',
      }
    : {
        workspace: 'Workspace',
        document: 'Current document',
        selection: 'Current selection',
        page: 'Current page',
        sheet: 'Current sheet',
      }
  const names = new Map<string, number>()
  for (const file of files) names.set(file.name, (names.get(file.name) ?? 0) + 1)
  return [
    {
      label: labels.workspace,
      detail: context?.workspaceName ?? (chinese ? '工作区' : 'Workspace'),
      kind: 'workspace' as const,
    },
    ...(active
      ? [
          {
            label: labels.document,
            detail: active.path,
            kind: 'current-document' as const,
            path: active.path,
          },
        ]
      : []),
    ...(active && editor?.selection
      ? [
          {
            label: labels.selection,
            detail: active.path,
            kind: 'current-selection' as const,
            path: active.path,
          },
        ]
      : []),
    ...(active && (editor?.currentPage || editor?.currentSlide)
      ? [
          {
            label: labels.page,
            detail: `${active.path} · ${
              editor.currentSlide
                ? `${chinese ? '幻灯片' : 'Slide'} ${editor.currentSlide}`
                : `${chinese ? '第' : 'Page'} ${editor.currentPage}${chinese ? ' 页' : ''}`
            }`,
            kind: 'current-page' as const,
            path: active.path,
          },
        ]
      : []),
    ...(active && editor?.currentSheet
      ? [
          {
            label: labels.sheet,
            detail: `${editor.currentSheet}${editor.currentRange ? `!${editor.currentRange}` : ''}`,
            kind: 'current-sheet-range' as const,
            path: active.path,
          },
        ]
      : []),
    ...files.map((file) => ({
      label:
        (names.get(file.name) ?? 0) > 1 || Object.values(labels).includes(file.name)
          ? file.path
          : file.name,
      detail: file.path,
      kind: 'file' as const,
      path: file.path,
    })),
  ]
}

export function buildAgentReferences(
  text: string,
  files: ChatReferenceFile[],
  context: WorkspaceAgentContext | null,
  lang: Lang | string = 'zh',
): WorkspaceAgentReference[] {
  const active = context?.activeFile
  const editor = active?.editor
  const chinese = isChineseWorkspaceLanguage(lang)
  const labels = chinese
    ? {
        workspace: '项目',
        document: '当前文档',
        selection: '当前选区',
        page: '当前页',
        sheet: '当前工作表',
      }
    : {
        workspace: 'Workspace',
        document: 'Current document',
        selection: 'Current selection',
        page: 'Current page',
        sheet: 'Current sheet',
      }
  const result: WorkspaceAgentReference[] = []
  const candidates = mentionCandidates(files, context, lang)
  const matched = matchedFileMentions(
    text,
    candidates.flatMap((candidate) =>
      candidate.kind === 'file' && candidate.path
        ? [candidate.label, candidate.path]
        : [candidate.label],
    ),
  )
  const mentioned = (label: string) => matched.has(label)
  if (mentioned(labels.workspace)) result.push({ kind: 'workspace', label: `@${labels.workspace}` })
  if (active && mentioned(labels.document))
    result.push({
      kind: 'current-document',
      label: `@${labels.document}`,
      path: active.path,
      type: active.type,
    })
  if (active && editor?.selection && mentioned(labels.selection))
    result.push({
      kind: 'current-selection',
      label: `@${labels.selection}`,
      path: active.path,
      type: active.type,
      selection: editor.selection,
    })
  if (active && mentioned(labels.page) && (editor?.currentPage || editor?.currentSlide))
    result.push({
      kind: 'current-page',
      label: `@${labels.page}`,
      path: active.path,
      type: active.type,
      ...(editor.currentPage ? { page: editor.currentPage } : {}),
      ...(editor.currentSlide ? { slide: editor.currentSlide } : {}),
    })
  if (active && editor?.currentSheet && mentioned(labels.sheet))
    result.push({
      kind: 'current-sheet-range',
      label: `@${labels.sheet}`,
      path: active.path,
      type: active.type,
      sheet: editor.currentSheet,
      ...(editor.currentRange ? { range: editor.currentRange } : {}),
    })
  for (const candidate of candidates.filter((candidate) => candidate.kind === 'file')) {
    if (!candidate.path || (!mentioned(candidate.label) && !mentioned(candidate.path))) continue
    const file = files.find((entry) => entry.path === candidate.path)
    if (file)
      result.push({
        kind: 'file',
        label: `@${mentioned(candidate.label) ? candidate.label : candidate.path}`,
        path: file.path,
        type: file.type,
      })
  }
  return result.slice(0, 24)
}

interface FileTreeNode {
  path: string
  name: string
  directory: boolean
  entry?: WorkspaceFileEntry
  children: FileTreeNode[]
  lazy?: boolean
}

function buildFileTree(entries: WorkspaceFileEntry[]): FileTreeNode[] {
  const root: FileTreeNode = { path: '', name: '', directory: true, children: [] }
  const byPath = new Map<string, FileTreeNode>([['', root]])
  const ensureDirectory = (path: string): FileTreeNode => {
    const normalized = path.replace(/\\/g, '/').replace(/^\.\//, '').replace(/\/$/, '')
    if (!normalized) return root
    const existing = byPath.get(normalized)
    if (existing) return existing
    const parentPath = normalized.slice(0, Math.max(0, normalized.lastIndexOf('/')))
    const parent = ensureDirectory(parentPath)
    const node: FileTreeNode = {
      path: normalized,
      name: fileName(normalized),
      directory: true,
      children: [],
    }
    parent.children.push(node)
    byPath.set(normalized, node)
    return node
  }
  for (const entry of entries) {
    const path = entry.path.replace(/\\/g, '/')
    const parentPath = path.slice(0, Math.max(0, path.lastIndexOf('/')))
    const parent = entry.directory ? ensureDirectory(path) : ensureDirectory(parentPath)
    if (entry.directory) {
      const node = byPath.get(path) ?? parent
      node.entry = entry
      continue
    }
    const node: FileTreeNode = {
      path,
      name: entry.name || fileName(path),
      directory: false,
      entry,
      children: [],
    }
    parent.children.push(node)
    byPath.set(path, node)
  }
  const sort = (nodes: FileTreeNode[]) => {
    nodes.sort((a, b) => {
      if (a.directory !== b.directory) return a.directory ? -1 : 1
      return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' })
    })
    nodes.forEach((node) => sort(node.children))
  }
  sort(root.children)
  return root.children
}

function FileTreeBranch({
  node,
  onOpen,
  depth = 0,
}: {
  node: FileTreeNode
  onOpen(path: string): void
  depth?: number
}) {
  const [loaded, setLoaded] = useState<WorkspaceFileEntry[] | null>(null)
  const [loadError, setLoadError] = useState('')
  const loading = useRef(false)
  if (node.directory)
    return (
      <details
        className="workspace-chat-tree-dir"
        open={!node.lazy && depth === 0}
        onToggle={(event) => {
          if (!node.lazy || !event.currentTarget.open || loaded || loading.current) return
          loading.current = true
          void workspaceAgentApi()
            .listFiles({ path: node.path, recursive: false })
            .then(setLoaded)
            .catch((cause) => setLoadError(String(cause)))
            .finally(() => {
              loading.current = false
            })
        }}
      >
        <summary title={node.path}>
          <span className="workspace-chat-tree-caret" aria-hidden="true" />
          <span className="workspace-chat-tree-dir-name">{node.name}</span>
        </summary>
        <div className="workspace-chat-tree-children">
          {loadError && (
            <div role="alert" className="workspace-chat-muted">
              {loadError}
            </div>
          )}
          {loaded?.map((entry) => (
            <FileTreeBranch
              key={entry.path}
              node={{
                path: entry.path,
                name: entry.name,
                directory: entry.directory,
                entry,
                children: [],
                lazy: entry.directory,
              }}
              onOpen={onOpen}
              depth={depth + 1}
            />
          ))}
          {node.children.map((child) => (
            <FileTreeBranch key={child.path} node={child} onOpen={onOpen} depth={depth + 1} />
          ))}
        </div>
      </details>
    )
  const entry = node.entry
  if (!entry) return null
  return (
    <button
      className="workspace-chat-file"
      style={{ paddingLeft: `${8 + depth * 14}px` }}
      onClick={() => onOpen(entry.path)}
      title={entry.path}
    >
      <span className={`workspace-chat-file-type ${entry.type}`}>{workspaceFileLabel(entry)}</span>
      <span className="workspace-chat-file-name">{entry.name || fileName(entry.path)}</span>
    </button>
  )
}

function workspaceAgentApi(): import('../../shared/workspace-agent-api').WorkspaceAgentApi {
  const api = window.workspaceAgent ?? window.aiOfficeWorkspaceAgent
  if (!api) throw new Error('Workspace Agent is unavailable.')
  return api
}

function dogVisualForTimeline(timeline: WorkspaceChatTimelineState): DogVisual {
  const latest = timeline.items[timeline.items.length - 1]
  if (latest?.kind === 'meme') return { mood: latest.mood, frame: latest.frame }
  if (latest?.kind === 'run')
    return dogVisualForEvent(
      latest.run.state === 'undo-failed'
        ? 'error'
        : latest.run.state === 'undone'
          ? 'cancelled'
          : 'done',
    )
  if (latest?.kind === 'tool' && latest.state === 'failed') return dogVisualForEvent('error')
  if (latest?.kind === 'file') {
    return dogVisualForEvent(
      latest.operation === 'read' ? 'read' : latest.operation === 'created' ? 'created' : 'write',
    )
  }
  if (latest?.kind === 'permission' && latest.state === 'pending')
    return dogVisualForEvent('approval')
  return dogVisualForEvent('idle')
}

export function WorkspaceChat({
  onClose,
  presentation = 'home',
  requestedSession,
  onOpenModelSettings,
  onPresentationChange,
  size = DEFAULT_CHAT_SIZE,
  onResize,
}: WorkspaceChatProps) {
  const { lang } = useI18n()
  const copy = useMemo(() => workspaceChatCopy(lang), [lang])
  const displayError = useCallback(
    (cause: unknown, fallback = copy.agentFailed): string =>
      workspaceErrorText(cause instanceof Error ? cause.message : String(cause), lang) || fallback,
    [copy.agentFailed, lang],
  )
  const [workspace, setWorkspace] = useState<WorkspaceDescriptor | null>(null)
  const [files, setFiles] = useState<WorkspaceFileEntry[]>([])
  const [sessions, setSessions] = useState<AgentSessionSummary[]>([])
  const [sessionManagerOpen, setSessionManagerOpen] = useState(false)
  const [advancedSessions, setAdvancedSessions] = useState(false)
  const [contextOpen, setContextOpen] = useState(false)
  const [contextPreview, setContextPreview] = useState<AgentContextPreview | null>(null)
  const [contextLoadingError, setContextLoadingError] = useState<string | null>(null)
  const [contextSaving, setContextSaving] = useState(false)
  const contextPanelId = useId()
  const contextTriggerRef = useRef<HTMLButtonElement>(null)
  const contextLayerRef = useRef<HTMLDivElement>(null)
  const memoryOpenerRef = useRef<HTMLButtonElement | null>(null)
  const zh = isChineseWorkspaceLanguage(lang)
  const modelBridge = useMemo<AiModelPickerBridge>(
    () => ({
      getSettings: () => window.aiOffice.aiGetSettings(),
      setSettings: (settings) => window.aiOffice.setAiSettings(settings),
      onSettingsChanged: (handler) => window.aiOffice.onAiSettingsChanged(handler),
      gskLoggedIn: async () => false,
      openModelSettings: () => onOpenModelSettings?.(),
    }),
    [onOpenModelSettings],
  )
  const [activationEpoch, setActivationEpoch] = useState(0)
  const [activeSessionId, setActiveSessionId] = useState<string | null>(null)
  const activeSessionRef = useRef(activeSessionId)
  activeSessionRef.current = activeSessionId
  const [attachingFiles, setAttachingFiles] = useState(false)
  const contextMutationLock = useRef(false)
  const referenceLoadVersion = useRef(0)
  const [contextSelection, setContextSelection] = useState<{
    sessionId: string
    value: SessionContextSelection
  } | null>(null)
  const selectedContext =
    contextSelection?.sessionId === activeSessionId ? contextSelection.value : null
  const referenceFiles = useMemo(
    () => sessionReferenceFiles(selectedContext?.filePaths ?? []),
    [selectedContext?.filePaths],
  )
  const closeContext = () => {
    setContextOpen(false)
    contextTriggerRef.current?.focus({ preventScroll: true })
  }
  useEscapeToClose(contextOpen, closeContext)
  useDismissablePopover(contextOpen, () => setContextOpen(false), {
    inside: () => [contextTriggerRef.current, contextLayerRef.current],
  })
  const mentionFiles = useMemo(
    () => mergeReferenceFiles(files, referenceFiles, workspace?.root),
    [files, referenceFiles, workspace?.root],
  )
  const fileTree = useMemo(
    () =>
      buildFileTree(
        files.filter(
          (file) =>
            !file.registered &&
            !referenceFiles.some(
              (reference) =>
                referencePathKey(reference.path, workspace?.root) ===
                referencePathKey(file.path, workspace?.root),
            ),
        ),
      ),
    [files, referenceFiles, workspace?.root],
  )
  const [dragOver, setDragOver] = useState(false)
  const filePickerRef = useRef<HTMLInputElement>(null)
  const [timeline, setTimeline] = useState(emptyTimelineState)
  const [drafts] = useState(() => new SessionDrafts(window.localStorage))
  useEffect(() => {
    const api = window.managedSessions
    if (!api) return
    let sequence = 0
    const refresh = () => {
      const current = ++sequence
      void api
        .list({ state: 'all' })
        .then((all) => {
          if (current === sequence) drafts.retain(all.map((item) => item.id))
        })
        .catch(() => {})
    }
    refresh()
    const off = api.onChanged(refresh)
    return () => {
      sequence++
      off()
    }
  }, [drafts])
  const [, refreshDraft] = useState(0)
  const input = drafts.read(activeSessionId)
  const setInput = useCallback(
    (value: string) => {
      drafts.write(activeSessionId, value)
      refreshDraft((revision) => revision + 1)
    },
    [drafts, activeSessionId],
  )
  const [attachedSkills, setAttachedSkills] = useState<string[]>([])
  const sendEpoch = useRef(0)
  const [agentContext, setAgentContext] = useState<WorkspaceAgentContext | null>(null)
  const [mentionRange, setMentionRange] = useState<MentionRange | null>(null)
  const [mentionIndex, setMentionIndex] = useState(0)
  const [running, setRunning] = useState(false)
  const [approvalMode, setApprovalMode] = useState<AgentApprovalMode>('per-action')
  const [approvalModeSaving, setApprovalModeSaving] = useState(false)
  const [approvalModeReady, setApprovalModeReady] = useState(false)
  const [approvalLoadError, setApprovalLoadError] = useState('')
  const [approvalLoadAttempt, setApprovalLoadAttempt] = useState(0)
  const [sessionAction, setSessionAction] = useState<{
    kind: 'rename' | 'delete'
    session: AgentSessionSummary
  } | null>(null)
  const [sessionName, setSessionName] = useState('')
  const [sessionSaving, setSessionSaving] = useState(false)
  const [sessionError, setSessionError] = useState('')
  const sessionSaveLock = useRef(false)
  const approvalSaveLock = useRef(false)
  const approvalModeRevision = useRef(0)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [undoingRunId, setUndoingRunId] = useState<string | null>(null)
  const [workspacePickerOpen, setWorkspacePickerOpen] = useState(false)
  const [workspacePickerMode, setWorkspacePickerMode] = useState<'project' | 'folder'>('project')
  const [projectOptions, setProjectOptions] = useState<WorkspaceProjectOption[]>([])
  const [projectOptionsLoading, setProjectOptionsLoading] = useState(false)
  const [sidebarWidth, setSidebarWidth] = useState(() =>
    readStoredNumber(SIDEBAR_SIZE_KEY, 220, 180, 420),
  )
  // A standalone Agent is the full workspace, not a chat-only popout. Keep
  // each presentation's explicit navigation choice across native reparenting.
  const [sidebarVisibility, setSidebarVisibility] = useState({
    docked: false,
    window: true,
    home: false,
  })
  const sidebarOpen = sidebarVisibility[presentation]
  const bootstrapped = useRef(false)
  const stopResize = useRef<() => void>(() => {})
  const scrollRef = useRef<HTMLDivElement>(null)
  const followStream = useRef(true)
  const fileRefreshTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const fileRefreshVersion = useRef(0)
  const workspaceLoadVersion = useRef(0)
  const inputRef = useRef<HTMLTextAreaElement>(null)
  const panelRef = useRef<HTMLElement>(null)
  const mainRef = useRef<HTMLElement>(null)
  const sidebarRef = useRef<HTMLElement>(null)
  const composerRef = useRef<HTMLFormElement>(null)
  const mentionListRef = useRef<HTMLDivElement>(null)
  const mentionListId = useId()
  const [layout, setLayout] = useState({
    panelWidth: size.width,
    width: size.width,
    height: size.height,
  })
  const [mentionPlacement, setMentionPlacement] = useState({ side: 'above', height: 280 })
  const initialDogVisual = dogVisualForEvent('idle')
  const [dogVisual, setDogVisual] = useState<DogVisual>(initialDogVisual)
  const dogVisualRef = useRef<DogVisual>(initialDogVisual)
  const dogReceivedAt = useRef(0)
  const pendingDogVisual = useRef<DogVisual | null>(null)
  const dogTransitionTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  const resizeComposer = useCallback(() => {
    const node = inputRef.current
    if (!node) return
    node.style.height = 'auto'
    const maxHeight = Math.min(190, Math.max(50, Math.floor(layout.height * 0.24)))
    const nextHeight = Math.min(maxHeight, Math.max(50, node.scrollHeight))
    node.style.height = `${nextHeight}px`
    node.style.overflowY = node.scrollHeight > maxHeight ? 'auto' : 'hidden'
  }, [layout.height])

  useLayoutEffect(() => {
    const main = mainRef.current,
      panel = panelRef.current
    if (!main || !panel) return
    const measure = () => {
      if (!main.clientWidth || !main.clientHeight) return
      const next = {
        panelWidth: panel.clientWidth,
        width: main.clientWidth,
        height: main.clientHeight,
      }
      setLayout((current) =>
        current.panelWidth === next.panelWidth &&
        current.width === next.width &&
        current.height === next.height
          ? current
          : next,
      )
    }
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(main)
    observer.observe(panel)
    return () => observer.disconnect()
  }, [loading, workspace?.id])

  useLayoutEffect(() => {
    resizeComposer()
  }, [input, resizeComposer, layout.width])

  const resizePanel = useCallback(
    (axis: ResizeAxis, event: ReactPointerEvent<HTMLElement>) => {
      if (!onResize && axis === 'chat-width') return
      event.preventDefault()
      stopResize.current()
      const target = event.currentTarget
      target.setPointerCapture?.(event.pointerId)
      // The native view's origin moves while its left edge is dragged.
      const startX = event.screenX
      const startSize = { ...size }
      const startSidebar = sidebarRef.current?.clientWidth ?? sidebarWidth
      const panel = panelRef.current
      const scale = panel?.clientWidth ? panel.getBoundingClientRect().width / panel.clientWidth : 1
      const move = (nextEvent: PointerEvent): void => {
        const dx = nextEvent.screenX - startX
        if (axis === 'chat-width') {
          onResize?.({ width: startSize.width - dx, height: startSize.height })
        }
        if (axis === 'sidebar') {
          const max = Math.max(
            180,
            Math.min(420, Math.floor(layout.panelWidth * 0.38), layout.panelWidth - 286),
          )
          const next = Math.min(max, Math.max(180, Math.round(startSidebar + dx / scale)))
          setSidebarWidth(next)
          persistNumber(SIDEBAR_SIZE_KEY, next)
        }
      }
      const stop = (): void => {
        window.removeEventListener('pointermove', move)
        window.removeEventListener('pointerup', stop)
        window.removeEventListener('pointercancel', stop)
        window.removeEventListener('blur', stop)
        target.removeEventListener('lostpointercapture', stop)
        if (target.hasPointerCapture?.(event.pointerId))
          target.releasePointerCapture(event.pointerId)
        document.body.classList.remove('workspace-chat-resizing')
      }
      document.body.classList.add('workspace-chat-resizing')
      window.addEventListener('pointermove', move)
      window.addEventListener('pointerup', stop, { once: true })
      window.addEventListener('pointercancel', stop, { once: true })
      window.addEventListener('blur', stop, { once: true })
      target.addEventListener('lostpointercapture', stop, { once: true })
      stopResize.current = stop
    },
    [onResize, sidebarWidth, size, layout.panelWidth],
  )

  useEffect(() => () => stopResize.current(), [])

  const applyDogVisual = useCallback((next: DogVisual) => {
    if (dogVisualRef.current.frame === next.frame && dogVisualRef.current.mood === next.mood) return
    dogVisualRef.current = next
    setDogVisual(next)
  }, [])

  const setDogEvent = useCallback(
    (event: DogEvent) => {
      const next = dogVisualForEvent(event)
      if (event === 'received') {
        if (dogTransitionTimer.current) clearTimeout(dogTransitionTimer.current)
        dogTransitionTimer.current = null
        pendingDogVisual.current = null
        dogReceivedAt.current = Date.now()
        applyDogVisual(next)
        return
      }

      // Keep the acknowledgement frame on screen long enough to be perceived
      // when the runtime emits `assistant.message` and `agent.status` back to
      // back. This is a bounded transition, never an automatic carousel.
      const elapsed = Date.now() - dogReceivedAt.current
      if (dogVisualRef.current.mood === 'received' && elapsed < 360) {
        pendingDogVisual.current = next
        if (!dogTransitionTimer.current) {
          dogTransitionTimer.current = setTimeout(() => {
            dogTransitionTimer.current = null
            const pending = pendingDogVisual.current
            pendingDogVisual.current = null
            if (pending) applyDogVisual(pending)
          }, 360 - elapsed)
        }
        return
      }
      pendingDogVisual.current = null
      applyDogVisual(next)
    },
    [applyDogVisual],
  )

  const resetDogVisual = useCallback(() => {
    if (dogTransitionTimer.current) clearTimeout(dogTransitionTimer.current)
    dogTransitionTimer.current = null
    pendingDogVisual.current = null
    dogReceivedAt.current = 0
    applyDogVisual(dogVisualForEvent('idle'))
  }, [applyDogVisual])

  const applyDogEvent = useCallback(
    (event: WorkspaceAgentEvent) => {
      const visualEvent = dogEventForAgentEvent(event, dogVisualRef.current)
      if (visualEvent) setDogEvent(visualEvent)
    },
    [setDogEvent],
  )

  const refreshAgentContext = useCallback(async () => {
    try {
      const next = await workspaceAgentApi().getContext()
      setAgentContext(next)
      return next
    } catch {
      setAgentContext(null)
      return null
    }
  }, [])
  useEffect(() => {
    void refreshAgentContext()
  }, [refreshAgentContext, requestedSession?.nonce])

  const visibleMentions = useMemo(() => {
    if (!mentionRange) return []
    const query = mentionRange.query.toLocaleLowerCase()
    return mentionCandidates(mentionFiles, agentContext, lang)
      .filter(
        (candidate) =>
          !query ||
          candidate.label.toLocaleLowerCase().includes(query) ||
          candidate.detail.toLocaleLowerCase().includes(query),
      )
      .slice(0, 12)
  }, [agentContext, mentionFiles, lang, mentionRange])
  const mentionsOpen = !!mentionRange && visibleMentions.length > 0
  useDismissablePopover(mentionsOpen, () => setMentionRange(null), {
    inside: () => [inputRef.current, mentionListRef.current],
    escape: true,
  })
  useLayoutEffect(() => {
    if (!mentionsOpen) return
    const input = inputRef.current,
      main = mainRef.current,
      composer = composerRef.current
    if (!input || !main || !composer) return
    const place = () => {
      const anchor = input.getBoundingClientRect(),
        bounds = main.getBoundingClientRect()
      const scale = main.clientWidth ? bounds.width / main.clientWidth : 1
      if (!scale) return
      const above = Math.max(0, (anchor.top - bounds.top) / scale - 8)
      const below = Math.max(0, (bounds.bottom - anchor.bottom) / scale - 8)
      const side = above >= below ? 'above' : 'below'
      const height = Math.floor(Math.min(300, side === 'above' ? above : below))
      setMentionPlacement((current) =>
        current.side === side && current.height === height ? current : { side, height },
      )
    }
    place()
    const observer = new ResizeObserver(place)
    observer.observe(input)
    observer.observe(main)
    observer.observe(composer)
    window.addEventListener('scroll', place, true)
    window.addEventListener('resize', place)
    return () => {
      observer.disconnect()
      window.removeEventListener('scroll', place, true)
      window.removeEventListener('resize', place)
    }
  }, [mentionsOpen, input, layout.width, layout.height])
  useLayoutEffect(() => {
    const list = mentionListRef.current
    const row = list?.children[Math.min(mentionIndex, visibleMentions.length - 1)] as
      HTMLElement | undefined
    if (!list || !row) return
    if (row.offsetTop < list.scrollTop) list.scrollTop = row.offsetTop
    else if (row.offsetTop + row.offsetHeight > list.scrollTop + list.clientHeight)
      list.scrollTop = row.offsetTop + row.offsetHeight - list.clientHeight
  }, [mentionIndex, visibleMentions, mentionPlacement.height])

  const refreshFiles = useCallback(
    async (showError = true) => {
      const version = ++fileRefreshVersion.current
      try {
        const nextFiles = await workspaceAgentApi().listFiles({ recursive: true })
        if (version === fileRefreshVersion.current) setFiles(nextFiles)
      } catch (cause) {
        if (showError && version === fileRefreshVersion.current) setError(displayError(cause))
      }
    },
    [displayError],
  )

  const scheduleFilesRefresh = useCallback(() => {
    if (fileRefreshTimer.current) clearTimeout(fileRefreshTimer.current)
    fileRefreshTimer.current = setTimeout(() => {
      fileRefreshTimer.current = null
      void refreshFiles(false)
    }, 180)
  }, [refreshFiles])

  const loadWorkspace = useCallback(
    async (descriptor: WorkspaceDescriptor | null) => {
      const loadVersion = ++workspaceLoadVersion.current
      const fileVersion = ++fileRefreshVersion.current
      setWorkspace(descriptor)
      if (!descriptor) {
        setFiles([])
        setAgentContext(null)
        setSessions([])
        setActiveSessionId(null)
        setTimeline(emptyTimelineState())
        return
      }
      const api = workspaceAgentApi()
      // File discovery can be slow on large or remote save folders. It must not
      // keep a new conversation or connection setup waiting for the file tree.
      void api
        .listFiles({ recursive: true })
        .then((nextFiles) => {
          if (
            loadVersion === workspaceLoadVersion.current &&
            fileVersion === fileRefreshVersion.current
          )
            setFiles(nextFiles)
        })
        .catch(() => {
          if (
            loadVersion === workspaceLoadVersion.current &&
            fileVersion === fileRefreshVersion.current
          )
            setFiles([])
        })
      const [nextSessions, nextContext] = await Promise.all([
        api.listSessions(),
        api.getContext().catch(() => null),
      ])
      if (loadVersion !== workspaceLoadVersion.current) return
      setSessions(nextSessions)
      setAgentContext(nextContext)
      setActiveSessionId((current) =>
        drafts.restoreSelection(
          current,
          nextSessions.map((session) => session.id),
        ),
      )
    },
    [drafts],
  )

  const refresh = useCallback(async () => {
    setLoading(true)
    setError(null)
    try {
      await loadWorkspace(await workspaceAgentApi().getWorkspace())
    } catch (cause) {
      setError(displayError(cause))
    } finally {
      setLoading(false)
    }
  }, [displayError, loadWorkspace])

  useEffect(() => {
    if (!requestedSession?.id) return
    sendEpoch.current++
    setAttachedSkills([])
    let alive = true
    const api = workspaceAgentApi()
    void api
      .resumeSession(requestedSession.id)
      .then(async (session) => {
        const ws = await api.getWorkspace()
        if (!alive) return
        await loadWorkspace(ws)
        if (alive) {
          setActiveSessionId(session.id)
          setActivationEpoch((value) => value + 1)
        }
      })
      .catch((cause) => {
        if (alive) setError(displayError(cause))
      })
    return () => {
      alive = false
    }
  }, [requestedSession, loadWorkspace, displayError])
  useEffect(() => {
    let alive = true
    const version = ++referenceLoadVersion.current
    setContextPreview(null)
    setContextLoadingError(null)
    if (activeSessionId)
      void workspaceAgentApi()
        .getContextSelection(activeSessionId)
        .then((selection) => {
          if (alive && version === referenceLoadVersion.current)
            setContextSelection({ sessionId: activeSessionId, value: selection })
        })
        .catch((cause) => {
          if (alive && version === referenceLoadVersion.current)
            setContextLoadingError(displayError(cause))
        })
    if (contextOpen && activeSessionId && workspaceAgentApi().previewContext)
      void workspaceAgentApi()
        .previewContext(activeSessionId)
        .then((p) => {
          if (alive && version === referenceLoadVersion.current) setContextPreview(p)
        })
        .catch(() => {})
    return () => {
      alive = false
    }
  }, [activeSessionId, contextOpen, sessionManagerOpen, displayError])
  const ensureCurrentSession = async () => {
    if (activeSessionRef.current) return activeSessionRef.current
    const epoch = sendEpoch.current
    const session = await workspaceAgentApi().createSession(workspace?.id)
    if (epoch !== sendEpoch.current) throw new Error('The selected conversation changed.')
    activeSessionRef.current = session.id
    drafts.adoptPending(session.id)
    setSessions((current) => [session, ...current.filter((s) => s.id !== session.id)])
    setActiveSessionId(session.id)
    return session.id
  }
  const loadSessionMemory = async (id: string) => {
    const current = await ensureCurrentSession(),
      api = workspaceAgentApi()
    if (current === id) return
    const selected = await api.getContextSelection(current)
    const next = {
      ...selected,
      sessionIds: [...new Set([...selected.sessionIds, id])],
      excludedSessionIds: selected.excludedSessionIds?.filter((source) => source !== id),
    }
    await api.setContextSelection(current, next)
    if (activeSessionRef.current === current)
      setContextSelection({ sessionId: current, value: next })
  }
  const showContext = async () => {
    if (contextOpen) {
      setContextOpen(false)
      return
    }
    try {
      await ensureCurrentSession()
      setMentionRange(null)
      setContextOpen(true)
    } catch (cause) {
      setError(displayError(cause))
    }
  }
  const openMemoryManager = (opener: HTMLButtonElement | null) => {
    memoryOpenerRef.current = opener
    setContextOpen(false)
    setAdvancedSessions(false)
    setSessionManagerOpen(true)
  }
  const closeMemoryManager = () => {
    setSessionManagerOpen(false)
    requestAnimationFrame(() => memoryOpenerRef.current?.focus({ preventScroll: true }))
  }
  const attachFiles = async (files: File[]) => {
    if (running || contextMutationLock.current || !files.length) return
    const epoch = sendEpoch.current
    contextMutationLock.current = true
    setAttachingFiles(true)
    try {
      setError(null)
      if (files.length > 16)
        throw new Error(
          isChineseWorkspaceLanguage(lang)
            ? '一次最多添加 16 个参考文件。'
            : 'Add up to 16 reference files at a time.',
        )
      const paths: string[] = []
      for (const file of files) {
        const path = window.aiOffice.pathForFile(file)
        if (path) paths.push(path)
        else {
          if (!/^image\/(png|jpeg|webp)$/i.test(file.type))
            throw new Error(
              isChineseWorkspaceLanguage(lang)
                ? '请从本机选择或拖入文件。'
                : 'Choose or drag a local file.',
            )
          const result = await window.aiOffice.addPastedImage(
            await file.arrayBuffer(),
            file.type.split('/')[1]?.replace('jpeg', 'jpg') ?? '',
          )
          paths.push(...result.accepted.map((item) => item.path))
        }
      }
      if (!paths.length || epoch !== sendEpoch.current) return
      const id = await ensureCurrentSession(),
        api = workspaceAgentApi()
      const previous = await api.getContextSelection(id)
      if (epoch !== sendEpoch.current || activeSessionRef.current !== id) return
      const selection = await api.attachContextFiles(id, paths)
      const nextFiles = mergeReferenceFiles(
        mergeReferenceFiles(mentionFiles, selection.addedPaths ?? [], workspace?.root),
        sessionReferenceFiles(selection.filePaths),
        workspace?.root,
      )
      const attached =
        selection.addedPaths?.map((entry) => entry.path) ??
        selection.filePaths.filter(
          (path) =>
            !previous.filePaths.includes(path) ||
            paths.some((source) => referencePathKey(source) === referencePathKey(path)),
        )
      drafts.write(
        id,
        appendFileMentions(
          drafts.read(id),
          attached,
          mentionCandidates(mentionFiles, agentContext, lang),
          mentionCandidates(nextFiles, agentContext, lang),
          workspace?.root,
        ),
      )
      if (epoch === sendEpoch.current && activeSessionRef.current === id) {
        ++referenceLoadVersion.current
        refreshDraft((revision) => revision + 1)
        setContextSelection({ sessionId: id, value: selection })
        setContextLoadingError(null)
        if (selection.addedPaths) {
          const keys = new Set(
            selection.addedPaths.map((entry) => referencePathKey(entry.path, workspace?.root)),
          )
          setFiles((current) => [
            ...current.filter((entry) => !keys.has(referencePathKey(entry.path, workspace?.root))),
            ...selection.addedPaths!,
          ])
        }
        setContextOpen(false)
        setMentionRange(null)
        requestAnimationFrame(() => {
          if (activeSessionRef.current !== id) return
          inputRef.current?.focus()
          const length = inputRef.current?.value.length ?? 0
          inputRef.current?.setSelectionRange(length, length)
        })
      }
    } catch (cause) {
      if (epoch === sendEpoch.current) setError(displayError(cause))
    } finally {
      contextMutationLock.current = false
      setAttachingFiles(false)
    }
  }

  const updateContextSelection = async (
    update: (selection: SessionContextSelection) => SessionContextSelection,
    removedFile?: string,
  ) => {
    const id = activeSessionId,
      epoch = sendEpoch.current
    if (!id || running || contextMutationLock.current) return
    contextMutationLock.current = true
    setContextSaving(true)
    try {
      const api = workspaceAgentApi(),
        selection = await api.getContextSelection(id)
      if (activeSessionRef.current !== id || epoch !== sendEpoch.current) return
      const next = update(selection)
      await api.setContextSelection(id, next)
      if (removedFile)
        drafts.write(
          id,
          removeFileMention(
            drafts.read(id),
            removedFile,
            mentionCandidates(mentionFiles, agentContext, lang),
          ),
        )
      if (activeSessionRef.current !== id || epoch !== sendEpoch.current) return
      const version = ++referenceLoadVersion.current
      setContextSelection({ sessionId: id, value: next })
      if (removedFile) refreshDraft((revision) => revision + 1)
      setContextLoadingError(null)
      setError(null)
      if (contextOpen)
        void api
          .previewContext(id)
          .then((preview) => {
            if (
              activeSessionRef.current === id &&
              epoch === sendEpoch.current &&
              version === referenceLoadVersion.current
            )
              setContextPreview(preview)
          })
          .catch(() => {})
    } catch (cause) {
      if (activeSessionRef.current === id && epoch === sendEpoch.current)
        throw new Error(displayError(cause), { cause })
    } finally {
      contextMutationLock.current = false
      setContextSaving(false)
    }
  }

  const removeReference = (path: string) =>
    updateContextSelection(
      (selection) => ({ ...selection, filePaths: selection.filePaths.filter((file) => file !== path) }),
      path,
    )

  const applyEventSideEffects = useCallback(
    (event: WorkspaceAgentEvent) => {
      applyDogEvent(event)
      if (
        event.type === 'document.created' ||
        event.type === 'file.modified' ||
        event.tool?.mutated
      )
        scheduleFilesRefresh()
      if (event.type === 'agent.status') {
        setRunning(event.status === 'running')
        if (event.status === 'running' || event.status === 'completed') setError(null)
      }
      if ((event.type === 'agent.status' && event.status !== 'running') || event.type === 'error') {
        void refreshAgentContext()
        const version = workspaceLoadVersion.current
        void workspaceAgentApi()
          .listSessions()
          .then((next) => {
            if (version === workspaceLoadVersion.current) setSessions(next)
          })
          .catch(() => {})
      }
      if (event.type === 'error') {
        setRunning(false)
        setError(displayError(event.error ?? event.text ?? copy.agentFailed))
      }
    },
    [
      applyDogEvent,
      copy.agentFailed,
      displayError,
      scheduleFilesRefresh,
      refreshAgentContext,
      workspace?.id,
    ],
  )

  const handleEvents = useCallback(
    (events: WorkspaceAgentEvent[]) => {
      setTimeline((current) =>
        events.reduce(
          (next, event) => reduceTimelineEvent(next, event, { includeDogMemes: true }),
          current,
        ),
      )
      for (const event of events) applyEventSideEffects(event)
    },
    [applyEventSideEffects],
  )

  useEffect(() => {
    if (bootstrapped.current) return
    bootstrapped.current = true
    void refresh()
  }, [refresh])

  useEffect(() => {
    let alive = true
    let timer: ReturnType<typeof setTimeout> | undefined
    let sequence = 0
    const refreshCatalog = () => {
      clearTimeout(timer)
      timer = setTimeout(() => {
        const current = ++sequence
        const version = workspaceLoadVersion.current
        void workspaceAgentApi()
          .listSessions()
          .then((next) => {
            if (alive && current === sequence && version === workspaceLoadVersion.current)
              setSessions(next)
          })
          .catch(() => {})
      }, 150)
    }
    const off = window.managedSessions?.onChanged(refreshCatalog)
    window.addEventListener('focus', refreshCatalog)
    refreshCatalog()
    return () => {
      alive = false
      clearTimeout(timer)
      off?.()
      window.removeEventListener('focus', refreshCatalog)
    }
  }, [workspace?.id])

  useEffect(() => {
    let alive = true
    let changed = false
    const getApprovalMode = workspaceAgentApi().getApprovalMode
    setApprovalModeReady(false)
    setApprovalLoadError('')
    if (typeof getApprovalMode !== 'function') {
      setApprovalLoadError(copy.agentFailed)
      return () => {}
    }
    const unsubscribe = workspaceAgentApi().onApprovalModeChanged?.((mode) => {
      if (!alive) return
      changed = true
      approvalModeRevision.current++
      setApprovalMode(mode)
      setApprovalModeReady(true)
      setApprovalLoadError('')
    })
    void getApprovalMode()
      .then((mode) => {
        if (alive && !changed) {
          setApprovalMode(mode)
          setApprovalModeReady(true)
        }
      })
      .catch((cause) => {
        if (alive && !changed) setApprovalLoadError(displayError(cause))
      })
    return () => {
      alive = false
      unsubscribe?.()
    }
  }, [displayError, approvalLoadAttempt, copy.agentFailed])

  useEffect(() => {
    followStream.current = true
    if (!activeSessionId) {
      setRunning(false)
      setTimeline(emptyTimelineState())
      resetDogVisual()
      return
    }
    let alive = true
    let historyReady = false
    const pendingEvents: WorkspaceAgentEvent[] = []
    const eventBuffer = createChatEventBuffer(handleEvents)
    setRunning(false)
    setTimeline(emptyTimelineState())
    resetDogVisual()
    const api = workspaceAgentApi()
    const unsubscribe = api.subscribe(activeSessionId, (event) => {
      if (!alive) return
      if (!historyReady) pendingEvents.push(event)
      else eventBuffer.push(event)
    })
    void Promise.all([
      api.getSessionMessages(activeSessionId),
      api.getSessionEvents(activeSessionId).catch(() => [] as WorkspaceAgentEvent[]),
      api.getSessionHistoryPrefix?.(activeSessionId) ?? Promise.resolve([]),
    ])
      .then(([history, persistedEvents, importedHistory]) => {
        if (!alive) return
        historyReady = true
        const buffered = pendingEvents.splice(0)
        // Harness events carry the durable sequence and preserve the actual
        // interleaving of assistant deltas, tool cards, and file activity.
        // Replaying those events as a base avoids appending each persisted
        // assistant delta a second time to the cumulative transcript text.
        // Legacy sessions may have no event log, so retain transcript replay
        // as the compatibility fallback.
        const hasConversationEvents = persistedEvents.some(
          (event) =>
            event.type === 'assistant.message' ||
            event.type === 'assistant.reasoning_summary' ||
            event.type === 'tool.started' ||
            event.type === 'tool.completed' ||
            event.type === 'tool.failed',
        )
        const importedTimeline = timelineStateFromHistory(importedHistory)
        // Imported messages are complete historical turns, not duplicates of this
        // Host's events (a new request may legitimately repeat the same words).
        importedTimeline.items = importedTimeline.items.map((item) => ({
          ...item,
          id: `imported-${item.id}`,
        }))
        const nextTimeline = hasConversationEvents
          ? replayTimelineEvents(importedTimeline, [...persistedEvents, ...buffered], {
              includeDogMemes: true,
            })
          : replayTimelineEvents(
              importedHistory.length ? importedTimeline : timelineStateFromHistory(history),
              buffered,
              {
                includeDogMemes: true,
              },
            )
        setTimeline(nextTimeline)
        if (!hasConversationEvents && persistedEvents.length === 0 && buffered.length === 0)
          applyDogVisual(dogVisualForTimeline(nextTimeline))
        for (const event of persistedEvents) applyEventSideEffects(event)
        for (const event of buffered) applyEventSideEffects(event)
      })
      .catch((cause) => {
        if (!alive) return
        setError(displayError(cause))
        historyReady = true
        handleEvents(pendingEvents.splice(0))
        pendingEvents.length = 0
      })
    return () => {
      alive = false
      eventBuffer.dispose()
      pendingEvents.length = 0
      unsubscribe()
    }
  }, [
    activeSessionId,
    activationEpoch,
    applyDogVisual,
    applyEventSideEffects,
    displayError,
    handleEvents,
    resetDogVisual,
  ])

  useEffect(() => {
    if (!workspace) return
    const unsubscribe =
      workspaceAgentApi().subscribeFiles?.(workspace.id, scheduleFilesRefresh) ?? (() => {})
    const unsubscribeLocations = window.aiOffice.onFolderChanged?.(() => scheduleFilesRefresh())
    const handleFocus = () => {
      scheduleFilesRefresh()
      void refreshAgentContext()
    }
    window.addEventListener('focus', handleFocus)
    return () => {
      unsubscribe()
      unsubscribeLocations?.()
      window.removeEventListener('focus', handleFocus)
    }
  }, [refreshAgentContext, refreshFiles, scheduleFilesRefresh, workspace])

  useEffect(() => {
    const node = scrollRef.current
    if (node && followStream.current) node.scrollTop = node.scrollHeight
  }, [timeline])

  // Keep the end in view across native resize/reparent and composer expansion,
  // but never pull a user away from older messages they intentionally scrolled to.
  useLayoutEffect(() => {
    const node = scrollRef.current
    if (node && followStream.current) node.scrollTop = node.scrollHeight
  }, [layout.width, layout.height, sidebarOpen, input])

  useEffect(
    () => () => {
      if (fileRefreshTimer.current) clearTimeout(fileRefreshTimer.current)
      if (dogTransitionTimer.current) clearTimeout(dogTransitionTimer.current)
    },
    [],
  )

  const createSession = async () => {
    const epoch = ++sendEpoch.current
    setAttachedSkills([])
    if (!workspace) return
    try {
      const session = await workspaceAgentApi().createSession(workspace.id)
      if (epoch !== sendEpoch.current) return
      setSessions((current) => [session, ...current.filter((item) => item.id !== session.id)])
      setActiveSessionId(session.id)
      setTimeline(emptyTimelineState())
      resetDogVisual()
      setError(null)
    } catch (cause) {
      setError(displayError(cause))
    }
  }

  const selectSession = (session: AgentSessionSummary) => {
    const epoch = ++sendEpoch.current
    setAttachedSkills([])
    setActiveSessionId(session.id)
    setError(null)
    if (session.workspaceId !== workspace?.id) {
      const api = workspaceAgentApi()
      void api
        .resumeSession(session.id)
        .then(async () => {
          if (epoch !== sendEpoch.current) return
          const descriptor = await api.getWorkspace()
          if (epoch !== sendEpoch.current) return
          await loadWorkspace(descriptor)
          if (epoch === sendEpoch.current) setActiveSessionId(session.id)
        })
        .catch((cause) => {
          if (epoch === sendEpoch.current) setError(displayError(cause))
        })
    }
  }

  const send = async (steer = false) => {
    if (contextMutationLock.current) return
    const epoch = ++sendEpoch.current
    const selectedSkills = [...attachedSkills]
    const draftOwner = activeSessionId
    const sentDraft = input
    const text = input.trim()
    if (!text) return
    followStream.current = true
    // Acknowledge the user's instruction immediately, before the async context
    // lookup and IPC round-trip begin.
    setDogEvent('received')
    setRunning(true)
    try {
      if (!steer && window.aiOffice.connectionSetup) {
        const connection = await window.aiOffice.connectionSetup.inspect()
        if (epoch !== sendEpoch.current) return
        if (!connection.usable) {
          setRunning(false)
          resetDogVisual()
          openOnboarding(true)
          return
        }
      }
      let sessionId = activeSessionId
      if (!sessionId && workspace) {
        const session = await workspaceAgentApi().createSession(workspace.id)
        if (epoch !== sendEpoch.current) return
        drafts.adoptPending(session.id)
        setSessions((current) => [session, ...current.filter((item) => item.id !== session.id)])
        setActiveSessionId(session.id)
        sessionId = session.id
      }
      if (!sessionId) {
        setRunning(false)
        setDogEvent('error')
        setError(copy.chooseWorkspaceError)
        return
      }
      const bodies = await Promise.all(
        selectedSkills.map(async (name) => ({ name, body: await window.aiOffice.skillText(name) })),
      )
      const instructions = bodies.length ? withAttachedSkills('', bodies) : undefined
      if (epoch !== sendEpoch.current) return
      const api = workspaceAgentApi()
      const context = api.prepareContext ? await api.prepareContext() : await refreshAgentContext()
      setAgentContext(context)
      if (epoch !== sendEpoch.current) return
      if (context?.documentContextError) throw new Error(context.documentContextError)
      const references = buildAgentReferences(text, mentionFiles, context, lang)
      setMentionRange(null)
      setError(null)
      const message = {
        requestId: crypto.randomUUID(),
        text,
        ...(context?.documentContext ? { contextId: context.documentContext.id } : {}),
        ...(references.length ? { references } : {}),
        ...(instructions ? { instructions } : {}),
      }
      if (steer) await workspaceAgentApi().steer(sessionId, message)
      else await workspaceAgentApi().send(sessionId, message)
      drafts.clearSent(draftOwner, sentDraft)
      drafts.clearSent(sessionId, sentDraft)
      refreshDraft((revision) => revision + 1)
    } catch (cause) {
      if (epoch !== sendEpoch.current) return
      setRunning(false)
      setDogEvent('error')
      setError(displayError(cause))
    }
  }

  const openWorkspacePicker = async (mode: 'project' | 'folder' = 'folder') => {
    if (mode === 'project' && window.managedSessions) {
      setWorkspacePickerOpen(false)
      setSessionManagerOpen(true)
      return
    }
    setWorkspacePickerMode(mode)
    setWorkspacePickerOpen(true)
    if (mode !== 'project') return
    setProjectOptionsLoading(true)
    try {
      const api = workspaceAgentApi()
      const next = await api.listProjectWorkspaces?.()
      setProjectOptions(Array.isArray(next) ? next : [])
    } catch (cause) {
      setProjectOptions([])
      setError(displayError(cause))
    } finally {
      setProjectOptionsLoading(false)
    }
  }

  const chooseWorkspace = async () => {
    try {
      setError(null)
      const next = await workspaceAgentApi().openWorkspace()
      if (next) {
        await loadWorkspace(next)
        setWorkspacePickerOpen(false)
      }
    } catch (cause) {
      setError(displayError(cause))
    }
  }

  const chooseProjectWorkspace = async (project: WorkspaceProjectOption) => {
    try {
      setError(null)
      setProjectOptionsLoading(true)
      const next = await workspaceAgentApi().setProjectWorkspace(project.id)
      await loadWorkspace(next)
      setWorkspacePickerOpen(false)
    } catch (cause) {
      setError(displayError(cause))
    } finally {
      setProjectOptionsLoading(false)
    }
  }

  const openFile = async (path: string) => {
    try {
      setError(null)
      await workspaceAgentApi().openFile(path)
    } catch (cause) {
      setError(displayError(cause))
    }
  }

  const cancelRun = async () => {
    sendEpoch.current++
    setRunning(false)
    if (!activeSessionId) return
    try {
      await workspaceAgentApi().cancel(activeSessionId)
    } catch (cause) {
      setError(displayError(cause))
    }
  }

  const beginSessionAction = (kind: 'rename' | 'delete', session: AgentSessionSummary) => {
    if (sessionSaveLock.current) return
    setSessionError('')
    setSessionName(workspaceSessionTitle(session.title, lang))
    setSessionAction({ kind, session })
  }

  const submitSessionAction = async () => {
    if (!sessionAction || sessionSaveLock.current) return
    if (sessionAction.kind === 'rename' && !sessionName.trim()) return
    const { session, kind } = sessionAction
    sessionSaveLock.current = true
    setSessionSaving(true)
    setSessionError('')
    try {
      if (kind === 'rename') {
        const next = await workspaceAgentApi().renameSession(session.id, sessionName.trim())
        setSessions((current) => current.map((item) => (item.id === next.id ? next : item)))
        setSessionAction(null)
        return
      }
      await workspaceAgentApi().deleteSession(session.id)
      const wasActive = activeSessionId === session.id
      setSessions((current) => {
        const next = current.filter((item) => item.id !== session.id)
        if (wasActive) setActiveSessionId(next[0]?.id ?? null)
        return next
      })
      if (wasActive) {
        setTimeline(emptyTimelineState())
        resetDogVisual()
      }
      setSessionAction(null)
    } catch (cause) {
      setSessionError(displayError(cause))
    } finally {
      sessionSaveLock.current = false
      setSessionSaving(false)
    }
  }

  const insertMention = (candidate: MentionCandidate) => {
    if (!mentionRange) return
    const next = `${input.slice(0, mentionRange.start)}@${candidate.label} ${input.slice(mentionRange.end)}`
    const cursor = mentionRange.start + candidate.label.length + 2
    setInput(next)
    setMentionRange(null)
    setMentionIndex(0)
    requestAnimationFrame(() => {
      inputRef.current?.focus()
      inputRef.current?.setSelectionRange(cursor, cursor)
    })
  }

  const resolvePermission = async (requestId: string, decision: ToolPermissionDecision) => {
    if (!activeSessionId) return
    try {
      setError(null)
      await workspaceAgentApi().resolvePermission(activeSessionId, requestId, decision)
    } catch (cause) {
      setError(displayError(cause))
    }
  }

  const changeApprovalMode = async (next: AgentApprovalMode) => {
    if (next === approvalMode || approvalSaveLock.current || !approvalModeReady) return
    approvalSaveLock.current = true
    const revision = approvalModeRevision.current
    try {
      setError(null)
      setApprovalModeSaving(true)
      const actual = await workspaceAgentApi().setApprovalMode(next)
      if (revision === approvalModeRevision.current) setApprovalMode(actual)
    } catch (cause) {
      throw new Error(displayError(cause), { cause })
    } finally {
      approvalSaveLock.current = false
      setApprovalModeSaving(false)
    }
  }

  const undoRun = async (runId: string) => {
    if (!activeSessionId || undoingRunId) return
    try {
      setError(null)
      setUndoingRunId(runId)
      await workspaceAgentApi().undoRun(activeSessionId, runId)
      await refreshFiles(false)
    } catch (cause) {
      setError(displayError(cause))
    } finally {
      setUndoingRunId(null)
    }
  }

  const openWordReview = async (runId: string, path: string) => {
    const api = workspaceAgentApi()
    if (!activeSessionId || !api.openWordReview) return
    try {
      await api.openWordReview(activeSessionId, runId, path)
    } catch (cause) {
      setError(displayError(cause))
    }
  }

  const panelStyle = {
    '--workspace-chat-sidebar-width': `${sidebarWidth}px`,
    '--conversation-height': `${layout.height}px`,
  } as CSSProperties
  const activeSession = sessions.find((session) => session.id === activeSessionId)
  const activeTitle =
    !activeSession ||
    (!activeSession.messageCount &&
      ['AI BigDog', 'AI 大狗', 'BigDog Agent', '大狗 Agent'].includes(activeSession.title))
      ? isChineseWorkspaceLanguage(lang)
        ? '新对话'
        : 'New conversation'
      : workspaceSessionTitle(activeSession.title || copy.activeSessionFallback, lang)
  const pendingPermission = timeline.items.find(
    (item) => item.kind === 'permission' && item.state === 'pending',
  )
  const activeTool = running
    ? timeline.items.find((item) => item.kind === 'tool' && item.state === 'running')
    : null
  const statusLabel = pendingPermission
    ? dogEventLabel('approval', lang)
    : activeTool?.kind === 'tool'
      ? workspaceToolLabel(activeTool.name, lang)
      : running
        ? dogEventLabel(
            dogVisual.mood === 'reading'
              ? 'read'
              : dogVisual.mood === 'writing'
                ? 'write'
                : 'thinking',
            lang,
          )
        : dogVisual.mood === 'complete'
          ? copy.complete
          : dogVisual.mood === 'error'
            ? copy.needsAttention
            : copy.ready

  return (
    <aside
      ref={panelRef}
      className={`workspace-chat-panel ai-bigdog-chat presentation-${presentation}${dragOver ? ' is-file-dragover' : ''}`}
      aria-label={copy.panelTitle}
      style={panelStyle}
      onDragEnter={(event) => {
        if (event.dataTransfer.types.includes('Files')) {
          event.preventDefault()
          event.stopPropagation()
          if (!running && !contextMutationLock.current) setDragOver(true)
        }
      }}
      onDragOver={(event) => {
        if (event.dataTransfer.types.includes('Files')) {
          event.preventDefault()
          event.stopPropagation()
          event.dataTransfer.dropEffect = running || contextMutationLock.current ? 'none' : 'copy'
        }
      }}
      onDragLeave={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setDragOver(false)
      }}
      onDrop={(event) => {
        if (event.dataTransfer.types.includes('Files')) {
          event.preventDefault()
          event.stopPropagation()
          setDragOver(false)
          void attachFiles(Array.from(event.dataTransfer.files))
        }
      }}
      onKeyDownCapture={(event) => {
        const modifier = /mac/i.test(navigator.platform) ? event.metaKey : event.ctrlKey
        if (
          presentation === 'window' &&
          modifier &&
          !event.altKey &&
          event.key.toLowerCase() === 'w'
        ) {
          event.preventDefault()
          event.stopPropagation()
          onClose()
        }
      }}
    >
      {dragOver && (
        <div className="workspace-chat-file-drop" aria-hidden="true">
          {isChineseWorkspaceLanguage(lang)
            ? '松开以添加 @文件引用'
            : 'Drop to add @file references'}
        </div>
      )}
      <input
        ref={filePickerRef}
        type="file"
        hidden
        multiple
        onChange={(event) => {
          const selected = Array.from(event.currentTarget.files ?? [])
          event.currentTarget.value = ''
          void attachFiles(selected)
        }}
      />
      {presentation === 'docked' && onResize && (
        <div
          className="workspace-chat-resize workspace-chat-resize-width"
          role="separator"
          aria-orientation="vertical"
          aria-label={copy.resizeChat}
          tabIndex={0}
          aria-valuenow={Math.round(size.width)}
          onKeyDown={(event) => {
            if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
              event.preventDefault()
              onResize({ ...size, width: size.width + (event.key === 'ArrowLeft' ? 20 : -20) })
            }
          }}
          onPointerDown={(event) => resizePanel('chat-width', event)}
        />
      )}
      <header className="workspace-chat-header" inert={sessionManagerOpen}>
        <div className="workspace-chat-title-wrap">
          <div className="workspace-chat-title-line">
            <DogMascot
              mood={dogVisual.mood}
              frame={dogVisual.frame}
              compact
              label={false}
              className="ai-bigdog-avatar"
            />
            <div className="workspace-chat-title">{copy.panelTitle}</div>
          </div>
          {presentation !== 'home' && (
            <div className="workspace-chat-workspace-name">
              <button
                type="button"
                className="workspace-chat-workspace-switch"
                onClick={() => void openWorkspacePicker()}
                aria-label={copy.changeWorkspace}
                aria-haspopup="dialog"
                aria-expanded={workspacePickerOpen}
                aria-controls="agent-workspace-picker"
                title={
                  workspace ? `${copy.changeWorkspace}: ${workspace.name}` : copy.chooseWorkspace
                }
              >
                <img src={workspaceFolderIcon} alt="" aria-hidden="true" />
                <span className="workspace-chat-workspace-label">{copy.workspacePanel}</span>
                <span className="workspace-chat-workspace-current">
                  {workspace ? workspace.name : copy.workspaceFallback}
                </span>
                {workspace && (
                  <span className="workspace-chat-workspace-source">
                    {workspace.managedProjectId || workspace.source === 'project'
                      ? copy.projectSpace
                      : copy.customFolder}
                  </span>
                )}
                <svg
                  width="12"
                  height="12"
                  viewBox="0 0 16 16"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  aria-hidden="true"
                >
                  <path d="m4 6 4 4 4-4" />
                </svg>
              </button>
            </div>
          )}
        </div>
        <div className="workspace-chat-header-actions">
          {window.managedSessions && (
            <button
              type="button"
              className="workspace-chat-icon-btn ai-bigdog-memory-button"
              onClick={(event) => openMemoryManager(event.currentTarget)}
              aria-label={isChineseWorkspaceLanguage(lang) ? '记忆管理' : 'Memory'}
            >
              {isChineseWorkspaceLanguage(lang) ? '\u8bb0\u5fc6\u7ba1\u7406' : 'Memory'}
            </button>
          )}
          {onPresentationChange && (
            <button
              type="button"
              className="workspace-chat-icon-btn"
              onClick={onPresentationChange}
              aria-label={presentation === 'window' ? copy.dockWindow : copy.openWindow}
              data-tip={presentation === 'window' ? copy.dockWindow : copy.openWindow}
            >
              <svg
                width="17"
                height="17"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.6"
                aria-hidden="true"
              >
                {presentation === 'window' ? (
                  <>
                    <rect x="3" y="4" width="18" height="16" rx="2" />
                    <path d="M14 4v16M6 12h5m-2-2 2 2-2 2" />
                  </>
                ) : (
                  <>
                    <path d="M10 5H5a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-5M14 3h7v7M21 3 10 14" />
                  </>
                )}
              </svg>
            </button>
          )}
          {presentation !== 'home' && (
            <button
              className="workspace-chat-icon-btn"
              onClick={onClose}
              aria-label={copy.hideChat}
              data-tip={copy.hideKeepsRunning}
            >
              &#x2715;
            </button>
          )}
        </div>
      </header>

      {sessionManagerOpen && !advancedSessions && window.managedSessions && (
        <MemoryPanel
          lang={lang}
          currentSessionId={activeSessionId}
          onClose={closeMemoryManager}
          onAdvanced={() => setAdvancedSessions(true)}
        />
      )}
      {sessionManagerOpen && advancedSessions && window.managedSessions && (
        <ManagedSessionsPanel
          currentSessionId={activeSessionId}
          onLoadMemory={async (id) => {
            await loadSessionMemory(id)
            setSessionManagerOpen(false)
            setAdvancedSessions(false)
          }}
          lang={lang}
          onClose={() => setAdvancedSessions(false)}
          onFolder={() => {
            setSessionManagerOpen(false)
            setAdvancedSessions(false)
            void openWorkspacePicker('folder')
          }}
          onResume={async (id) => {
            await loadWorkspace(await workspaceAgentApi().getWorkspace())
            setActiveSessionId(id)
            setActivationEpoch((value) => value + 1)
            setSessionManagerOpen(false)
            setAdvancedSessions(false)
          }}
        />
      )}
      {sessionAction && (
        <form
          className="workspace-chat-inline-review workspace-chat-session-review"
          onSubmit={(event) => {
            event.preventDefault()
            void submitSessionAction()
          }}
          onKeyDown={(event) => {
            if (event.key === 'Escape' && !sessionSaving) {
              event.stopPropagation()
              setSessionAction(null)
            }
          }}
        >
          {sessionAction.kind === 'rename' ? (
            <label>
              {copy.renameSession(sessionName)}
              <input
                autoFocus
                value={sessionName}
                disabled={sessionSaving}
                onChange={(event) => setSessionName(event.target.value)}
              />
            </label>
          ) : (
            <p>{copy.deleteSession(workspaceSessionTitle(sessionAction.session.title, lang))}</p>
          )}
          {sessionError && <p role="alert">{sessionError}</p>}
          <div className="workspace-chat-inline-actions">
            <button type="button" disabled={sessionSaving} onClick={() => setSessionAction(null)}>
              {copy.cancel}
            </button>
            <button
              type="submit"
              className="confirm"
              disabled={sessionSaving || (sessionAction.kind === 'rename' && !sessionName.trim())}
            >
              {sessionSaving
                ? copy.saving
                : sessionAction.kind === 'rename'
                  ? copy.save
                  : copy.delete}
            </button>
          </div>
        </form>
      )}

      {workspacePickerOpen && (
        <div
          className="workspace-chat-picker"
          id="agent-workspace-picker"
          role="dialog"
          aria-label={copy.workspacePickerTitle}
          onKeyDown={(event) => {
            if (event.key === 'Escape') {
              event.stopPropagation()
              setWorkspacePickerOpen(false)
            }
          }}
        >
          <div className="workspace-chat-picker-head">
            <strong>{copy.workspacePickerTitle}</strong>
            <button
              type="button"
              className="workspace-chat-icon-btn"
              onClick={() => setWorkspacePickerOpen(false)}
              aria-label={copy.close}
              data-tip={copy.close}
            >
              &#x2715;
            </button>
          </div>
          <div className="workspace-chat-picker-tabs" role="tablist">
            <button
              type="button"
              role="tab"
              aria-selected={workspacePickerMode === 'project'}
              className={workspacePickerMode === 'project' ? 'active' : ''}
              onClick={() => void openWorkspacePicker('project')}
            >
              {copy.projectWorkspaces}
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={workspacePickerMode === 'folder'}
              className={workspacePickerMode === 'folder' ? 'active' : ''}
              onClick={() => setWorkspacePickerMode('folder')}
            >
              {copy.customFolder}
            </button>
          </div>
          {workspacePickerMode === 'project' ? (
            <div className="workspace-chat-project-options">
              {projectOptionsLoading ? (
                <div className="workspace-chat-muted">{copy.loadingWorkspace}</div>
              ) : projectOptions.length > 0 ? (
                projectOptions.map((project) => (
                  <button
                    type="button"
                    className="workspace-chat-project-option"
                    key={project.id}
                    onClick={() => void chooseProjectWorkspace(project)}
                    disabled={projectOptionsLoading}
                  >
                    <img src={workspaceFolderIcon} alt="" aria-hidden="true" />
                    <span>
                      <strong>{project.name}</strong>
                      <small>{copy.projectFileCount(project.fileCount)}</small>
                    </span>
                  </button>
                ))
              ) : (
                <div className="workspace-chat-muted">{copy.noProjectWorkspaces}</div>
              )}
            </div>
          ) : (
            <div className="workspace-chat-folder-picker">
              <p>{copy.chooseWorkspaceDescription}</p>
              <button
                type="button"
                className="workspace-chat-primary-btn"
                onClick={() => void chooseWorkspace()}
              >
                {copy.chooseFolder}
              </button>
            </div>
          )}
        </div>
      )}

      {loading ? (
        <div className="workspace-chat-empty">{copy.loadingWorkspace}</div>
      ) : !workspace ? (
        <div className="workspace-chat-empty workspace-chat-empty-action">
          <DogIllustration scene="welcome" className="workspace-chat-empty-mascot" />
          <div className="workspace-chat-empty-title">{copy.chooseWorkspace}</div>
          <p>{copy.chooseWorkspaceDescription}</p>
          <button className="workspace-chat-primary-btn" onClick={() => void openWorkspacePicker()}>
            {copy.chooseFolder}
          </button>
        </div>
      ) : (
        <div
          inert={sessionManagerOpen}
          className={`workspace-chat-grid ${sidebarOpen ? 'sidebar-open' : 'sidebar-closed'}`}
        >
          <aside
            className="workspace-chat-sidebar"
            ref={sidebarRef}
            id="agent-workspace-sidebar"
            hidden={!sidebarOpen}
          >
            <div className="workspace-chat-section-head">
              <span>{copy.sessions}</span>
              <button
                className="workspace-chat-small-btn"
                onClick={() => void createSession()}
                aria-label={copy.newSession}
                data-tip={copy.newSession}
              >
                +
              </button>
            </div>
            <div className="workspace-chat-session-list">
              {sessions.length === 0 && (
                <div className="workspace-chat-muted">{copy.noSessions}</div>
              )}
              {sessions.map((session) => (
                <div
                  className={`workspace-chat-session ${session.id === activeSessionId ? 'active' : ''}`}
                  key={session.id}
                >
                  <button
                    className="workspace-chat-session-main"
                    onClick={() => selectSession(session)}
                  >
                    <span className="workspace-chat-session-name">
                      {workspaceSessionTitle(session.title, lang)}
                    </span>
                    <span className="workspace-chat-session-meta">
                      {copy.messageCount(session.messageCount)}
                    </span>
                  </button>
                  <button
                    className="workspace-chat-session-menu"
                    disabled={sessionSaving}
                    onClick={() => beginSessionAction('rename', session)}
                    aria-label={`${copy.rename} ${workspaceSessionTitle(session.title, lang)}`}
                    data-tip={copy.rename}
                  >
                    &#x22EF;
                  </button>
                  <button
                    className="workspace-chat-session-menu danger"
                    disabled={sessionSaving}
                    onClick={() => beginSessionAction('delete', session)}
                    aria-label={`${copy.delete} ${workspaceSessionTitle(session.title, lang)}`}
                    data-tip={copy.delete}
                  >
                    &#x2715;
                  </button>
                </div>
              ))}
            </div>
            <div className="workspace-chat-section-head workspace-chat-files-head">
              <span>{copy.files}</span>
              <button
                className="workspace-chat-small-btn"
                onClick={() => void refreshFiles()}
                aria-label={copy.refreshFiles}
                data-tip={copy.refreshFiles}
              >
                &#x21BB;
              </button>
            </div>
            <div className="workspace-chat-file-list">
              {referenceFiles.length > 0 && (
                <div className="workspace-chat-reference-list">
                  <div className="workspace-chat-muted">
                    {isChineseWorkspaceLanguage(lang) ? '本会话引用' : 'Conversation references'}
                  </div>
                  {referenceFiles.map((file) => (
                    <button
                      type="button"
                      className="workspace-chat-file"
                      key={file.path}
                      title={file.path}
                      onClick={() => void openFile(file.path)}
                    >
                      <span className={`workspace-chat-file-type ${file.type}`}>
                        {workspaceFileLabel(file)}
                      </span>
                      <span className="workspace-chat-file-name">{file.name}</span>
                    </button>
                  ))}
                </div>
              )}
              {files.length === 0 && referenceFiles.length === 0 && (
                <div className="workspace-chat-muted">{copy.noFiles}</div>
              )}
              {fileTree.map((node) => (
                <FileTreeBranch
                  key={node.path}
                  node={node}
                  onOpen={(path) => void openFile(path)}
                />
              ))}
              {files
                .filter((entry) => entry.registered)
                .map((entry) => (
                  <FileTreeBranch
                    key={`location:${entry.path}`}
                    node={{
                      path: entry.path,
                      name: entry.name,
                      directory: entry.directory,
                      entry,
                      children: [],
                      lazy: entry.directory,
                    }}
                    onOpen={(path) => void openFile(path)}
                  />
                ))}
            </div>
          </aside>
          <div
            className="workspace-chat-inner-resize workspace-chat-sidebar-resize"
            hidden={!sidebarOpen}
            role="separator"
            aria-orientation="vertical"
            aria-label={copy.resizeWorkspace}
            onPointerDown={(event) => resizePanel('sidebar', event)}
          />

          <section className="workspace-chat-main" ref={mainRef}>
            <div className="workspace-chat-main-head">
              <button
                type="button"
                className="workspace-chat-icon-btn workspace-chat-navigation-toggle"
                aria-label={copy.toggleSidebar}
                data-tip={copy.toggleSidebar}
                aria-expanded={sidebarOpen}
                aria-controls="agent-workspace-sidebar"
                onClick={() =>
                  setSidebarVisibility((current) => ({
                    ...current,
                    [presentation]: !current[presentation],
                  }))
                }
              >
                <svg
                  width="17"
                  height="17"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.6"
                  aria-hidden="true"
                >
                  <rect x="3" y="4" width="18" height="16" rx="2" />
                  <path d="M9 4v16" />
                </svg>
                <span>
                  {presentation === 'home'
                    ? isChineseWorkspaceLanguage(lang)
                      ? '历史与文件'
                      : 'History & files'
                    : copy.workspacePanel}
                </span>
              </button>
              <div className="workspace-chat-active-title" title={activeTitle}>
                {activeTitle}
              </div>
              <ConversationRecords
                lang={lang}
                sessionId={activeSessionId}
                revision={timeline.items
                  .filter((item) => item.kind === 'run')
                  .map((item) => (item.kind === 'run' ? `${item.run.id}:${item.run.state}` : ''))
                  .join('|')}
                undoingRunId={undoingRunId}
                onUndo={(runId) => void undoRun(runId)}
                onWordReview={(runId, path) => void openWordReview(runId, path)}
                onOpen={(path, artifactId) =>
                  artifactId && activeSessionId && window.aiOfficeTasks
                    ? void window.aiOfficeTasks
                        .openArtifact(activeSessionId, artifactId)
                        .catch((cause) => setError(displayError(cause)))
                    : void openFile(path)
                }
              />
              <button
                type="button"
                className="workspace-chat-icon-btn"
                onClick={() => void createSession()}
                aria-label={copy.newSession}
                data-tip={copy.newSession}
              >
                +
              </button>
            </div>
            <div
              className="workspace-chat-scroll"
              ref={scrollRef}
              onScroll={(event) => {
                const node = event.currentTarget
                followStream.current = node.scrollHeight - node.scrollTop - node.clientHeight < 64
              }}
            >
              {timeline.items.length === 0 && !running && (
                <div className="workspace-chat-welcome">
                  <DogIllustration scene="welcome" className="workspace-chat-welcome-mascot" />
                  <div className="workspace-chat-welcome-title">{copy.welcomeTitle}</div>
                  <p>{copy.welcomeDescription}</p>
                </div>
              )}
              {timelineForDisplay(timeline, running || !!pendingPermission).map((item) => {
                if (item.kind === 'work')
                  return <WorkspaceToolSteps key={item.id} group={item} lang={lang} />
                if (item.kind === 'message') {
                  return (
                    <div className={`workspace-chat-message ${item.role}`} key={item.id}>
                      <div className="workspace-chat-message-role">
                        {item.role === 'assistant' && (
                          <img
                            className="workspace-chat-message-avatar"
                            src={chatIcon}
                            alt=""
                            aria-hidden="true"
                          />
                        )}
                        {item.role === 'user' ? copy.you : copy.agent}
                      </div>
                      <div className="workspace-chat-message-text">
                        {item.role === 'assistant' ? (
                          <WorkspaceMarkdown value={item.text} />
                        ) : (
                          item.text
                        )}
                      </div>
                    </div>
                  )
                }
                if (item.kind === 'meme') {
                  return (
                    <div
                      className={`workspace-chat-dog-meme workspace-chat-dog-meme-${item.mood}`}
                      key={item.id}
                      role="status"
                      aria-label={dogEventLabel(item.event, lang)}
                    >
                      <img
                        src={item.frame}
                        className="workspace-chat-dog-meme-image"
                        alt={dogEventLabel(item.event, lang)}
                        loading="lazy"
                      />
                      <div className="workspace-chat-dog-meme-copy">
                        <span className="workspace-chat-dog-meme-role">
                          <img src={chatIcon} alt="" aria-hidden="true" />
                          {copy.agent}
                        </span>
                        <span className="workspace-chat-dog-meme-label">
                          {dogEventLabel(item.event, lang)}
                        </span>
                      </div>
                    </div>
                  )
                }
                if (item.kind === 'reasoning') {
                  return (
                    <details className="workspace-chat-reasoning" key={item.id}>
                      <summary>{copy.reasoningSummary}</summary>
                      <div>{item.text}</div>
                    </details>
                  )
                }
                if (item.kind === 'permission') {
                  return (
                    <div className={`workspace-chat-permission ${item.state}`} key={item.id}>
                      <div className="workspace-chat-permission-copy">
                        <strong>{workspacePermissionText(item.title, lang)}</strong>
                        <span>{workspacePermissionText(item.description, lang)}</span>
                        {item.state === 'pending' && item.round && (
                          <small title={item.round.baseVersion}>
                            {copy.roundApprovalScope} · {item.round.baseVersion.slice(0, 12)}
                          </small>
                        )}
                      </div>
                      {item.state === 'pending' ? (
                        <div className="workspace-chat-permission-actions">
                          <button
                            type="button"
                            onClick={() => void resolvePermission(item.requestId, 'deny')}
                          >
                            {copy.deny}
                          </button>
                          <button
                            type="button"
                            className={item.round ? undefined : 'primary'}
                            onClick={() => void resolvePermission(item.requestId, 'allow-once')}
                          >
                            {copy.allowOnce}
                          </button>
                          {item.round && (
                            <button
                              type="button"
                              className="primary"
                              onClick={() => void resolvePermission(item.requestId, 'allow-round')}
                            >
                              {copy.allowRound}
                            </button>
                          )}
                        </div>
                      ) : (
                        <span className="workspace-chat-permission-state">
                          {item.reason === 'expired'
                            ? copy.expired
                            : item.reason === 'cancelled'
                              ? copy.workStopped
                            : item.state === 'allowed'
                              ? item.round?.approved ? copy.allowedRound : copy.allowedOnce
                              : copy.denied}
                        </span>
                      )}
                    </div>
                  )
                }
                if (item.kind === 'run') {
                  return (
                    <RunReviewCard
                      key={item.id}
                      run={item.run}
                      lang={lang}
                      onUndo={(runId) => void undoRun(runId)}
                      onWordReview={(runId, path) => void openWordReview(runId, path)}
                      compact
                      onOpen={(path, artifactId) =>
                        artifactId && activeSessionId && window.aiOfficeTasks
                          ? void window.aiOfficeTasks
                              .openArtifact(activeSessionId, artifactId)
                              .catch((cause) => setError(displayError(cause)))
                          : void openFile(path)
                      }
                      undoing={undoingRunId === item.run.id}
                      undoDisabled={undoingRunId !== null || running}
                    />
                  )
                }
              })}
              {activeSessionId && (
                <AssistantTaskShelf
                  key={activeSessionId}
                  lang={lang}
                  sessionId={activeSessionId}
                  onOpenFile={(path) => void openFile(path)}
                />
              )}
            </div>
            {error && (
              <div className="workspace-chat-error" role="alert">
                {error}
              </div>
            )}
            <form
              ref={composerRef}
              className="workspace-chat-composer"
              onSubmit={(event) => {
                event.preventDefault()
                void send(running)
              }}
            >
              <details
                className={`workspace-chat-progress ${running ? 'running' : ''} ${pendingPermission ? 'waiting' : ''}`}
              >
                <summary>
                  <span className="workspace-chat-status-dot" aria-hidden="true" />
                  <span role="status">{statusLabel}</span>
                  <span className="workspace-chat-progress-hint">{copy.statusDetails}</span>
                </summary>
                <div className="workspace-chat-progress-detail">
                  <DogMascot mood={dogVisual.mood} frame={dogVisual.frame} compact label={false} />
                  <span>{copy.hideKeepsRunning}</span>
                  {pendingPermission && (
                    <button
                      type="button"
                      className="workspace-chat-secondary-btn"
                      onClick={() => {
                        scrollRef.current
                          ?.querySelector('.workspace-chat-permission.pending')
                          ?.scrollIntoView({ block: 'center' })
                      }}
                    >
                      {copy.viewApproval}
                    </button>
                  )}
                </div>
              </details>
              <div className="workspace-chat-input-anchor">
                <textarea
                  ref={inputRef}
                  aria-autocomplete="list"
                  aria-controls={mentionsOpen ? mentionListId : undefined}
                  aria-activedescendant={
                    mentionsOpen
                      ? `${mentionListId}-${Math.min(mentionIndex, visibleMentions.length - 1)}`
                      : undefined
                  }
                  onInput={resizeComposer}
                  value={input}
                  onFocus={(event) => {
                    void refreshAgentContext()
                    setMentionRange(
                      mentionRangeAtCursor(
                        event.currentTarget.value,
                        event.currentTarget.selectionStart,
                      ),
                    )
                  }}
                  onBlur={() => window.setTimeout(() => setMentionRange(null), 80)}
                  onChange={(event) => {
                    const value = event.currentTarget.value
                    setInput(value)
                    setMentionRange(mentionRangeAtCursor(value, event.currentTarget.selectionStart))
                    setMentionIndex(0)
                  }}
                  onSelect={(event) => {
                    const node = event.currentTarget
                    setMentionRange(
                      node.selectionStart === node.selectionEnd
                        ? mentionRangeAtCursor(node.value, node.selectionStart)
                        : null,
                    )
                  }}
                  onKeyDown={(event) => {
                    if (event.nativeEvent.isComposing || event.keyCode === 229) return
                    if (mentionRange && visibleMentions.length > 0) {
                      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                        event.preventDefault()
                        const direction = event.key === 'ArrowDown' ? 1 : -1
                        setMentionIndex(
                          (current) =>
                            (current + direction + visibleMentions.length) % visibleMentions.length,
                        )
                        return
                      }
                      if (event.key === 'Escape') {
                        event.preventDefault()
                        setMentionRange(null)
                        return
                      }
                      if (event.key === 'Enter' && !event.shiftKey) {
                        event.preventDefault()
                        insertMention(
                          visibleMentions[Math.min(mentionIndex, visibleMentions.length - 1)]!,
                        )
                        return
                      }
                    }
                    if (event.key === 'Enter' && !event.shiftKey) {
                      event.preventDefault()
                      void send(running)
                    }
                  }}
                  placeholder={copy.askPlaceholder}
                  onPaste={(event) => {
                    if (event.clipboardData.files.length) {
                      event.preventDefault()
                      void attachFiles(Array.from(event.clipboardData.files))
                    }
                  }}
                  aria-label={copy.messageAria}
                  rows={1}
                />
                {mentionsOpen && (
                  <div
                    ref={mentionListRef}
                    id={mentionListId}
                    className="workspace-chat-mention-picker"
                    role="listbox"
                    aria-label={
                      isChineseWorkspaceLanguage(lang)
                        ? '引用文件与上下文'
                        : 'File and context references'
                    }
                    data-side={mentionPlacement.side}
                    style={{ maxHeight: mentionPlacement.height }}
                  >
                    {visibleMentions.map((candidate, index) => (
                      <button
                        type="button"
                        role="option"
                        tabIndex={-1}
                        id={`${mentionListId}-${index}`}
                        aria-selected={index === Math.min(mentionIndex, visibleMentions.length - 1)}
                        className={
                          index === Math.min(mentionIndex, visibleMentions.length - 1)
                            ? 'active'
                            : ''
                        }
                        key={`${candidate.kind}:${candidate.path ?? candidate.label}`}
                        title={`${candidate.label}\n${candidate.detail}`}
                        onMouseDown={(event) => event.preventDefault()}
                        onClick={() => insertMention(candidate)}
                      >
                        <span>@{candidate.label}</span>
                        <small>{candidate.detail}</small>
                      </button>
                    ))}
                  </div>
                )}
              </div>
              {referenceFiles.length > 0 && (
                <div
                  className="workspace-chat-reference-chips"
                  aria-label={
                    isChineseWorkspaceLanguage(lang)
                      ? '会话参考文件'
                      : 'Conversation reference files'
                  }
                >
                  {referenceFiles.map((file) => (
                    <div className="workspace-chat-reference-chip" key={file.path}>
                      <button
                        type="button"
                        title={file.path}
                        onClick={() => void openFile(file.path)}
                      >
                        @{file.name}
                      </button>
                      <button
                        type="button"
                        disabled={running || attachingFiles || contextSaving}
                        aria-label={`${isChineseWorkspaceLanguage(lang) ? '移除引用' : 'Remove reference'} ${file.name}`}
                        onClick={() =>
                          void removeReference(file.path).catch((cause) =>
                            setError(displayError(cause)),
                          )
                        }
                      >
                        ×
                      </button>
                    </div>
                  ))}
                </div>
              )}
              {attachingFiles && (
                <div className="workspace-chat-muted" role="status">
                  {isChineseWorkspaceLanguage(lang)
                    ? '正在添加文件引用…'
                    : 'Adding file references…'}
                </div>
              )}
              <div className="workspace-chat-composer-actions">
                {(agentContext?.documentContext || agentContext?.documentContextError) && (
                  <div className="workspace-chat-current-document">
                    <span>
                      {agentContext.documentContext?.name ??
                        displayError(agentContext.documentContextError)}
                    </span>
                    {agentContext.documentContext?.dirty && (
                      <span>
                        {isChineseWorkspaceLanguage(lang)
                          ? '含未保存内容'
                          : 'Includes unsaved changes'}
                      </span>
                    )}
                    <button
                      type="button"
                      disabled={running}
                      aria-label={
                        isChineseWorkspaceLanguage(lang)
                          ? '移除当前文档上下文'
                          : 'Remove current document context'
                      }
                      onClick={() => {
                        void workspaceAgentApi()
                          .clearDocumentContext()
                          .then(refreshAgentContext)
                          .catch((cause) => setError(displayError(cause)))
                      }}
                    >
                      ×
                    </button>
                  </div>
                )}
                <div className="workspace-chat-context-summary">
                  <button
                    ref={contextTriggerRef}
                    type="button"
                    className="workspace-chat-context-trigger"
                    aria-label={zh ? '文件与记忆' : 'Files & memory'}
                    aria-expanded={contextOpen}
                    aria-controls={contextOpen ? contextPanelId : undefined}
                    onClick={() => void showContext()}
                  >
                    {zh ? '文件与记忆' : 'Files & memory'}
                    {referenceFiles.length > 0 && (
                      <span data-file-count="">
                        {referenceFiles.length}{' '}
                        {zh ? '个文件' : referenceFiles.length === 1 ? 'file' : 'files'}
                      </span>
                    )}
                    <svg
                      viewBox="0 0 16 16"
                      width="13"
                      height="13"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="1.5"
                      aria-hidden="true"
                    >
                      <path d="m4 10 4-4 4 4" />
                    </svg>
                  </button>
                  {!contextLoadingError && (!activeSessionId || selectedContext) && (
                    <span
                      className="workspace-chat-context-status"
                      data-enabled={selectedContext?.workspaceMemory ?? true}
                    >
                      {selectedContext?.workspaceMemory === false
                        ? zh ? '自动记忆已关' : 'Auto memory off'
                        : zh ? '自动记忆已开' : 'Auto memory on'}
                    </span>
                  )}
                </div>
                {presentation === 'home' && (
                  <div className="ai-bigdog-model-actions">
                    <AiModelPicker bridge={modelBridge} lang={lang} />
                    {timeline.usage && <AiUsageChip usage={timeline.usage} session lang={lang} />}
                    <AiSkillsMenu
                      lang={lang}
                      listSkills={async () =>
                        (await window.aiOffice.skillsList()).map(({ name, description }) => ({
                          name,
                          description,
                        }))
                      }
                      selected={attachedSkills}
                      onToggle={(name) =>
                        setAttachedSkills((previous) =>
                          previous.includes(name)
                            ? previous.filter((n) => n !== name)
                            : [...previous, name],
                        )
                      }
                      onAddFile={() => {
                        if (!running && !contextMutationLock.current) filePickerRef.current?.click()
                      }}
                    />
                  </div>
                )}
                {attachedSkills.length > 0 && (
                  <AiSkillChips
                    names={attachedSkills}
                    lang={lang}
                    onRemove={(name) =>
                      setAttachedSkills((previous) => previous.filter((n) => n !== name))
                    }
                  />
                )}

                <AgentApprovalControl
                  mode={approvalMode}
                  ready={approvalModeReady}
                  saving={approvalModeSaving}
                  onChange={changeApprovalMode}
                  copy={copy}
                  loadError={approvalLoadError}
                  onRetry={() => setApprovalLoadAttempt((attempt) => attempt + 1)}
                  actions={
                    <div className="workspace-chat-send-actions">
                      {running && (
                        <button
                          type="button"
                          className="workspace-chat-secondary-btn"
                          onClick={() => void cancelRun()}
                        >
                          {copy.stop}
                        </button>
                      )}
                      <button
                        type="submit"
                        className="workspace-chat-primary-btn"
                        title={copy.enterHint}
                        disabled={
                          attachingFiles ||
                          contextSaving ||
                          !input.trim() ||
                          (!activeSessionId && !workspace)
                        }
                      >
                        {running ? copy.steer : copy.send}
                      </button>
                    </div>
                  }
                />
              </div>
            </form>
            {contextOpen && activeSessionId && (
              <div ref={contextLayerRef} className="workspace-chat-context-layer">
                <ConversationContextPanel
                  key={activeSessionId}
                  id={contextPanelId}
                  sessionId={activeSessionId}
                  lang={lang}
                  selection={selectedContext}
                  preview={contextPreview}
                  files={referenceFiles}
                  busy={running || attachingFiles || contextSaving}
                  error={contextLoadingError}
                  onClose={closeContext}
                  onManageMemory={
                    window.managedSessions ? () => openMemoryManager(contextTriggerRef.current) : undefined
                  }
                  onAddFiles={() => filePickerRef.current?.click()}
                  onOpenFile={(path) => {
                    setContextOpen(false)
                    void openFile(path)
                  }}
                  onRemoveFile={removeReference}
                  onChange={updateContextSelection}
                />
              </div>
            )}
          </section>
        </div>
      )}
    </aside>
  )
}
