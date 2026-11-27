import { createHash, randomUUID } from 'node:crypto'
import { open, stat } from 'node:fs/promises'
import type { DocumentSnapshot, LiveDocumentAdapter } from '../workspace-agent/service'
import {
  documentChanges,
  evidenceSnapshotDigest,
  observeDocument,
  observeResourceChanges,
} from '../workspace-agent/change-evidence'
import { saveWitness, type SaveWitness } from '@genoffice/electron-utils/save-witness'
import type { AgentRunFileChange } from '../../shared/workspace-agent-api'
import type { TaskTargetVersion } from '../../shared/assistant-task-api'

export interface DocumentObservation {
  currentWordBytes?: () => Promise<Uint8Array>
  version: TaskTargetVersion
  liveSnapshot?: boolean
  snapshot?: DocumentSnapshot
  packageSnapshot?: DocumentSnapshot
  witness?: SaveWitness
}
const CAP = 64 * 1024 * 1024
function identity(info: Awaited<ReturnType<typeof stat>>): string {
  return `${info.dev}:${info.ino}:${info.birthtimeMs}`
}
function digest(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

/** The caller supplies a host-resolved path and an adapter belonging to that exact file. */
export async function observeTaskDocument(input: {
  path: string
  type?: TaskTargetVersion['type']
  live?: LiveDocumentAdapter | null
  editorIdentity?: string
  dirty?: boolean
}): Promise<DocumentObservation> {
  const { path, type, live, editorIdentity, dirty } = input
  const version: TaskTargetVersion = { path, type, exists: false, complete: true }
  let packageSnapshot: DocumentSnapshot | undefined
  const file = await open(path, 'r').catch((error) => {
    if (error?.code === 'ENOENT') return null
    throw error
  })
  if (file) {
    try {
      const info = await file.stat()
      if (!info.isFile()) throw new Error('task_target_not_file')
      version.exists = true
      version.fileIdentity = identity(info)
      if (info.size > CAP) version.complete = false
      else {
        const bytes = Buffer.alloc(info.size)
        let offset = 0
        while (offset < bytes.length) {
          const result = await file.read(bytes, offset, bytes.length - offset, offset)
          if (!result.bytesRead) break
          offset += result.bytesRead
        }
        const after = await stat(path)
        if (
          offset !== bytes.length ||
          identity(info) !== identity(after) ||
          info.mtimeMs !== after.mtimeMs ||
          info.size !== after.size
        )
          version.complete = false
        else {
          version.diskDigest = digest(bytes)
          if (type)
            packageSnapshot = {
              path,
              type,
              bytesBase64: bytes.toString('base64'),
              createdAt: new Date().toISOString(),
            }
        }
      }
    } finally {
      await file.close()
    }
  }
  let snapshot: DocumentSnapshot | undefined
  if (live) {
    snapshot = await live.snapshot()
    version.modelDigest = evidenceSnapshotDigest(snapshot, true)
    version.editorIdentity = editorIdentity
    // Tab chrome can still show the pre-save flag while the writer has already committed.
    const liveDirty = await live.taskDirty?.()
    version.dirty = !version.exists || (liveDirty ?? dirty ?? false)
  }
  return {
    version,
    ...(type === 'docx' && live?.exportBytes
      ? { currentWordBytes: () => live.exportBytes!() }
      : {}),
    snapshot: snapshot ?? packageSnapshot,
    packageSnapshot,
    liveSnapshot: !!live,
    witness: version.exists ? saveWitness(path) : undefined,
  }
}

/** Clean documents may reopen on a new boot; unsaved content must retain its original live identity. */
export function sameTaskTarget(expected: TaskTargetVersion, current: TaskTargetVersion): boolean {
  if (
    !expected.complete ||
    !current.complete ||
    expected.path !== current.path ||
    expected.exists !== current.exists
  )
    return false
  if (
    expected.exists &&
    (expected.fileIdentity !== current.fileIdentity || expected.diskDigest !== current.diskDigest)
  )
    return false
  if (expected.dirty || current.dirty)
    return (
      expected.editorIdentity === current.editorIdentity &&
      !!expected.modelDigest &&
      expected.modelDigest === current.modelDigest
    )
  if (
    expected.editorIdentity &&
    expected.editorIdentity === current.editorIdentity &&
    expected.modelDigest &&
    current.modelDigest
  )
    return expected.modelDigest === current.modelDigest
  return true
}
/** Advance an attempt-local grant only across its own edits to the same file generation. */
export function nextRoundVersion(
  previous: TaskTargetVersion,
  next?: TaskTargetVersion,
): TaskTargetVersion | undefined {
  return next?.complete &&
    previous.path === next.path &&
    previous.diskDigest === next.diskDigest &&
    previous.fileIdentity === next.fileIdentity &&
    (!previous.editorIdentity || next.editorIdentity === previous.editorIdentity)
    ? next
    : undefined
}

/** Retain the round only across the same editor model or the exact staged bytes that were saved. */
export function roundAfterWordSave(
  previous: TaskTargetVersion,
  current: DocumentObservation,
): TaskTargetVersion | undefined {
  const next = current.version
  if (
    previous.type !== 'docx' ||
    next.type !== 'docx' ||
    previous.path !== next.path ||
    !next.complete ||
    !next.exists ||
    next.dirty
  )
    return
  if (sameTaskTarget(previous, next)) return next
  if (!previous.modelDigest || !next.diskDigest) return
  if (current.liveSnapshot) {
    if (
      previous.editorIdentity &&
      previous.editorIdentity === next.editorIdentity &&
      previous.modelDigest === next.modelDigest &&
      current.witness?.complete &&
      current.witness.sha256 === next.diskDigest
    )
      return next
  } else if (previous.modelDigest === next.diskDigest) return next
}

/** Match the host writer's promoted bytes to the file that was read back after a save. */
export function verifiedWordSave(
  before: DocumentObservation | undefined,
  after: DocumentObservation,
): boolean {
  const version = after.version
  if (!version.complete || !version.exists || !version.diskDigest || version.dirty) return false
  if (after.liveSnapshot) {
    if (
      after.witness?.complete &&
      after.witness.sha256 === version.diskDigest &&
      after.witness.id !== before?.witness?.id
    )
      return true
    // Saving an already clean, unchanged document can be a no-op.
    return !!before && !before.version.dirty && sameTaskTarget(before.version, version)
  }
  return !before?.version.dirty || before.version.modelDigest === version.diskDigest
}

export function observedEffect(
  before: DocumentObservation[],
  after: DocumentObservation[],
): boolean {
  return (
    before.length !== after.length ||
    before.some((value) => {
      const next = after.find((item) => item.version.path === value.version.path)
      return (
        !next ||
        value.version.exists !== next.version.exists ||
        value.version.diskDigest !== next.version.diskDigest ||
        value.version.modelDigest !== next.version.modelDigest
      )
    })
  )
}

export async function taskFileChanges(
  before: DocumentObservation[],
  after: DocumentObservation[],
): Promise<AgentRunFileChange[]> {
  const changes: AgentRunFileChange[] = []
  for (const current of after) {
    const previous =
      before.find((item) => item.version.path === current.version.path) ??
      before.find(
        (item) =>
          !item.version.exists &&
          current.version.exists &&
          !!item.version.editorIdentity &&
          item.version.editorIdentity === current.version.editorIdentity,
      )
    if (previous && !observedEffect([previous], [current])) continue
    const a = previous?.version,
      b = current.version
    const live = !!b.editorIdentity
    const type = b.type ?? a?.type ?? 'unknown'
    const left =
      previous?.snapshot &&
      (await observeDocument(previous.snapshot, previous.liveSnapshot === true).catch(
        () => undefined,
      ))
    const right =
      current.snapshot &&
      (await observeDocument(current.snapshot, current.liveSnapshot === true).catch(
        () => undefined,
      ))
    const diff = documentChanges(left, right)
    const readable = b.complete && (!!right || (!b.exists && !live))
    const diskWritten = !!b.diskDigest && b.diskDigest !== a?.diskDigest
    const saved = (!live && !!b.diskDigest) || (live && diskWritten && b.dirty === false)
    const checkedPackage =
      saved &&
      !!b.diskDigest &&
      current.witness?.sha256 === b.diskDigest &&
      current.witness.id !== previous?.witness?.id
    const resources =
      diskWritten || operationForDisk(a, b)
        ? await observeResourceChanges(previous?.packageSnapshot, current.packageSnapshot)
        : undefined
    const operation =
      !a?.exists && b.exists ? 'created' : a?.exists && !b.exists && !live ? 'deleted' : 'modified'
    changes.push({
      path: b.path,
      documentPath: b.path,
      ...(b.editorIdentity ? { documentId: b.editorIdentity } : {}),
      ...(a && a.path !== b.path ? { previousPath: a.path } : {}),
      type,
      operation,
      tools: [],
      summaries: [],
      undoable: false,
      evidence: {
        version: 1,
        id: randomUUID(),
        recordedAt: new Date().toISOString(),
        type,
        source: live ? 'live_editor' : 'document',
        effect: 'changed',
        beforeVersion: a?.modelDigest ?? a?.diskDigest,
        afterVersion: b.modelDigest ?? b.diskDigest,
        diskVersion: b.diskDigest,
        saveState: saved
          ? 'saved'
          : live && !b.exists
            ? 'memory_only'
            : live
              ? 'pending_save'
              : 'unknown',
        verification:
          (checkedPackage && current.witness?.complete) || (readable && (!saved || !live))
            ? 'passed'
            : 'unknown',
        basis: !saved && live ? 'live_snapshot' : saved ? 'disk_readback' : 'unavailable',
        verificationScope:
          checkedPackage || (saved && !live)
            ? 'package_bytes'
            : !saved && live
              ? 'editor_model'
              : 'unknown',
        ...diff,
        resources: resources?.details,
        omitted: diff.omitted + (resources?.omitted ?? 0),
        incomplete: diff.incomplete || resources?.incomplete === true || !readable,
        ...(!readable && !checkedPackage
          ? { issue: 'observation_failed' as const }
          : saved && live && (!checkedPackage || !current.witness?.complete)
            ? { issue: 'save_unverified' as const }
            : {}),
      },
    })
  }
  return changes
}
function operationForDisk(
  before: TaskTargetVersion | undefined,
  after: TaskTargetVersion,
): boolean {
  return before?.exists !== after.exists
}

/** Merge only the same target or a proven first save of the same live document. */
export function mergeTaskFileChange(
  changes: Map<string, AgentRunFileChange>,
  value: AgentRunFileChange,
): void {
  const change = structuredClone(value)
  const previous =
    changes.get(change.path) ??
    (change.previousPath && change.documentId
      ? [...changes.values()].find(
          (item) => item.path === change.previousPath && item.documentId === change.documentId,
        )
      : undefined)
  if (previous) {
    if (previous.path !== change.path) changes.delete(previous.path)
    change.tools = [...new Set([...previous.tools, ...change.tools])]
    change.summaries = [...new Set([...previous.summaries, ...change.summaries])]
    if (previous.operation === 'created') change.operation = 'created'
    if (previous.evidence && change.evidence) {
      const a = previous.evidence,
        b = change.evidence
      b.beforeVersion = a.beforeVersion
      const details = [...a.details, ...b.details],
        resources = [...(a.resources ?? []), ...(b.resources ?? [])]
      b.details = details.slice(0, 80)
      b.resources = resources.slice(0, 80)
      b.omitted += a.omitted + Math.max(0, details.length - 80) + Math.max(0, resources.length - 80)
      b.incomplete ||= a.incomplete || b.omitted > 0
    }
  }
  changes.set(change.path, change)
}
