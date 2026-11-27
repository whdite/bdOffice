import { useEffect, useRef, useState } from 'react'
import type { AgentRunReview, RunReviewPage } from '../../shared/workspace-agent-api'
import type {
  ChangeLocation,
  FileChangeEvidence,
  LocatedChange,
} from '../../shared/change-evidence'
import { workspaceChatCopy, isChineseWorkspaceLanguage } from './workspace-chat-strings'
import './run-evidence.css'

function labels(lang: string) {
  return isChineseWorkspaceLanguage(lang)
    ? {
        result: '任务成果',
        notCreated: '未创建',
        noChange: '未发生变更',
        creationUnknown: '创建结果待核实',
        undoStates: {
          restored: '已恢复本轮前内容',
          removed: '已移除本轮新增文件',
          conflict: '内容已变化，未撤销',
          not_attempted: '未执行撤销',
          unknown: '撤销结果待核实',
        },
        openFile: '打开文件',
        partial: '部分完成',
        failed: '任务需要处理',
        cancelled: '已停止，保留以下结果',
        unchanged: '未检测到内容修改',
        unknown: '修改结果待核实',
        before: '修改前',
        after: '修改后',
        created: '新增内容',
        deleted: '已移除',
        format: '格式或布局变化',
        formatDetails: '格式与布局细节',
        file: '文件内容',
        block: '文档块',
        slide: '幻灯片',
        element: '元素',
        page: '页',
        line: '行',
        sheet: '工作表',
        memory_only: '仅在编辑器中',
        pending_save: '待保存',
        saved: '已保存',
        save_unknown: '保存状态待核实',
        live: '编辑器内容已核验',
        disk: '磁盘内容已核验',
        fileVerified: '文件已核验',
        unverified: '待核实',
        changed: '核验发现变化',
        recorded: '记录时间',
        limited: '仅展示部分差异，请结合原文件检查。',
        saveFailed: '保存未成功，请检查文件后继续保存。',
        observationFailed: '无法确认最终内容，请检查原文件。',
        saveUnverified: '保存后的内容尚未通过读回核对。',
        toolFailed: '执行中出现失败，以下仅展示实际观察到的结果。',
        unavailable: '暂无法提供精确位置差异，请检查原文件。',
        notStored: '成果记录未能独立保存，请保留当前结果。',
        conflict: '存在后续修改，未撤销',
        conflictHelp: '文件或编辑器内容已变化。已阻止旧快照覆盖，请先检查新修改。',
        history: '成果记录',
        more: '加载更早成果',
        loading: '正在读取…',
        empty: '这个对话还没有文件成果。',
        historyFailed: '无法读取成果记录，请稍后重试。',
        retry: '重试',
        legacy: '历史摘要，未记录逐项核验证据',
        advanced: '执行摘要',
        omitted: (count: number) => `另有 ${count} 处差异未展开`,
        excerpt: '内容较长，显示变化附近的摘录',
      }
    : {
        result: 'Task results',
        notCreated: 'Not created',
        noChange: 'No change',
        creationUnknown: 'Creation not verified',
        undoStates: {
          restored: 'Pre-run content restored',
          removed: 'New file removed',
          conflict: 'Content changed; not restored',
          not_attempted: 'Undo not attempted',
          unknown: 'Undo result needs verification',
        },
        openFile: 'Open file',
        partial: 'Partially completed',
        failed: 'Task needs attention',
        cancelled: 'Stopped; results retained',
        unchanged: 'No content change detected',
        unknown: 'Change needs verification',
        before: 'Before',
        after: 'After',
        created: 'Added content',
        deleted: 'Removed',
        format: 'Formatting or layout changed',
        formatDetails: 'Formatting and layout details',
        file: 'File content',
        block: 'Document block',
        slide: 'Slide',
        element: 'Element',
        page: 'Page',
        line: 'Line',
        sheet: 'Sheet',
        memory_only: 'In the editor only',
        pending_save: 'Pending save',
        saved: 'Saved',
        save_unknown: 'Save needs verification',
        live: 'Editor content verified',
        disk: 'Disk content verified',
        fileVerified: 'File verified',
        unverified: 'Needs verification',
        changed: 'Changes found during verification',
        recorded: 'Recorded',
        limited: 'Some differences are not displayed. Review the original file too.',
        saveFailed: 'Saving did not succeed. Review the file and save again.',
        observationFailed: 'The final content could not be confirmed. Review the original file.',
        saveUnverified: 'Saved content has not passed the readback check.',
        toolFailed: 'An operation failed. Only observed results are shown below.',
        unavailable: 'Precise location differences are unavailable. Review the original file.',
        notStored: 'The independent result record could not be saved. Keep this result.',
        conflict: 'Newer changes; undo blocked',
        conflictHelp:
          'The file or editor has changed. Restoring the old snapshot was blocked to preserve newer edits.',
        history: 'Result history',
        more: 'Load earlier results',
        loading: 'Loading…',
        empty: 'No file results in this conversation yet.',
        historyFailed: 'Result history could not be read. Try again later.',
        retry: 'Retry',
        legacy: 'Historical summary; no item-level verification recorded',
        advanced: 'Operation summaries',
        omitted: (count: number) => `${count} more differences not displayed`,
        excerpt: 'Excerpt around the change',
      }
}
function locationLabel(location: ChangeLocation, copy: ReturnType<typeof labels>): string {
  switch (location.kind) {
    case 'block':
      return `${copy.block} ${location.index}`
    case 'cell':
      return `${location.sheet}!${location.range}`
    case 'sheet':
      return `${copy.sheet} ${location.sheet}`
    case 'slide':
      return `${copy.slide} ${location.slide}${location.elementId ? ` · ${copy.element}` : ''}`
    case 'page':
      return `${copy.page} ${location.page}`
    case 'line':
      return `${copy.line} ${location.line}`
    case 'resource':
      return [
        location.page
          ? `${copy.page} ${location.page}`
          : location.slide
            ? `${copy.slide} ${location.slide}`
            : location.sheet,
        location.name,
      ]
        .filter(Boolean)
        .join(' · ')
    default:
      return copy.file
  }
}
function Evidence({
  value,
  lang,
  historical,
}: {
  value: FileChangeEvidence
  lang: string
  historical?: boolean
}) {
  const copy = labels(lang)
  const verified =
    value.verification === 'passed'
      ? historical
        ? isChineseWorkspaceLanguage(lang)
          ? '来源记录已核验'
          : 'Verified on the source device'
        : value.basis === 'disk_readback'
          ? copy.disk
          : value.basis === 'file_hash'
            ? copy.fileVerified
            : copy.live
      : value.verification === 'failed'
        ? copy.changed
        : copy.unverified
  const issue =
    value.issue === 'save_failed'
      ? copy.saveFailed
      : value.issue === 'save_unverified'
        ? copy.saveUnverified
        : value.issue === 'observation_failed'
          ? copy.observationFailed
          : value.issue === 'tool_failed'
            ? copy.toolFailed
            : undefined
  const difference = (detail: LocatedChange) => (
    <div className="run-evidence-change" key={detail.id}>
      <strong title={detail.location.kind === 'slide' ? detail.location.elementId : undefined}>
        {locationLabel(detail.location, copy)}
      </strong>
      {detail.kind === 'format' ? (
        <details>
          <summary>{copy.formatDetails}</summary>
          <div className="run-evidence-diff">
            <div>
              <span>{copy.before}</span>
              <pre>{detail.before}</pre>
            </div>
            <div>
              <span>{copy.after}</span>
              <pre>{detail.after}</pre>
            </div>
          </div>
        </details>
      ) : detail.before !== undefined || detail.after !== undefined ? (
        <div className="run-evidence-diff">
          <div>
            <span>{copy.before}</span>
            <pre>{detail.before || '∅'}</pre>
          </div>
          <div>
            <span>{copy.after}</span>
            <pre>{detail.after || '∅'}</pre>
          </div>
        </div>
      ) : (
        <p>{copy.unavailable}</p>
      )}
      {detail.truncated && <small>{copy.excerpt}</small>}
    </div>
  )
  return (
    <div className="run-evidence">
      <div
        className="run-evidence-status"
        title={`${copy.recorded}: ${new Date(value.recordedAt).toLocaleString()}`}
      >
        <span data-save-state={value.saveState}>
          {value.saveState === 'unknown' ? copy.save_unknown : copy[value.saveState]}
        </span>
        <span
          data-verification={value.verification}
          data-verification-scope={value.verificationScope}
        >
          {verified}
        </span>
        {value.effect !== 'changed' && (
          <span>{value.effect === 'unchanged' ? copy.unchanged : copy.unknown}</span>
        )}
        <time dateTime={value.recordedAt}>
          {copy.recorded}: {new Date(value.recordedAt).toLocaleTimeString()}
        </time>
      </div>
      {issue && <p className="run-evidence-issue">{issue}</p>}
      {value.details.map(difference)}
      {!!value.resources?.length && (
        <details className="run-evidence-resources">
          <summary>
            {isChineseWorkspaceLanguage(lang)
              ? '样式、图表、图像与页面资源'
              : 'Styles, charts, images and page resources'}
          </summary>
          {value.resources.map(difference)}
        </details>
      )}
      {!value.details.length && !value.resources?.length && value.effect !== 'unchanged' && (
        <p>{copy.unavailable}</p>
      )}
      {value.omitted > 0 && <p>{copy.omitted(value.omitted)}</p>}
      {value.incomplete && <p>{copy.limited}</p>}
    </div>
  )
}
interface Props {
  run: AgentRunReview
  lang: string
  undoing: boolean
  undoDisabled?: boolean
  onUndo(runId: string): void
  onOpen?(path: string, artifactId?: string): void
  onWordReview?(runId: string, path: string): void
  compact?: boolean
}
export function RunReviewCard({
  run,
  lang,
  undoing,
  undoDisabled,
  onUndo,
  onOpen,
  onWordReview,
  compact,
}: Props) {
  const copy = labels(lang),
    chat = workspaceChatCopy(lang)
  const [expanded, setExpanded] = useState(false)
  const zh = isChineseWorkspaceLanguage(lang)
  const showDetails = !compact || expanded
  const word = run.changes.length === 1 ? run.changes[0] : undefined
  const showSaveSummary = compact && !expanded && run.state === 'completed' && !!run.changes.length
  const allSaved = run.changes.every((change) => change.evidence?.saveState === 'saved')
  const allVerified =
    !run.delivery?.outstanding.length &&
    run.changes.every((change) => change.evidence?.verification === 'passed')
  const canReviewWord = (change: AgentRunReview['changes'][number]) =>
    !run.historical && change.type === 'docx' && change.operation !== 'deleted' && !!change.evidence
  const changed = run.changes.filter(
    (change) => !change.evidence || change.evidence.effect === 'changed',
  ).length
  const title =
    run.state === 'undone'
      ? chat.agentRunUndone
      : run.state === 'undo-failed'
        ? run.undoConflict
          ? copy.conflict
          : chat.agentRunRollbackAttention
        : run.outcome === 'partial'
          ? copy.partial
          : run.outcome === 'failed'
            ? copy.failed
            : run.outcome === 'cancelled'
              ? copy.cancelled
              : changed
                ? chat.updatedFiles(changed)
                : copy.result
  return (
    <div
      className={`workspace-chat-run-review ${run.state}`}
      data-outcome={run.outcome}
      data-result-origin={run.historical ? 'imported' : 'local'}
    >
      <div className="workspace-chat-run-head">
        <div>
          <strong>{title}</strong>
          <span title={run.changes.map((change) => change.path).join(', ')}>
            {run.changes.map((change) => change.path.split(/[\\/]/).pop()).join(', ')}
          </span>
          {showSaveSummary && (
            <small>
              {[
                allSaved ? copy.saved : copy.save_unknown,
                allVerified ? (zh ? '已核验' : 'Verified') : copy.unverified,
              ].join(' · ')}
            </small>
          )}
        </div>
        <div className="run-evidence-actions">
          {compact && word && onWordReview && canReviewWord(word) && (
            <button type="button" onClick={() => onWordReview(run.id, word.path)}>
              {zh ? '审阅修订' : 'Review revisions'}
            </button>
          )}
          {compact && (
            <button type="button" aria-expanded={expanded} onClick={() => setExpanded(!expanded)}>
              {expanded ? (zh ? '收起详情' : 'Hide details') : zh ? '查看详情' : 'View details'}
            </button>
          )}
          <button
            type="button"
            disabled={!run.undoable || undoDisabled || undoing}
            onClick={() => onUndo(run.id)}
          >
            {undoing ? chat.undoing : chat.undoRun}
          </button>
        </div>
      </div>
      {run.undoConflict && <p className="run-evidence-issue">{copy.conflictHelp}</p>}
      {run.historical && (
        <p>
          {isChineseWorkspaceLanguage(lang)
            ? '迁移的历史成果，核验状态来自源记录。'
            : 'Imported result history; verification describes the source record.'}
        </p>
      )}
      {run.evidenceStored === false && <p className="run-evidence-issue">{copy.notStored}</p>}
      {showDetails && run.delivery && (
        <div className="run-evidence-delivery">
          {run.delivery.checked.map((line) => (
            <div key={line}>{line}</div>
          ))}
          {run.delivery.outstanding.map((line) => (
            <div key={line}>
              {zh ? '待完成 · ' : 'Outstanding · '}
              {line}
            </div>
          ))}
        </div>
      )}
      {showDetails && (
        <div className="workspace-chat-run-files">
          {run.changes.map((change) => (
            <div className="workspace-chat-run-file" key={change.path}>
              <div>
                <span className="workspace-chat-run-operation">
                  {change.evidence?.effect === 'unchanged'
                    ? change.operation === 'created'
                      ? copy.notCreated
                      : copy.noChange
                    : change.operation === 'created' && change.evidence?.effect === 'unknown'
                      ? copy.creationUnknown
                      : chat.operation(change.operation)}
                </span>
                <strong title={change.path}>{change.path.split(/[\\/]/).pop()}</strong>
                {onWordReview && canReviewWord(change) && (!compact || !word) && (
                  <button
                    type="button"
                    className="run-evidence-open"
                    onClick={() => onWordReview(run.id, change.path)}
                  >
                    {zh ? '审阅修订' : 'Review revisions'}
                  </button>
                )}
                {onOpen &&
                  (!run.historical || change.artifactId) &&
                  change.operation !== 'deleted' &&
                  change.evidence?.saveState === 'saved' && (
                    <button
                      type="button"
                      className="run-evidence-open"
                      onClick={() => onOpen(change.path, change.artifactId)}
                    >
                      {copy.openFile}
                    </button>
                  )}
              </div>
              {change.undoState && (
                <p className="run-evidence-undo-state" data-undo-state={change.undoState}>
                  {copy.undoStates[change.undoState]}
                </p>
              )}
              {change.evidence?.beforeVersion && change.evidence.afterVersion && (
                <small
                  className="run-evidence-version"
                  title={change.evidence.beforeVersion + ' → ' + change.evidence.afterVersion}
                >
                  {zh ? '基线' : 'Base'} <code>{change.evidence.beforeVersion.slice(0, 12)}</code>
                  {' → '}
                  {zh ? '结果' : 'Result'} <code>{change.evidence.afterVersion.slice(0, 12)}</code>
                </small>
              )}
              {change.evidence ? (
                <Evidence value={change.evidence} lang={lang} historical={run.historical} />
              ) : (
                <p>{copy.legacy}</p>
              )}
              {change.summaries.length > 0 && (
                <details>
                  <summary>{copy.advanced}</summary>
                  {change.summaries.map((summary) => (
                    <p key={summary}>{summary}</p>
                  ))}
                </details>
              )}
            </div>
          ))}
        </div>
      )}
      {run.error && !run.undoConflict && (
        <div className="workspace-chat-run-error">{run.error}</div>
      )}
    </div>
  )
}

/** Optional history reads the independent records, not just the retained rollback list. */
export function RunEvidenceHistory({
  sessionId,
  lang,
  onUndo,
  undoingRunId,
  revision,
  onOpen,
  onWordReview,
  expanded,
}: {
  sessionId: string
  lang: string
  onUndo(runId: string): void
  undoingRunId: string | null
  revision: string
  onOpen?(path: string, artifactId?: string): void
  onWordReview?(runId: string, path: string): void
  expanded?: boolean
}) {
  const copy = labels(lang),
    [open, setOpen] = useState(false)
  const visible = expanded ?? open
  const [page, setPage] = useState<RunReviewPage>({ reviews: [] }),
    [busy, setBusy] = useState(false),
    [failed, setFailed] = useState(false)
  const generation = useRef({ cancelled: false }),
    loading = useRef(false)
  const [retry, setRetry] = useState(0)
  useEffect(() => {
    const epoch = { cancelled: false }
    generation.current = epoch
    if (!visible) return
    setBusy(true)
    setFailed(false)
    loading.current = true
    const api = window.aiOfficeWorkspaceAgent ?? window.workspaceAgent
    if (!api?.listRunReviews) {
      setBusy(false)
      loading.current = false
      setFailed(true)
      return
    }
    void api
      .listRunReviews(sessionId)
      .then((result) => {
        if (!epoch.cancelled) setPage(result)
      })
      .catch(() => {
        if (!epoch.cancelled) setFailed(true)
      })
      .finally(() => {
        if (!epoch.cancelled) {
          setBusy(false)
          loading.current = false
        }
      })
    return () => {
      epoch.cancelled = true
    }
  }, [sessionId, visible, revision, retry])
  const more = async () => {
    if (loading.current || !page.nextCursor) return
    const epoch = generation.current,
      api = window.aiOfficeWorkspaceAgent ?? window.workspaceAgent
    loading.current = true
    setBusy(true)
    try {
      const result = await api!.listRunReviews(sessionId, page.nextCursor)
      if (!epoch.cancelled)
        setPage((current) => ({
          reviews: [...current.reviews, ...result.reviews],
          nextCursor: result.nextCursor,
        }))
    } catch {
      if (!epoch.cancelled) setFailed(true)
    } finally {
      if (!epoch.cancelled) {
        loading.current = false
        setBusy(false)
      }
    }
  }
  const content = visible ? (
    <div>
      {page.reviews.map((run) => (
        <RunReviewCard
          key={run.id}
          run={run}
          lang={lang}
          onUndo={onUndo}
          onOpen={onOpen}
          onWordReview={onWordReview}
          undoing={undoingRunId === run.id}
          undoDisabled={undoingRunId !== null}
        />
      ))}
      {busy && <p role="status">{copy.loading}</p>}
      {!busy && !page.reviews.length && !failed && <p>{copy.empty}</p>}
      {failed && (
        <p role="alert">
          {copy.historyFailed}{' '}
          <button type="button" onClick={() => setRetry((value) => value + 1)}>
            {copy.retry}
          </button>
        </p>
      )}
      {page.nextCursor && (
        <button type="button" disabled={busy} onClick={() => void more()}>
          {copy.more}
        </button>
      )}
    </div>
  ) : null
  if (expanded !== undefined)
    return (
      <section className="run-evidence-history" aria-label={copy.history}>
        {content}
      </section>
    )
  return (
    <details
      className="run-evidence-history"
      onToggle={(event) => setOpen(event.currentTarget.open)}
    >
      <summary>{copy.history}</summary>
      {content}
    </details>
  )
}
