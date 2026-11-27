import type { WebContents } from 'electron'
import type {
  DocumentCommand,
  DocumentReadRequest,
  DocumentSnapshot,
  DocumentSummary,
  LiveDocumentAdapter,
  MutationResult,
} from './service'
import type {
  WorkspaceEditorContext,
  WorkspaceFileContent,
  WorkspaceFileType,
  WorkspaceSearchResult,
} from '../../shared/workspace-agent-api'
import { renderRenderSlide } from './pptx-raster'
import type { RenderSlide } from '@genoffice/pptx-render'
import { documentCommandError } from '../../shared/document-tool-contracts'
import { normalizePptxVisionRepairPatches, type PptxVisionRepairPatch } from './vision'

type JsonRecord = Record<string, unknown>

const MAX_CONTEXT_BLOCKS = 80
const MAX_CONTEXT_SLIDES = 12

function isRecord(value: unknown): value is JsonRecord {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function stringValue(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value : fallback
}

function integerValue(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isInteger(value) ? value : undefined
}

function jsonSnapshot(path: string, type: WorkspaceFileType, value: unknown): DocumentSnapshot {
  return {
    path,
    type,
    bytesBase64: Buffer.from(JSON.stringify(value), 'utf8').toString('base64'),
    createdAt: new Date().toISOString(),
  }
}

function snapshotJson(snapshot: DocumentSnapshot): unknown {
  return JSON.parse(Buffer.from(snapshot.bytesBase64, 'base64').toString('utf8')) as unknown
}

/** Execute a small, argument-only expression in an embedded editor renderer. */
async function executeRenderer<T>(
  webContents: WebContents,
  body: string,
  args: readonly unknown[] = [],
): Promise<T> {
  if (webContents.isDestroyed()) throw new Error('The editor renderer is no longer available.')
  const encodedArgs = JSON.stringify(args).replace(/</g, '\\u003c')
  const code = `(async () => { const __args = ${encodedArgs}; ${body} })()`
  return (await webContents.executeJavaScript(code, true)) as T
}

async function waitForRenderer(webContents: WebContents, body: string): Promise<void> {
  // A Shell tab can be returned before its renderer has mounted the editor.
  // Keep the Agent call deterministic during that short startup window.
  for (let attempt = 0; attempt < 30; attempt += 1) {
    // Retry readiness only. A rejected mutation may already have written data.
    if (await executeRenderer<boolean>(webContents, body)) return
    if (attempt < 29) await new Promise((resolve) => setTimeout(resolve, 40))
  }
  throw new Error('The editor workspace bridge is not ready.')
}

async function bridgeCall<T>(
  webContents: WebContents,
  method: string,
  args: readonly unknown[] = [],
): Promise<T> {
  const methodLiteral = JSON.stringify(method)
  await waitForRenderer(
    webContents,
    `return typeof window.__genofficeWorkspaceAgent?.[${methodLiteral}] === 'function';`,
  )
  return executeRenderer<T>(
    webContents,
    `const bridge = window.__genofficeWorkspaceAgent;
     if (!bridge || typeof bridge[${methodLiteral}] !== 'function') throw new Error('The editor workspace bridge is not ready.');
     return await bridge[${methodLiteral}](...__args);`,
    args,
  )
}

async function slidesCall<T>(
  webContents: WebContents,
  method: string,
  args: readonly unknown[] = [],
): Promise<T> {
  const methodLiteral = JSON.stringify(method)
  await waitForRenderer(
    webContents,
    `return typeof window.slidesApi?.[${methodLiteral}] === 'function';`,
  )
  return executeRenderer<T>(
    webContents,
    `const api = window.slidesApi;
     if (!api || typeof api[${methodLiteral}] !== 'function') throw new Error('The Slides editor is not ready.');
     return await api[${methodLiteral}](...__args);`,
    args,
  )
}

function renderNodeText(value: unknown): string {
  if (!isRecord(value)) return ''
  const text = value.text
  if (isRecord(text) && Array.isArray(text.lines)) {
    return text.lines
      .map((line) => {
        if (!isRecord(line) || !Array.isArray(line.runs)) return ''
        return line.runs
          .filter(isRecord)
          .filter((run) => run.isBullet !== true)
          .map((run) => stringValue(run.text))
          .join('')
      })
      .filter(Boolean)
      .join('\n')
  }
  if (value.type === 'group' && Array.isArray(value.children)) {
    return value.children.map(renderNodeText).filter(Boolean).join('\n')
  }
  if (value.type === 'table' && Array.isArray(value.cells)) {
    return value.cells.map(renderNodeText).filter(Boolean).join(' | ')
  }
  if (value.type === 'chart' && Array.isArray(value.labels)) {
    return value.labels
      .filter(isRecord)
      .map((label) => stringValue(label.text))
      .filter(Boolean)
      .join('\n')
  }
  return ''
}

function slideTexts(value: unknown): string[] {
  if (!Array.isArray(value)) return []
  return value.map((slide) =>
    isRecord(slide) && Array.isArray(slide.nodes)
      ? slide.nodes.map(renderNodeText).filter(Boolean).join('\n')
      : '',
  )
}

function countRenderNodes(value: unknown): number {
  if (!isRecord(value)) return 0
  const children = Array.isArray(value.children) ? value.children : []
  return 1 + children.reduce((count, child) => count + countRenderNodes(child), 0)
}

function summarizeRenderNodes(
  value: unknown,
  slide: number,
  output: NonNullable<DocumentSummary['elements']>,
): void {
  if (!isRecord(value) || output.length >= 500) return
  const box = isRecord(value.box) ? value.box : undefined
  const id = stringValue(
    value.id,
    stringValue(value.sourceId, `slide-${slide}-${output.length + 1}`),
  )
  const text = renderNodeText(value)
  const item: NonNullable<DocumentSummary['elements']>[number] = {
    id,
    type: stringValue(value.type, 'unknown'),
    slide,
    ...(box && typeof box.x === 'number' ? { x: box.x } : {}),
    ...(box && typeof box.y === 'number' ? { y: box.y } : {}),
    ...(box && typeof box.w === 'number' ? { width: box.w } : {}),
    ...(box && typeof box.h === 'number' ? { height: box.h } : {}),
    ...(text ? { text: text.slice(0, 500) } : {}),
  }
  if (Array.isArray(value.children)) item.children = value.children.length
  output.push(item)
  if (Array.isArray(value.children))
    for (const child of value.children) summarizeRenderNodes(child, slide, output)
}

function maxCharsOf(request: DocumentReadRequest | undefined): number {
  const raw = request?.maxChars
  return Math.min(Math.max(integerValue(raw) ?? 80_000, 1), 500_000)
}

class SlidesLiveAdapter implements LiveDocumentAdapter {
  readonly type = 'pptx' as const

  constructor(private readonly webContents: WebContents) {}

  async getContext(): Promise<WorkspaceEditorContext> {
    return (await bridgeCall<WorkspaceEditorContext>(this.webContents, 'getContext')) ?? {}
  }

  private async slides(): Promise<unknown[]> {
    const value = await slidesCall<unknown>(this.webContents, 'getRenderSlides')
    return Array.isArray(value) ? value : []
  }

  async inspect(): Promise<DocumentSummary> {
    const slides = await this.slides()
    const texts = slideTexts(slides)
    const elements: NonNullable<DocumentSummary['elements']> = []
    slides.forEach((slide, index) => {
      if (!isRecord(slide) || !Array.isArray(slide.nodes)) return
      for (const node of slide.nodes) summarizeRenderNodes(node, index + 1, elements)
    })
    const elementCount = slides.reduce<number>(
      (total, slide) =>
        total +
        (isRecord(slide) && Array.isArray(slide.nodes)
          ? slide.nodes.reduce((count, node) => count + countRenderNodes(node), 0)
          : 0),
      0,
    )
    return {
      path: '',
      type: this.type,
      title: texts[0]?.split(/\r?\n/, 1)[0] ?? '',
      textLength: texts.reduce((total, text) => total + text.length, 0),
      slideCount: texts.length,
      elementCount,
      slides: texts.map((text, index) => ({
        index: index + 1,
        title: text.split(/\r?\n/, 1)[0] ?? '',
        text: text.slice(0, 1_000),
      })),
      ...(elements.length ? { elements } : {}),
    }
  }

  async search(query: string): Promise<WorkspaceSearchResult[]> {
    const needle = String(query ?? '')
      .trim()
      .toLocaleLowerCase()
    if (!needle) return []
    const texts = slideTexts(await this.slides())
    return texts
      .map((text, index) => ({ text, index }))
      .filter(({ text }) => text.toLocaleLowerCase().includes(needle))
      .slice(0, 100)
      .map(({ text, index }) => ({
        path: '',
        file: '',
        type: this.type,
        preview: text.slice(0, 500),
        location: { slide: index + 1 },
      }))
  }

  async read(request: DocumentReadRequest = {}): Promise<WorkspaceFileContent> {
    const texts = slideTexts(await this.slides())
    const requestedIndices = request.slides?.length
      ? request.slides
      : texts.map((_text, index) => index + 1)
    if (requestedIndices.some((page) => !Number.isInteger(page) || page < 1 || page > texts.length))
      throw new Error('Every requested slide must exist (1-based).')
    const selective = (request.slides?.length ?? 0) > 0
    const indices = selective ? requestedIndices : requestedIndices.slice(0, MAX_CONTEXT_SLIDES)
    const text = indices
      .map((index) => {
        const value = texts[index - 1]
        return value ? `Slide ${index}\n${value}` : ''
      })
      .filter(Boolean)
      .join('\n\n')
    const maxChars = maxCharsOf(request)
    return {
      path: '',
      type: this.type,
      text: text.slice(0, maxChars),
      truncated: text.length > maxChars || indices.length < requestedIndices.length,
    }
  }

  async mutate(command: DocumentCommand): Promise<MutationResult> {
    const invalid = documentCommandError(this.type, command)
    if (invalid) return { ok: false, changed: 0, description: invalid }
    const kind = stringValue(command.kind, stringValue(command.operation, 'replace_text'))
    if (kind === 'regenerate_slide') {
      return {
        ok: false,
        changed: 0,
        description: 'Live Slides sessions do not support slide regeneration yet.',
      }
    }
    const find = stringValue(command.find, stringValue(command.query))
    if (!find) return { ok: false, changed: 0, description: 'Provide find/query text.' }
    const replace = stringValue(
      command.replace,
      stringValue(command.text, stringValue(command.content)),
    )
    if (find === replace && command.matchCase !== false)
      return { ok: true, changed: 0, description: 'Text is already unchanged.' }
    // Workspace Agent document APIs expose user-facing page numbers (1-based),
    // while the Slides renderer IPC contract uses its internal zero-based
    // `slideIndex`. Keep the conversion at this adapter boundary so headless
    // and live sessions share the same semantic command shape.
    const slide = integerValue(command.slide) ?? integerValue(command.slideIndex)
    const slides = Array.isArray(command.slides)
      ? [...new Set(command.slides as number[])]
      : slide !== undefined
        ? [slide]
        : undefined
    const slideCount = slides ? (await this.slides()).length : 0
    if (slides && slides.some((page) => page > slideCount)) {
      return { ok: false, changed: 0, description: 'The requested slide does not exist.' }
    }
    let count = 0
    for (const selected of slides ?? [undefined]) {
      const result = await slidesCall<unknown>(this.webContents, 'findReplace', [
        {
          find,
          replace,
          ...(selected === undefined ? {} : { slideIndex: selected - 1 }),
          matchCase: command.matchCase !== false,
        },
      ])
      count += isRecord(result) && typeof result.count === 'number' ? result.count : 0
    }
    return count > 0
      ? { ok: true, changed: count, description: `Updated ${count} slide text run(s).` }
      : { ok: false, changed: 0, description: 'No matching slide text.' }
  }

  async save(): Promise<void> {
    const result = await slidesCall<unknown>(this.webContents, 'save')
    if (!isRecord(result) || result.ok !== true)
      throw new Error(
        stringValue(
          isRecord(result) ? result.error : undefined,
          'The Slides document could not be saved.',
        ),
      )
  }

  async snapshot(): Promise<DocumentSnapshot> {
    return jsonSnapshot('', this.type, await this.slides())
  }

  async undo(): Promise<void> {
    await slidesCall(this.webContents, 'undo')
  }

  async rollback(): Promise<void> {
    await this.undo()
  }

  async beginAgentRun(): Promise<void> {
    const opened = await slidesCall<boolean>(this.webContents, 'beginHistoryBatch')
    if (!opened) throw new Error('The Slides history transaction could not be opened.')
  }

  async endAgentRun(): Promise<unknown> {
    return slidesCall<number | null>(this.webContents, 'endHistoryBatch')
  }

  async restoreAgentRun(
    _snapshot: DocumentSnapshot,
    token: unknown,
    mutationCount: number,
  ): Promise<void> {
    if (typeof token === 'number') {
      const restored = await bridgeCall<unknown>(this.webContents, 'restoreAgentRun', [token])
      if (restored !== true && (!isRecord(restored) || restored.ok !== true))
        throw new Error('The Slides rollback point is no longer available.')
      return
    }
    for (let index = 0; index < Math.max(1, mutationCount); index += 1)
      await slidesCall(this.webContents, 'undo')
  }

  async renderSlide(slide: number) {
    const slides = await this.slides()
    const value = slides[slide - 1]
    if (!isRecord(value) || !Array.isArray(value.nodes))
      throw new Error(`Slide ${slide} does not exist.`)
    return renderRenderSlide(value as unknown as RenderSlide, slide)
  }

  async applyVisionRepair(
    slide: number,
    patches: readonly PptxVisionRepairPatch[],
  ): Promise<MutationResult> {
    if (!Number.isInteger(slide) || slide < 1)
      return { ok: false, changed: 0, description: 'The requested slide does not exist.' }
    const slides = await this.slides()
    const value = slides[slide - 1]
    if (!isRecord(value) || !Array.isArray(value.nodes))
      return { ok: false, changed: 0, description: `Slide ${slide} does not exist.` }
    const render = await renderRenderSlide(value as unknown as RenderSlide, slide)
    const elements = render.elements ?? []
    const allowed = new Set(elements.map((element) => element.id))
    const normalized = normalizePptxVisionRepairPatches(patches, allowed)
    if (!normalized.length)
      return { ok: false, changed: 0, description: 'No safe Vision repair patches.' }
    const byId = new Map(elements.map((element) => [element.id, element]))
    const geometry = normalized.flatMap((patch) => {
      if (
        patch.xRatio === undefined &&
        patch.yRatio === undefined &&
        patch.widthRatio === undefined &&
        patch.heightRatio === undefined
      )
        return []
      const current = byId.get(patch.elementId)
      if (!current) return []
      const xRatio = patch.xRatio ?? current.xRatio
      const yRatio = patch.yRatio ?? current.yRatio
      const widthRatio = patch.widthRatio ?? current.widthRatio
      const heightRatio = patch.heightRatio ?? current.heightRatio
      if (xRatio + widthRatio > 1 || yRatio + heightRatio > 1)
        throw new Error(`Vision repair would move ${patch.elementId} outside the slide.`)
      return [
        {
          sourceId: patch.elementId,
          xPx: xRatio * render.width,
          yPx: yRatio * render.height,
          wPx: widthRatio * render.width,
          hPx: heightRatio * render.height,
          rotationDeg: current.rotationDeg ?? 0,
        },
      ]
    })
    const fontPatches = normalized.filter((patch) => patch.fontSizePt !== undefined)
    for (const patch of fontPatches) {
      const current = byId.get(patch.elementId)
      if (!current || !['text', 'shape', 'table'].includes(current.type)) {
        throw new Error(
          `Vision font repair requires a text, shape, or table element: ${patch.elementId}.`,
        )
      }
    }
    const opened = await slidesCall<boolean>(this.webContents, 'beginHistoryBatch')
    if (!opened) throw new Error('The Slides history transaction could not be opened.')
    let changed = 0
    try {
      if (geometry.length) {
        const result = await slidesCall<unknown>(this.webContents, 'batchEditTransform', [
          {
            slideIndex: slide - 1,
            fitWidthPx: render.width,
            items: geometry,
          },
        ])
        if (!result) throw new Error('The live Slides geometry repair could not be applied.')
        changed += geometry.length
      }
      for (const patch of fontPatches) {
        const result = await slidesCall<unknown>(this.webContents, 'setElementFont', [
          {
            slideIndex: slide - 1,
            sourceIds: [patch.elementId],
            fontSizePt: patch.fontSizePt,
          },
        ])
        if (!result)
          throw new Error(`The live Slides font repair could not be applied to ${patch.elementId}.`)
        changed += 1
      }
      await slidesCall(this.webContents, 'endHistoryBatch')
    } catch (error) {
      // End the batch before undoing so the renderer does not leave a dangling
      // history transaction if a later patch fails.
      await slidesCall(this.webContents, 'endHistoryBatch').catch(() => undefined)
      throw error
    }
    return { ok: changed > 0, changed, description: `Applied ${changed} Vision repair change(s).` }
  }
}

class DocsLiveAdapter implements LiveDocumentAdapter {
  async exportBytes(): Promise<Uint8Array> {
    return Buffer.from(await bridgeCall<string>(this.webContents, 'currentWordPackage'), 'base64')
  }
  async cancelWord(id: string): Promise<void> {
    await bridgeCall(this.webContents, 'cancelWord', [id])
  }
  readonly type = 'docx' as const

  constructor(private readonly webContents: WebContents) {}

  async getContext(): Promise<WorkspaceEditorContext> {
    return (await bridgeCall<WorkspaceEditorContext>(this.webContents, 'getContext')) ?? {}
  }

  async taskDirty(): Promise<boolean> {
    return bridgeCall(this.webContents, 'taskDirty')
  }

  private async blocks(): Promise<
    Array<{ id: string; index: number; text: string; type: string }>
  > {
    const raw = await bridgeCall<unknown>(this.webContents, 'inspect')
    if (!isRecord(raw) || !Array.isArray(raw.blocks)) return []
    return raw.blocks.filter(isRecord).map((block, index) => ({
      id: stringValue(block.id, `block-${index}`),
      index,
      text: stringValue(block.text),
      type: stringValue(block.type, 'docParagraph'),
    }))
  }

  async inspect(): Promise<DocumentSummary> {
    const raw = await bridgeCall<unknown>(this.webContents, 'inspect')
    if (!isRecord(raw)) return { path: '', type: this.type, textLength: 0, blocks: [] }
    const rawBlocks = Array.isArray(raw.blocks) ? raw.blocks : []
    const blocks = rawBlocks.filter(isRecord).map((block, index) => ({
      id: stringValue(block.id, `block-${index}`),
      type: stringValue(block.type, 'paragraph'),
      text: stringValue(block.text).slice(0, 500),
      ...(isRecord(block.formatting) ? { formatting: block.formatting } : {}),
    }))
    const textLength =
      typeof raw.textLength === 'number' && Number.isFinite(raw.textLength)
        ? Math.max(0, Math.trunc(raw.textLength))
        : blocks.reduce((total, block) => total + block.text.length, 0)
    const toc = blocks
      .filter((block) => /heading/i.test(block.type))
      .slice(0, 200)
      .map((block) => ({ id: block.id, title: block.text }))
    return {
      path: '',
      type: this.type,
      title: stringValue(raw.title, blocks.find((block) => /heading/i.test(block.type))?.text),
      textLength,
      pageEstimate: Math.max(1, Math.ceil(Math.max(1, textLength) / 1_800)),
      ...(toc.length ? { toc } : {}),
      blocks: blocks.slice(0, 200),
      ...(typeof raw.sections === 'number'
        ? { sections: Math.max(1, Math.trunc(raw.sections)) }
        : {}),
      ...(typeof raw.comments === 'number'
        ? { comments: Math.max(0, Math.trunc(raw.comments)) }
        : {}),
      ...(typeof raw.revisions === 'boolean' ? { revisions: raw.revisions } : {}),
      ...(isRecord(raw.theme) ? { theme: raw.theme as DocumentSummary['theme'] } : {}),
    }
  }

  async search(query: string): Promise<WorkspaceSearchResult[]> {
    return (await bridgeCall<WorkspaceSearchResult[]>(this.webContents, 'search', [query])) ?? []
  }

  async read(request: DocumentReadRequest = {}): Promise<WorkspaceFileContent> {
    const selective =
      (request.blockIds?.length ?? 0) > 0 ||
      request.blockStart !== undefined ||
      request.blockEnd !== undefined
    const bridgeRequest = selective ? request : { ...request, maxBlocks: MAX_CONTEXT_BLOCKS }
    return (
      (await bridgeCall<WorkspaceFileContent>(this.webContents, 'read', [bridgeRequest])) ?? {
        path: '',
        type: this.type,
        text: '',
      }
    )
  }

  async mutate(command: DocumentCommand): Promise<MutationResult> {
    const invalid = documentCommandError(this.type, command)
    if (invalid) return { ok: false, changed: 0, description: invalid }
    return (
      (await bridgeCall<MutationResult>(this.webContents, 'mutate', [command])) ?? {
        ok: false,
        changed: 0,
        description: 'The Docs editor did not return a mutation result.',
      }
    )
  }

  async save(): Promise<void> {
    const result = await bridgeCall<unknown>(this.webContents, 'save')
    if (result !== true && (!isRecord(result) || result.ok !== true))
      throw new Error(
        stringValue(
          isRecord(result) ? result.error : undefined,
          'The Docs document could not be saved.',
        ),
      )
  }

  async snapshot(): Promise<DocumentSnapshot> {
    return jsonSnapshot('', this.type, await bridgeCall(this.webContents, 'snapshot'))
  }

  async undo(): Promise<void> {
    await bridgeCall(this.webContents, 'undo')
  }

  async rollback(): Promise<void> {
    await this.undo()
  }

  async restoreAgentRun(
    snapshot: DocumentSnapshot,
    _token?: unknown,
    _mutationCount?: number,
    expected?: DocumentSnapshot,
  ): Promise<void> {
    const restored = await bridgeCall<unknown>(this.webContents, 'restoreAgentRun', [
      snapshotJson(snapshot),
      ...(expected ? [snapshotJson(expected)] : []),
    ])
    if (restored !== true && (!isRecord(restored) || restored.ok !== true))
      throw new Error('The Docs rollback snapshot could not be restored.')
  }

  async review(command: DocumentCommand): Promise<unknown> {
    return bridgeCall(this.webContents, 'review', [command])
  }
}

class SheetsLiveAdapter implements LiveDocumentAdapter {
  constructor(
    private readonly webContents: WebContents,
    readonly type: 'xlsx' | 'xlsm' = 'xlsx',
  ) {}

  async getContext(): Promise<WorkspaceEditorContext> {
    return (await bridgeCall<WorkspaceEditorContext>(this.webContents, 'getContext')) ?? {}
  }

  async inspect(): Promise<DocumentSummary> {
    return (
      (await bridgeCall<DocumentSummary>(this.webContents, 'inspect')) ?? {
        path: '',
        type: this.type,
        textLength: 0,
        sheets: [],
      }
    )
  }

  async search(query: string): Promise<WorkspaceSearchResult[]> {
    return (await bridgeCall<WorkspaceSearchResult[]>(this.webContents, 'search', [query])) ?? []
  }

  async read(request: DocumentReadRequest = {}): Promise<WorkspaceFileContent> {
    return (
      (await bridgeCall<WorkspaceFileContent>(this.webContents, 'read', [request])) ?? {
        path: '',
        type: this.type,
        text: '',
      }
    )
  }

  async mutate(command: DocumentCommand): Promise<MutationResult> {
    const invalid = documentCommandError(this.type, command)
    if (invalid) return { ok: false, changed: 0, description: invalid }
    return (
      (await bridgeCall<MutationResult>(this.webContents, 'mutate', [command])) ?? {
        ok: false,
        changed: 0,
        description: 'The Sheets editor did not return a mutation result.',
      }
    )
  }

  async save(): Promise<void> {
    const result = await bridgeCall<unknown>(this.webContents, 'save')
    if (result !== true && (!isRecord(result) || result.ok !== true))
      throw new Error(
        stringValue(
          isRecord(result) ? result.error : undefined,
          'The Sheets workbook could not be saved.',
        ),
      )
  }

  async snapshot(): Promise<DocumentSnapshot> {
    return jsonSnapshot('', this.type, await bridgeCall(this.webContents, 'snapshot'))
  }

  async undo(): Promise<void> {
    await bridgeCall(this.webContents, 'undo')
  }

  async rollback(): Promise<void> {
    await this.undo()
  }

  async restoreAgentRun(
    _snapshot: DocumentSnapshot,
    _token: unknown,
    mutationCount: number,
  ): Promise<void> {
    const restored = await bridgeCall<unknown>(this.webContents, 'restoreAgentRun', [
      Math.max(1, mutationCount),
    ])
    if (restored !== true && (!isRecord(restored) || restored.ok !== true))
      throw new Error('The Sheets rollback steps could not be applied.')
  }
}

class RendererBridgeLiveAdapter implements LiveDocumentAdapter {
  constructor(
    readonly type: 'pdf' | 'md' | 'html',
    protected readonly webContents: WebContents,
  ) {}

  async getContext(): Promise<WorkspaceEditorContext> {
    return (await bridgeCall<WorkspaceEditorContext>(this.webContents, 'getContext')) ?? {}
  }

  async inspect(): Promise<DocumentSummary> {
    return (
      (await bridgeCall<DocumentSummary>(this.webContents, 'inspect')) ?? {
        path: '',
        type: this.type,
        textLength: 0,
      }
    )
  }

  async search(query: string): Promise<WorkspaceSearchResult[]> {
    return (await bridgeCall<WorkspaceSearchResult[]>(this.webContents, 'search', [query])) ?? []
  }

  async read(request: DocumentReadRequest = {}): Promise<WorkspaceFileContent> {
    return (
      (await bridgeCall<WorkspaceFileContent>(this.webContents, 'read', [request])) ?? {
        path: '',
        type: this.type,
        text: '',
      }
    )
  }

  async mutate(command: DocumentCommand): Promise<MutationResult> {
    const invalid = documentCommandError(this.type, command)
    if (invalid) return { ok: false, changed: 0, description: invalid }
    return (
      (await bridgeCall<MutationResult>(this.webContents, 'mutate', [command])) ?? {
        ok: false,
        changed: 0,
        description: `The ${this.type.toUpperCase()} editor did not return a mutation result.`,
      }
    )
  }

  async save(): Promise<void> {
    const result = await bridgeCall<unknown>(this.webContents, 'save')
    if (result !== true && (!isRecord(result) || result.ok !== true)) {
      throw new Error(
        stringValue(
          isRecord(result) ? result.error : undefined,
          `The ${this.type.toUpperCase()} document could not be saved.`,
        ),
      )
    }
  }

  async snapshot(): Promise<DocumentSnapshot> {
    return jsonSnapshot('', this.type, await bridgeCall(this.webContents, 'snapshot'))
  }

  async undo(): Promise<void> {
    await bridgeCall(this.webContents, 'undo')
  }

  async rollback(): Promise<void> {
    await bridgeCall(this.webContents, 'rollback')
  }

  async restoreAgentRun(snapshot: DocumentSnapshot): Promise<void> {
    const restored = await bridgeCall<unknown>(this.webContents, 'restoreAgentRun', [
      snapshotJson(snapshot),
    ])
    if (restored !== true && (!isRecord(restored) || restored.ok !== true)) {
      throw new Error(`The ${this.type.toUpperCase()} rollback snapshot could not be restored.`)
    }
  }
}

class PdfLiveAdapter extends RendererBridgeLiveAdapter {
  constructor(webContents: WebContents) {
    super('pdf', webContents)
  }

  async exportBytes(): Promise<Uint8Array> {
    const base64 = await bridgeCall<string>(this.webContents, 'materializeBase64')
    if (typeof base64 !== 'string' || !base64)
      throw new Error('The live PDF materialization failed.')
    return new Uint8Array(Buffer.from(base64, 'base64'))
  }
}

export function createLiveDocumentAdapter(
  type: WorkspaceFileType,
  webContents: WebContents,
): LiveDocumentAdapter | null {
  let adapter: LiveDocumentAdapter | null = null
  if (type === 'docx') adapter = new DocsLiveAdapter(webContents)
  else if (type === 'pptx') adapter = new SlidesLiveAdapter(webContents)
  else if (type === 'xlsx' || type === 'xlsm') adapter = new SheetsLiveAdapter(webContents, type)
  else if (type === 'pdf') adapter = new PdfLiveAdapter(webContents)
  else if (type === 'md') adapter = new RendererBridgeLiveAdapter('md', webContents)
  else if (type === 'html') adapter = new RendererBridgeLiveAdapter('html', webContents)
  if (adapter) {
    Object.defineProperty(adapter, 'instanceId', { value: webContents.id })
    adapter.taskDirty = () =>
      executeRenderer<boolean | undefined>(
        webContents,
        'return typeof window.__genofficeWorkspaceAgent?.taskDirty === "function" ? window.__genofficeWorkspaceAgent.taskDirty() : undefined;',
      )
    if (type !== 'pptx' && type !== 'html')
      adapter.mutateGuarded = async (command, expected) => {
        const invalid = documentCommandError(type, command)
        if (invalid) throw new Error(invalid)
        const expectedJson = Buffer.from(expected.bytesBase64, 'base64').toString('utf8')
        const guarded = await executeRenderer<{ result: MutationResult; after: unknown }>(
          webContents,
          `
        const bridge = window.__genofficeWorkspaceAgent;
        if (!bridge?.snapshot || !bridge?.mutate) throw new Error('The editor bridge is unavailable.');
        const before = bridge.snapshot();
        if (before && typeof before.then === 'function') throw new Error('Synchronous document revision checking is unavailable.');
        if (JSON.stringify(before) !== __args[1]) throw new Error('The document changed before the edit. Refresh its context.');
        const result = await bridge.mutate(__args[0]);
        return { result, after: await bridge.snapshot() };
      `,
          [command, expectedJson],
        )
        return {
          result: guarded.result,
          snapshot: jsonSnapshot(expected.path, type, guarded.after),
        }
      }
  }
  return adapter
}
