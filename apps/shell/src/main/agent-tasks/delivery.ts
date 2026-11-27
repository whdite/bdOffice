import { basename } from 'node:path'
import { parseAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { WordDelivery, type WordBaseline } from '../workspace-agent/word-delivery'
import { deliveryInstruction, wordBriefForFile, type AgentToolCall } from '@genoffice/agent-core'
import type { TaskTargetVersion } from '../../shared/assistant-task-api'
import { observeDeliveryDocument as observeDocument } from '../workspace-agent/change-evidence'
import type { DocumentObservation } from './document-observation'

interface Output {
  target: TaskTargetVersion
  created?: boolean
  before?: WordBaseline
}
type TextFact = {
  page?: number
  blockId?: string
  label: string
  value: string
  outputs?: string[]
}
export interface DeliveryState {
  outputs: Output[]
  sources: Array<{
    target: TaskTargetVersion
    sheet: string
    facts?: Array<{ cell: string; label?: string; outputs?: string[] }>
  }>
  textSources?: Array<{ target: TaskTargetVersion; facts: TextFact[] }>
}
const deliverable = /\.(?:docx|xlsx|xlsm|pptx|md|html?|pdf)$/i
function explicitSourceCells(goal: string): string[] {
  const withoutRanges = goal.replace(/\$?[A-Z]{1,3}\$?[1-9]\d*\s*:\s*\$?[A-Z]{1,3}\$?[1-9]\d*/g, '')
  const cells = new Set<string>()
  for (const match of withoutRanges.matchAll(
    /(?<![A-Za-z0-9])\$?([A-Z]{1,3})\$?([1-9]\d{0,5})(?![A-Za-z0-9])/g,
  )) {
    const cell = match[1]! + match[2]!
    if (
      /^(?:FY\d{2,4}|Q[1-4])$/.test(cell) &&
      !/(?:单元格|cell|!)\s*$/i.test(
        withoutRanges.slice(Math.max(0, match.index! - 10), match.index),
      )
    )
      continue
    cells.add(cell)
  }
  return [...cells]
}
function outputNames(goal: string): string[] {
  const at =
    /(?:生成(?:并保存)?|新建|输出(?:到|为)|交付为|写入|同步(?:到|至)|generate|create|deliver to|write (?:to|into))/i.exec(
      goal,
    )
  if (!at) return []
  const work =
    goal
      .slice(at.index + at[0].length)
      .split(/(?:基于|依据|根据|参考|参照|\bbased on\b|\busing\b|\bfrom\b)/i)[0] ?? ''
  const names = new Set<string>()
  const remaining = work
    .replace(
      /[“"'「`]([^“”"'「」`\n]{1,260}\.(?:docx|xlsx|pptx|md|html?|pdf))[”"'」`]/gi,
      (_, file: string) => {
        names.add(file.split(/[\\/]/).at(-1)!)
        return ' '
      },
    )
    .replace(/(\.(?:docx|xlsx|pptx|md|html?|pdf))\s*(?:以及|和|与|及|、)\s*/gi, '$1 ')
  const all = [
    ...remaining.matchAll(
      /(?:^|[\\/“"'「`\s])([^\\/"'“”「」`\s，。；、]{1,130}\.(?:docx|xlsx|pptx|md|html?|pdf))(?=$|[\s”"'」`，,。；、])/gi,
    ),
  ]
  for (const match of all) names.add(match[1]!)
  return [...names]
}
async function readView(doc: DocumentObservation) {
  return doc.snapshot && observeDocument(doc.snapshot, !!doc.liveSnapshot)
}
function figureMatches(sentence: string, value: string, header: string): boolean {
  const numeric = value.replace(/,/g, '')
  if (!/^-?\d+(?:\.\d+)?$/.test(numeric)) return sentence.includes(value)
  const currency = /(?:美元|人民币|元|USD|CNY)/i.test(header)
  const scale = (unit: string) =>
    /亿/.test(unit) ? 1e8 : /万/.test(unit) ? 1e4 : /千/.test(unit) ? 1e3 : 1
  const expected = Number(numeric) * (currency ? scale(header) : 1)
  for (const m of sentence
    .normalize('NFKC')
    .matchAll(
      /([-−]?\d+(?:,\d{3})*(?:\.\d+)?)(?:\s*((?:万|千|亿)?(?:美元|人民币|元)|USD|CNY|%))?/gi,
    )) {
    const unit = m[2] ?? ''
    if (/美元|USD/i.test(header) && /人民币|CNY/i.test(unit)) continue
    if (/人民币|CNY/i.test(header) && /美元|USD/i.test(unit)) continue
    const actual =
      Number(m[1]!.replace(/,/g, '').replace('−', '-')) * (currency ? scale(unit || header) : 1)
    if (Math.abs(actual - expected) < 1e-8) return true
  }
  return false
}

/** One durable, compact baseline per affected file. Reads use the originating host's file binding. */
export class DeliveryReview {
  observations: string[] = []
  readonly state: DeliveryState
  private readonly goal: string
  private readonly zh: boolean
  private readonly word: WordDelivery
  private readonly cells: string[]
  private readonly reconcile: boolean
  private readonly names: string[]
  private readonly expectedWordCount: number
  private readonly authority: string
  private readonly pageCache = new Map<string, number>()
  constructor(instruction: string, state?: DeliveryState) {
    this.goal = deliveryInstruction(instruction)
    this.zh = /[\u3400-\u9fff]/.test(this.goal)
    this.word = new WordDelivery(this.goal)
    this.reconcile =
      /为准|准数|核准.{0,12}一致|(?:多份|文件间|跨文件).{0,40}一致|唯一来源|以.{0,60}为(?:依据|基准)|source of truth|authoritative|match.{0,30}(?:cell|workbook|spreadsheet)|(?:预算表|工作簿|Excel|spreadsheet).{0,60}(?:同步|更新到|更新至|sync|propagate)/i.test(
        this.goal,
      ) &&
      (!/模板|版式|排版|样式|template|formatting/i.test(this.goal) ||
        /金额|核准|数[据字值]|指标|单元格|预算|期间|口径|amount|figures|budget|numbers|source of truth/i.test(
          this.goal,
        ))
    this.cells = this.reconcile ? explicitSourceCells(this.goal) : []
    this.names = outputNames(this.goal)
    const count =
      /(?:生成|创建|新建|另建|输出)(?:并保存)?\s*([2-9]|两|二|三|四|五|六|七|八|九)\s*(?:份|个)\s*(?:Word|DOCX|文档|报告|备忘录)/i.exec(
        this.goal,
      )?.[1] ??
      /(?:create|generate)\s+([2-9]|two|three|four)\s+(?:separate\s+)?(?:Word|DOCX|documents)/i.exec(
        this.goal,
      )?.[1]
    this.expectedWordCount = count
      ? Number(count) ||
        { two: 2, three: 3, four: 4 }[count.toLowerCase()] ||
        '一二三四五六七八九'.indexOf(count.replace('两', '二')) + 1
      : 0
    this.authority =
      /(?:以|根据|依据|using|use)\s*(.{1,600}?)(?:为(?:唯一)?准|为依据|as (?:the )?(?:authority|source of truth))/i
        .exec(this.goal)?.[1]
        ?.toLowerCase() ?? ''
    this.state = state ?? { outputs: [], sources: [] }
  }
  private wordFor(path: string) {
    const brief = wordBriefForFile(this.goal, path)
    return brief === this.goal ? this.word : new WordDelivery(brief)
  }
  private message(zh: string, en: string) {
    return this.zh ? zh : en
  }
  wantsSource(call: AgentToolCall) {
    if (!this.reconcile) return false
    if (['xlsx.read', 'xlsx.read_range'].includes(call.name))
      return typeof call.input.sheet === 'string'
    return (
      ['docx.read', 'pdf.read', 'pdf.read_pages'].includes(call.name) &&
      Array.isArray(call.input.deliveryFacts) &&
      call.input.deliveryFacts.length > 0
    )
  }
  async record(
    call: AgentToolCall,
    before: DocumentObservation[],
    after: DocumentObservation[],
    source: boolean,
  ) {
    if (source) {
      const target = after[0]?.version,
        sheet = String(call.input.sheet)
      if (!target) return false
      const filename = basename(target.path).toLowerCase()
      if (/\.(?:docx|pdf|xlsx|xlsm)\b/i.test(this.authority) && !this.authority.includes(filename))
        return false
      if (
        !this.authority.includes(filename) &&
        this.state.outputs.some((o) => o.target.path === target.path)
      )
        return false
      if (/^(?:docx|pdf)\./.test(call.name)) {
        const sources = (this.state.textSources ??= [])
        let reference = sources.find((s) => s.target.path === target.path)
        if (!reference) {
          reference = { target, facts: [] }
          sources.push(reference)
        }
        const proposed = Array.isArray(call.input.deliveryFacts) ? call.input.deliveryFacts : []
        for (const v of proposed.slice(0, 60)) {
          if (
            !v ||
            typeof v.label !== 'string' ||
            !v.label.trim() ||
            typeof v.value !== 'string' ||
            !v.value.trim() ||
            !(target.type === 'docx'
              ? typeof v.blockId === 'string'
              : Number.isInteger(v.page) && v.page > 0)
          )
            continue
          const fact: TextFact = {
            label: v.label.trim().slice(0, 200),
            value: v.value.trim().slice(0, 240),
            ...(target.type === 'docx' ? { blockId: v.blockId } : { page: v.page }),
            ...(Array.isArray(v.outputs)
              ? {
                  outputs: v.outputs
                    .filter((x: unknown) => typeof x === 'string')
                    .slice(0, 32)
                    .map((x: string) => basename(x)),
                }
              : {}),
          }
          const previous = reference.facts.findIndex(
            (f) => f.page === fact.page && f.blockId === fact.blockId && f.label === fact.label,
          )
          if (previous < 0) reference.facts.push(fact)
          else reference.facts[previous] = fact
        }
        return true
      }
      if (!/\.xls[xm]$/i.test(target.path)) return false
      let src = this.state.sources.find((s) => s.target.path === target.path && s.sheet === sheet)
      if (!src) {
        src = { target, sheet, facts: [] }
        this.state.sources.push(src)
      }
      const requested = Array.isArray(call.input.deliveryFacts) ? call.input.deliveryFacts : []
      for (const fact of requested.slice(0, 60)) {
        if (!fact || typeof fact !== 'object' || !/^[A-Z]{1,3}[1-9]\d{0,5}$/.test(fact.cell))
          continue
        const rule = {
          cell: String(fact.cell),
          ...(typeof fact.label === 'string' ? { label: fact.label.slice(0, 200) } : {}),
          ...(Array.isArray(fact.outputs)
            ? {
                outputs: fact.outputs
                  .filter((s: unknown) => typeof s === 'string')
                  .slice(0, 32)
                  .map((s: string) => basename(s)),
              }
            : {}),
        }
        const prior = (src.facts ??= []).findIndex((f) => f.cell === rule.cell)
        if (prior < 0) src.facts.push(rule)
        else src.facts[prior] = rule
      }
      if (!this.cells.length && !src.facts?.length) {
        const view = await readView(after[0]!).catch(() => undefined)
        const range = String(call.input.range ?? '').replace(/\$/g, '')
        const locs = view?.units.filter(
          (u) => u.location.kind === 'cell' && u.location.sheet === sheet,
        )
        const byAddress = new Map(
          (locs ?? []).flatMap((u) =>
            u.location.kind === 'cell' ? [[u.location.range, u] as const] : [],
          ),
        )
        let bounds
        try {
          bounds = range ? parseRange(range) : undefined
        } catch {
          return true
        }
        for (const u of locs ?? []) {
          if (u.location.kind !== 'cell' || !/^-?\d+(?:\.\d+)?$/.test(u.cellValue ?? u.text))
            continue
          const cell = u.location.range
          const { row, column } = parseAddress(cell)
          if (
            bounds &&
            (row < bounds.startRow ||
              row > bounds.endRow ||
              column < bounds.startColumn ||
              column > bounds.endColumn)
          )
            continue
          const rowLabel = byAddress.get('A' + cell.replace(/^[A-Z]+/, ''))?.text
          const colLabel = byAddress.get(cell.replace(/\d+$/, '1'))?.text
          if (
            cell === range ||
            (rowLabel &&
              this.goal.includes(rowLabel) &&
              colLabel &&
              this.goal.includes(colLabel.replace(/[（(].*$/, '')))
          )
            (src.facts ??= []).push({ cell, label: rowLabel })
        }
      }
      return true
    }
    const authored =
      (['docx.create', 'create_document', 'write_document'].includes(call.name) ||
        (call.name === 'docx.call' && call.input.name === 'write_document')) &&
      !/空白|空文档|空文件|blank|empty (?:document|file)/i.test(this.goal) &&
      /报告|周报|月报|信函|备忘录|纪要|总结|通知|撰写|起草|report|draft|briefing|memo|letter/i.test(
        this.goal,
      )
    if (
      !this.word.active &&
      !this.reconcile &&
      !authored &&
      !this.names.length &&
      !this.expectedWordCount
    )
      return false
    let changed = false
    for (const doc of after) {
      const target = doc.version
      if (
        target.type !== 'docx' &&
        !((this.reconcile || this.names.length) && deliverable.test(target.path))
      )
        continue
      const existing = this.state.outputs.find(
        (o) =>
          o.target.path === target.path ||
          (target.editorIdentity && o.target.editorIdentity === target.editorIdentity),
      )
      if (existing) {
        if (authored && target.type === 'docx')
          existing.before = { ...existing.before, authored: true }
        existing.target = {
          ...target,
          ...(existing.target.generatedFile ? { generatedFile: true } : {}),
        }
      } else {
        const original = before.find(
          (item) =>
            item.version.path === target.path ||
            (target.editorIdentity && item.version.editorIdentity === target.editorIdentity),
        )
        const word = this.wordFor(target.path)
        this.state.outputs.push({
          target,
          created:
            target.generatedFile === true ||
            call.name.endsWith('.create') ||
            !original?.version.exists,
          before:
            target.type === 'docx'
              ? { ...(await word.baseline(original)), ...(authored ? { authored: true } : {}) }
              : undefined,
        })
      }
      changed = true
    }
    return changed
  }
  async check(
    inspect: (target: TaskTargetVersion) => Promise<DocumentObservation>,
    render?: (doc: DocumentObservation) => Promise<number>,
  ): Promise<string[]> {
    this.observations = []
    const { outputs, sources } = this.state
    if (!outputs.length)
      return (this.names.length || this.expectedWordCount || this.word.pages) &&
        /(?:生成并保存|新建并保存|另建并保存|create and save)/i.test(this.goal) &&
        !/[?？]|如何|怎么|为什么|\b(?:how|why)\b/i.test(this.goal)
        ? [
            this.message(
              '请先生成并核对要求的交付文件。',
              'Create and check the requested deliverables.',
            ),
          ]
        : []
    const issues: string[] = []
    const createdWord = outputs.filter((o) => o.created && o.target.type === 'docx').length
    if (createdWord < this.expectedWordCount)
      issues.push(
        this.message(
          `要求生成 ${this.expectedWordCount} 份 Word 文件，已生成 ${createdWord} 份。只补齐缺少的文件。`,
          `Expected ${this.expectedWordCount} Word files, created ${createdWord}. Produce only the missing files.`,
        ),
      )
    const expected: Array<{
      origin: string
      value: string
      label: string
      header: string
      outputs?: string[]
    }> = []
    const files = new Map<string, Promise<DocumentObservation>>()
    const views = new Map<DocumentObservation, ReturnType<typeof readView>>()
    const load = (target: TaskTargetVersion) => {
      if (!files.has(target.path))
        files.set(
          target.path,
          inspect(target).catch(() => ({ version: { ...target, complete: false } })),
        )
      return files.get(target.path)!
    }
    const content = (doc: DocumentObservation) => {
      if (!views.has(doc))
        views.set(
          doc,
          readView(doc).catch(() => undefined),
        )
      return views.get(doc)!
    }
    const cells = new Set([
      ...this.cells,
      ...sources.flatMap((s) => s.facts?.map((f) => f.cell) ?? []),
    ])
    const textSources = this.state.textSources ?? []
    if (
      this.reconcile &&
      !cells.size &&
      !textSources.some((s) => s.facts.length) &&
      (/\.xlsx?|\.docx|\bPDF\b|Excel|工作簿|预算表|spreadsheet|workbook/i.test(this.goal) ||
        sources.length)
    )
      issues.push(
        this.message(
          '读取指定来源，在读取工具的 deliveryFacts 中注明原文位置、指标和使用它的成品。',
          'Read the authoritative source with deliveryFacts identifying the measure, original location and intended outputs.',
        ),
      )
    for (const source of textSources) {
      const doc = await load(source.target)
      const view =
        doc.version.type === 'pdf' && doc.packageSnapshot && !doc.version.dirty
          ? await observeDocument(doc.packageSnapshot, false).catch(() => undefined)
          : await content(doc)
      for (const fact of source.facts) {
        const origin = basename(source.target.path) + ' ' + (fact.blockId ?? 'p.' + fact.page)
        const at = view?.units.find((u) =>
          fact.blockId
            ? u.key === fact.blockId
            : u.location.kind === 'page' && u.location.page === fact.page,
        )
        if (at && at.text.includes(fact.label) && figureMatches(at.text, fact.value, ''))
          expected.push({
            origin,
            value: fact.value,
            label: fact.label,
            header: '',
            outputs: fact.outputs,
          })
        else
          issues.push(
            origin +
              ' · ' +
              this.message('请重新定位并核对原文中的', 'Locate the source text again for ') +
              fact.label,
          )
      }
    }
    for (const cell of cells) {
      const candidates = sources.filter(
        (s) =>
          s.facts?.some((f) => f.cell === cell) ||
          (this.cells.includes(cell) &&
            (sources.length === 1 ||
              this.goal.includes(basename(s.target.path)) ||
              this.goal.includes(s.sheet))),
      )
      for (const source of candidates) {
        const origin = basename(source.target.path) + ' ' + source.sheet + '!' + cell
        const doc = await load(source.target)
        const view = await content(doc)
        const row = cell.replace(/^[A-Z]+/, '')
        const at = (range: string) =>
          view?.units.find(
            (u) =>
              u.location.kind === 'cell' &&
              u.location.sheet === source.sheet &&
              u.location.range === range,
          )
        const item = at(cell),
          fact = source.facts?.find((f) => f.cell === cell)
        const value = (item?.cellValue ?? item?.text)?.trim()
        if (value && !value.startsWith('='))
          expected.push({
            origin,
            value,
            label: fact?.label ?? (cell === 'A' + row ? '' : (at('A' + row)?.text.trim() ?? '')),
            header: at(cell.replace(/\d+$/, '1'))?.text ?? '',
            outputs: this.cells.includes(cell) ? undefined : fact?.outputs,
          })
        else
          issues.push(
            origin + ' ' + this.message('需要可核对的单元格值。', 'needs a verifiable cell value.'),
          )
      }
      if (!candidates.length)
        issues.push(
          this.message('请从指定来源读取单元格 ', 'Read the specified source cell ') + cell,
        )
    }
    for (const name of this.names) {
      if (!outputs.some((o) => basename(o.target.path).toLowerCase() === name.toLowerCase()))
        issues.push(name + ' ' + this.message('尚未交付。', 'has not been delivered.'))
    }
    for (const output of outputs) {
      const name = basename(output.target.path)
      const desc = (zh: string, en: string) => issues.push(name + ' · ' + this.message(zh, en))
      const current = await load(output.target)
      if (!current.version.complete || (!current.version.exists && !current.snapshot)) {
        desc('无法读取该文件的当前版本。', 'The current file version is unavailable.')
        continue
      }
      if (current.version.type === 'docx') {
        const word = await this.wordFor(current.version.path).verify(
          current,
          output.before,
          render,
          this.pageCache,
          content,
          this.zh,
        )
        issues.push(...word.issues)
        this.observations.push(...word.checked)
      }
      if (!expected.length) continue
      const view = await content(current)
      if (!view || !view.units.length) {
        desc('无法回读当前成品。', 'Cannot inspect the current output.')
        continue
      }
      const rows = new Map<string, string[]>()
      for (const u of view.units) {
        const loc = u.location
        const key =
          loc.kind === 'cell'
            ? [loc.sheet, loc.range.replace(/^[A-Z]+/, '')].join('!')
            : loc.kind === 'slide'
              ? 'slide:' + loc.slide
              : u.key
        const row = rows.get(key) ?? []
        row.push((u.cellValue ?? u.text).normalize('NFKC'))
        rows.set(key, row)
      }
      const blocks = Array.from(rows.values(), (row) => row.join(' '))
      const relevant = expected.filter(
        (f) => !f.outputs?.length || f.outputs.some((p) => p.toLowerCase() === name.toLowerCase()),
      )
      const verified = new Map<string, { origin: string; value: string }>()
      for (const { origin, value, label, header } of relevant) {
        const key = JSON.stringify([label, header])
        const previous = verified.get(key)
        if (label && previous && previous.origin !== origin && previous.value !== value) {
          desc(
            `“${label}”在 ${previous.origin} 和 ${origin} 的值不同。先核对期间、单位和权威版本。`,
            `${label} differs between ${previous.origin} and ${origin}. Resolve the period, unit and authoritative version.`,
          )
          continue
        }
        if (label) verified.set(key, { origin, value })
        const linked = label ? blocks.filter((t) => t.includes(label)) : []
        const candidates = linked.length ? linked : label && this.goal.includes(label) ? [] : blocks
        if (!candidates.some((t) => figureMatches(t, value, header)))
          desc(
            `请核对“${label || origin}”，源值 ${origin} = ${value}。`,
            `Reconcile ${label || origin} with ${origin} = ${value}.`,
          )
        else this.observations.push(`${name} · ${origin} = ${value}`)
      }
    }
    return [...new Set(issues)].slice(0, 12)
  }
}
