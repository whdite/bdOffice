import { createHash } from 'node:crypto'
import { basename } from 'node:path'
import { wordPageLimit, wordLayout, pageMismatch } from '@genoffice/agent-core'
import type { DocumentObservation } from '../agent-tasks/document-observation'
import {
  observeDeliveryDocument as observeDocument,
  type EvidenceObservation,
} from './change-evidence'

type Unit = EvidenceObservation['units'][number]
type Heading = { name: string; level: number }
export interface WordBaseline {
  authored?: boolean
  requestedHeadings?: Heading[]
  range?: { start: number; tail: number; prefix: string; suffix: string }
  headings?: Heading[]
  otherBlocks?: string
  otherFormatting?: string
  live?: boolean
  unavailable?: boolean
  facts?: Array<{ key: string; index: number; values: string[] }>
}
const polish =
  /润色|精简|精练|不改变原意|保留原意|polish|copy[- ]?edit|preserve (?:the )?(?:meaning|facts)|improve (?:the )?wording/i
const mandatory =
  /不得不|必须|务必|须要|需(?:要)?(?:事[先前]|经|取得|由)|须(?!知)|应当|强制|(?<!无)需要|\b(?:must|shall|required|mandatory)\b/i
const prohibited =
  /不得(?!不)|严禁|禁止|不允许|不可(?:以)?|\b(?:must not|shall not|prohibited|forbidden|may not)\b/i
const prior =
  /事[先前]|预先|提前|之前|前.{0,5}(?:取得|获批)|\b(?:prior|beforehand|in advance|before)\b/i
const written = /书面|书写形式|\b(?:written|in writing)\b/i
const exception = /除非|除外|例外|除.{1,12}外|\b(?:except|unless)\b/i
const limits = /最高|至多|不超过|上限|最多|\b(?:maximum|at most|up to|cap(?:ped)?)\b/i
const negative = /无需|无须|不必|不需|免于|\b(?:need not|not required|no obligation)\b/i
const denied =
  /尚未|仍未|未获|未达|未能|并未|未曾|并非|并不|\b(?:did not|has not|have not|not yet|never)\b/i
function numbers(text: string): string[] {
  return Array.from(
    text
      .normalize('NFKC')
      .matchAll(
        /[-−]?\d+(?:,\d{3})*(?:\.\d+)?(?:\s*(?:万|亿|千|%)?(?:元|美元|人民币|万元|亿元|%))?/g,
      ),
    ([value]) =>
      value
        .replace(/,/g, '')
        .replace(/\s/g, '')
        .replace('−', '-')
        .replace(/([\d.]+)([万亿千])(?=元|美元)/, (_, n, unit) =>
          String(Number(n) * ({ 万: 1e4, 亿: 1e8, 千: 1e3 }[unit as string] ?? 1)),
        ),
  )
}
function facts(text: string): string[] {
  const groups: [RegExp, string][] = [
    [prohibited, '禁止 / prohibition'],
    [negative, '免除义务 / exemption'],
    [denied, '否定 / negation'],
    [prior, '事先 / prior'],
    [written, '书面 / written'],
    [exception, '例外 / exception'],
    [limits, '上限 / upper limit'],
  ]
  if (!negative.test(text) && !prohibited.test(text))
    groups.push([mandatory, '强制义务 / obligation'])
  return [
    ...numbers(text),
    ...groups.filter(([pattern]) => pattern.test(text)).map(([, label]) => label),
  ]
}
function scopedHeadings(goal: string, units: Unit[] = []) {
  const wanted = new Map<string, number>()
  if (!/(?:标题|heading)/i.test(goal)) return wanted
  for (let clause of goal.split(/[。；;！\n]/)) {
    for (const { text, headingLevel } of units) {
      if (
        headingLevel &&
        text &&
        text.length <= 120 &&
        clause.includes(text) &&
        !clause.includes('“' + text + '”') &&
        !clause.includes('"' + text + '"')
      )
        clause = clause.replace(text, () => '“' + text + '”')
    }
    for (const match of clause.matchAll(
      /[“「『"'‘]([^”」』"'’\n]{1,120})[”」』"'’]([^，,。；;"“”]{0,75})/g,
    )) {
      const tail = match[2] ?? ''
      const levels = [...tail.matchAll(/([一二三四五六1-6])\s*级|(?:level|heading)\s*([1-6])/gi)]
      const last = levels.at(-1)
      if (last) {
        const n = (last[1] ?? last[2])!
        wanted.set(match[1]!, Number(n) || '一二三四五六'.indexOf(n) + 1)
      }
    }
  }
  return wanted
}
const digest = (values: unknown) =>
  createHash('sha256').update(JSON.stringify(values)).digest('hex')
function otherBlocks(units: Unit[], names: Map<string, number>): string {
  return digest(units.filter((u) => !names.has(u.text)).map((u) => [u.text, u.headingLevel ?? 0]))
}
function scopeDigest(units: Unit[]) {
  return digest(units.map((u) => [u.text, u.headingLevel ?? 0]))
}
function paragraphRange(goal: string, units: Unit[]) {
  if (!/只|仅|其余|其他.*不变|\bonly\b/i.test(goal)) return
  const ordinal = /第\s*([1-9]\d*|[一二两三四五六七八九十])\s*(?:自然)?段/.exec(goal)
  const index = ordinal
    ? Number(ordinal[1]) || '一二三四五六七八九十'.indexOf(ordinal[1]!.replace('两', '二')) + 1
    : /\bfirst paragraph\b/i.test(goal)
      ? 1
      : 0
  if (!index) return
  const paragraphs = units.filter(
    (u) => u.location.kind === 'block' && !u.headingLevel && u.text.trim(),
  )
  const start = units.indexOf(paragraphs[index - 1]!)
  if (start < 0) return
  return {
    start,
    tail: units.length - start - 1,
    prefix: scopeDigest(units.slice(0, start)),
    suffix: scopeDigest(units.slice(start + 1)),
  }
}
function protectedBaseline(
  view: EvidenceObservation,
  headings: Map<string, number>,
  local: boolean,
  meaning: boolean,
  live: boolean,
  goal: string,
): WordBaseline {
  const range = paragraphRange(goal, view.units)
  return {
    ...(range ? { range } : {}),
    live,
    ...(local
      ? {
          headings: view.units
            .filter((u) => u.headingLevel)
            .map((u) => ({ name: u.text, level: u.headingLevel! })),
          otherBlocks: otherBlocks(view.units, headings),
          otherFormatting: digest(
            view.units.filter((u) => !headings.has(u.text)).map((u) => u.digest),
          ),
        }
      : {}),
    ...(meaning
      ? {
          facts: view.units.flatMap((u, index) => {
            const values = facts(u.text)
            return values.length ? [{ key: u.key, index, values }] : []
          }),
        }
      : {}),
  }
}

/** Word owns local editing invariants and physical layout acceptance for both editor and file workers. */
export class WordDelivery {
  readonly headings: Map<string, number>
  readonly local: boolean
  readonly meaning: boolean
  readonly localParagraph: boolean
  readonly pages
  readonly layout
  constructor(private readonly goal: string) {
    this.pages = wordPageLimit(goal)
    this.layout = wordLayout(goal)
    this.headings = scopedHeadings(this.goal)
    this.local =
      /标题|heading/i.test(goal) && /只|仅|其[余他].{0,12}(?:保持|不变)|\bonly\b/i.test(this.goal)
    this.localParagraph =
      /只|仅|\bonly\b/i.test(goal) &&
      /第\s*[\d一二两三四五六七八九十]+\s*段|first paragraph/i.test(goal)
    this.meaning =
      polish.test(this.goal) &&
      !/(?:修改|变更|调整)(?:金额|数字|义务|上限)|(?:change|update) the (?:amount|limit|obligation)/i.test(
        this.goal,
      )
  }
  get active(): boolean {
    return !!(
      this.pages ||
      Object.keys(this.layout).length ||
      this.headings.size ||
      this.local ||
      this.localParagraph ||
      this.meaning
    )
  }
  async baseline(original?: DocumentObservation): Promise<WordBaseline | undefined> {
    if (!original?.snapshot || !(this.local || this.meaning || this.localParagraph)) return
    const view = await observeDocument(original.snapshot, !!original.liveSnapshot).catch(
      () => undefined,
    )
    if (!view || view.incomplete) return { unavailable: true }
    const headings = this.local ? scopedHeadings(this.goal, view.units) : this.headings
    return {
      ...protectedBaseline(
        view,
        headings,
        this.local && headings.size > 0,
        this.meaning,
        !!original.liveSnapshot,
        this.goal,
      ),
      requestedHeadings: Array.from(headings, ([name, level]) => ({ name, level })),
    }
  }
  async verify(
    current: DocumentObservation,
    before: WordBaseline | undefined,
    render: ((doc: DocumentObservation) => Promise<number>) | undefined,
    pageCache: Map<string, number>,
    read: (doc: DocumentObservation) => Promise<EvidenceObservation | undefined>,
    zh: boolean,
  ) {
    const name = basename(current.version.path),
      issues: string[] = [],
      checked: string[] = []
    const message = (cn: string, en: string) => (zh ? cn : en)
    const desc = (cn: string, en: string) => issues.push(name + ' · ' + message(cn, en))
    const { pages, layout } = this
    if (Object.keys(layout).length) {
      const disk = !current.version.dirty && current.packageSnapshot
      const snap = disk || current.snapshot
      const observed =
        snap &&
        (await observeDocument(snap, !disk && !!current.liveSnapshot, true).catch(() => undefined))
          ?.wordLayout
      const { titleSize, titleAlign, columns } = layout
      if (
        titleSize &&
        (!observed?.title?.sizes.length || !observed.title.sizes.every((s) => s === titleSize))
      )
        desc(`主标题字号应为 ${titleSize} 磅。`, `Set the title to ${titleSize} pt.`)
      if (titleAlign && observed?.title?.align !== titleAlign)
        desc(`主标题对齐方式应为 ${titleAlign}。`, `Set title alignment to ${titleAlign}.`)
      if (columns) {
        const tables = observed?.tables.filter((t) => t.length === columns.length) ?? []
        if (
          !tables.length ||
          tables.some((t) =>
            t.some((w, i) => Math.abs((w * 100) / t.reduce((a, b) => a + b, 0) - columns[i]!) > 1),
          )
        )
          desc(
            `表格列宽应为 ${columns.join(' / ')}%。`,
            `Set table widths to ${columns.join(' / ')}%.`,
          )
      }
    }
    if (before?.unavailable)
      desc('原始版本的范围或事实未能核对。', 'The original scope or facts could not be verified.')
    if (pages) {
      const key = JSON.stringify([
        current.version.path,
        current.version.modelDigest,
        current.version.diskDigest,
        current.version.dirty,
      ])
      let n = pageCache.get(key)
      if (n === undefined && render) {
        n = await render(current).catch(() => undefined)
        if (n !== undefined) pageCache.set(key, n)
      }
      if (!n)
        desc('当前版本页数尚未取得，请核对分页。', 'The current layout has not been rendered.')
      else {
        const issue = pageMismatch(pages, n, zh)
        if (issue) issues.push(name + ' · ' + issue)
        else checked.push(name + ' · ' + n + (zh ? ' 页' : ' page(s)'))
      }
    }

    const headings = before?.requestedHeadings
      ? new Map(before.requestedHeadings.map((h) => [h.name, h.level]))
      : this.headings
    if (!headings.size && !this.meaning && !before?.range && !before?.authored)
      return { issues, checked }
    const view = await read(current)
    if (!view || view.incomplete) {
      desc('无法完整回读当前文档。', 'Cannot inspect all of the current document.')
      return { issues, checked }
    }
    const range = before?.range
    if (before?.authored) {
      const blocks = view.units.filter((u) => u.text.trim())
      if (!blocks.length)
        desc('文档缺少要求生成的正文。', 'The requested document has no body content.')
      else if (
        blocks.length >= 6 &&
        /周报|月报|评审报告|分析报告|business report|briefing/i.test(this.goal) &&
        !blocks.some((u) => u.headingLevel)
      )
        desc('请用 Word 标题样式组织报告。', 'Give the report native heading styles.')
      else checked.push(name + ' · ' + message('正文结构已回读', 'Document body inspected'))
    }
    if (
      range &&
      (scopeDigest(view.units.slice(0, range.start)) !== range.prefix ||
        scopeDigest(view.units.slice(view.units.length - range.tail)) !== range.suffix)
    )
      desc('恢复目标段落以外的内容。', 'Restore the content outside the specified paragraph.')
    for (const [title, level] of headings) {
      const actual = view.units.find((u) => u.text === title)
      if (actual?.headingLevel !== level)
        desc(
          `“${title}”应为 ${level} 级，当前 ${actual?.headingLevel ?? '—'}。`,
          `“${title}” should be heading ${level}, observed ${actual?.headingLevel ?? '—'}.`,
        )
      else checked.push(`${name} · ${title} → H${level}`)
    }
    if (
      before?.otherBlocks &&
      (before.otherBlocks !== otherBlocks(view.units, headings) ||
        (before.live === !!current.liveSnapshot &&
          before.otherFormatting !==
            digest(view.units.filter((u) => !headings.has(u.text)).map((u) => u.digest))))
    ) {
      const off = before.headings?.filter(
        (h) =>
          !headings.has(h.name) &&
          view.units.find((u) => u.text === h.name)?.headingLevel !== h.level,
      )
      desc(
        '恢复指定范围外的段落和标题层级。',
        'Restore paragraphs and headings outside the requested scope.',
      )
      if (off?.length) issues.push(off.map((h) => `${h.name} → H${h.level}`).join(', '))
    }
    let verifiedFacts = 0
    const byKey = new Map(view.units.map((u) => [u.key, u]))
    for (const rule of before?.facts ?? []) {
      const unit = byKey.get(rule.key) ?? view.units[rule.index]
      const observed = facts(unit?.text ?? '')
      const missing = rule.values.filter((v) => !observed.includes(v))
      if (missing.length)
        desc(
          `第 ${rule.index + 1} 段须保留原文的 ${missing.join(', ')}。`,
          `Restore in block ${rule.index + 1}: ${missing.join(', ')}.`,
        )
      else verifiedFacts += rule.values.length
    }
    if (verifiedFacts)
      checked.push(
        name +
          ' · ' +
          message(
            `核对了 ${verifiedFacts} 项原文数字和约束`,
            `${verifiedFacts} original figures and conditions checked`,
          ),
      )

    return { issues, checked }
  }
}
