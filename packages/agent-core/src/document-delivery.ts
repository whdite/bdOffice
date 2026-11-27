/** Mechanical requirements come from the user's brief. Quoted skill examples are not user requests. */
export function deliveryInstruction(instruction: string): string {
  return instruction.startsWith('The user attached the following skill')
    ? instruction
        .split('</skill>')
        .at(-1)!
        .replace(/^\s*User request:\s*/i, '')
    : instruction
}

export const DOCUMENT_DELIVERY_GUIDE =
  'Own the deliverable from one brief. Choose suitable structure and readable layout for new work. Carry the audience, sources, template, scope and explicit limits into the writer. For local edits keep other content and formatting intact; for polishing preserve numbers, names, obligations, negations and exceptions. For multiple outputs track each source, period, unit and destination, and compare the actual results. Inspect and repair discrepancies within the current authorization before finishing. Ask once only if a consequential choice remains unresolved after inspecting available evidence.'

/** Select the relevant clauses when one brief assigns different limits to multiple Word files. */
export function wordBriefForFile(instruction: string, path: string): string {
  const raw = deliveryInstruction(instruction),
    name = path.split(/[\\/]/).at(-1)!.toLowerCase()
  if (!name || !raw.toLowerCase().includes(name)) return raw
  return raw
    .split(/[。；;\n]/)
    .filter((clause) => !/\.docx\b/i.test(clause) || clause.toLowerCase().includes(name))
    .join('。')
}

function count(value: string): number {
  if (/^\d+$/.test(value)) return Number(value)
  const digits = '零一二三四五六七八九'
  const [tens, ones] = value.replace('两', '二').split('十')
  return ones === undefined
    ? digits.indexOf(tens!)
    : (tens ? digits.indexOf(tens) : 1) * 10 + (ones ? digits.indexOf(ones) : 0)
}

export interface WordPageLimit {
  min: number
  max: number
}
export function pageMismatch(
  limit: WordPageLimit,
  actual: number,
  zh: boolean,
): string | undefined {
  if (actual >= limit.min && actual <= limit.max) return
  const wanted =
    limit.min === limit.max
      ? String(limit.min)
      : Number.isFinite(limit.max)
        ? limit.min + '–' + limit.max
        : limit.min + '+'
  return zh
    ? `要求 ${wanted} 页，实际 ${actual} 页。请调整列宽、间距或篇幅后重新分页。`
    : `Expected ${wanted} pages, rendered ${actual}. Adjust columns, spacing or length and recheck.`
}

export interface WordLayout {
  titleSize?: number
  titleAlign?: 'left' | 'center' | 'right'
  columns?: number[]
}
export function wordLayout(instruction: string): WordLayout {
  const goal = deliveryInstruction(instruction)
  const title =
    /(?:主标题|文档标题|(?<!级)标题|\btitle\b)([^。\n;；]{0,70})/i
      .exec(goal)?.[1]
      ?.split(/表格|正文|body text|table/i)[0] ?? ''
  const size = /(\d+(?:\.\d+)?)\s*(?:磅|pt\b|points?\b)/i.exec(title)?.[1]
  const columns = /(?:表格|列宽|columns?|table)([^。\n;；]{0,150})/i.exec(goal)?.[1] ?? ''
  const percents = Array.from(columns.matchAll(/(\d+(?:\.\d+)?)\s*[%％]/g), (m) => Number(m[1]))
  return {
    ...(size && Number(size) >= 1 && Number(size) <= 400 ? { titleSize: Number(size) } : {}),
    ...(/居中|cent(?:er|re)/i.test(title)
      ? { titleAlign: 'center' as const }
      : /右对齐|right.align/i.test(title)
        ? { titleAlign: 'right' as const }
        : /左对齐|left.align/i.test(title)
          ? { titleAlign: 'left' as const }
          : {}),
    ...(percents.length >= 2 &&
    percents.every((n) => n > 0) &&
    Math.abs(percents.reduce((a, b) => a + b, 0) - 100) < 1
      ? { columns: percents }
      : {}),
  }
}

/** Output length only. Page references, ranges to edit, and a source's page count are not limits. */
export function wordPageLimit(instruction: string): WordPageLimit | undefined {
  const goal = deliveryInstruction(instruction)
  const number = '(\\d{1,3}|[一二两三四五六七八九十]{1,3})'
  for (const clause of goal.split(/[，,。！？;；\n]/)) {
    if (
      /(?:pptx?|slides?|幻灯片|演示文稿|PDF)/i.test(clause) &&
      !/(?:docx?|word|周报|报告|文档|备忘录)/i.test(clause)
    )
      continue
    if (
      /(?:第\s*[\d一二三四五六七八九十]+\s*页|(?:读取|参见|引用|来源|见附件|摘取|提取|extract|read|refer to)\s*.{0,15}(?:页|pages?\b))/i.test(
        clause,
      ) &&
      !/(?:控制在|限制在|压缩到|不超过|within|at most)/i.test(clause)
    )
      continue
    const range = new RegExp(number + '\\s*[-–—至到]\\s*' + number + '\\s*页').exec(clause)
    const zh = new RegExp(
      '(?:^|(?:生成|制作|写成|撰写|写|做成|另建|新建|保存|排成|排版为|压缩到|控制在|限定为|限制在|保持|共|总共|须为|最多|不超过|至少|篇幅|要求|限定|长度|为|给我|的|word|docx|报告|周报|文档|备忘录|，|、)\\s*[”：:]?\\s*(?:一份|一篇)?\\s*)' +
        number +
        '\\s*页(?:内|以内|之内|纸)?',
      'i',
    ).exec(clause)
    const en =
      /\b(?:(at most|up to|no more than|at least|within)\s+)?(one|two|three|\d{1,3})(?:\s*[-–]\s*(\d{1,3}))?[- ]pages?\b/i.exec(
        clause,
      )
    if (!range && !zh && !en) continue
    if (
      en &&
      /\b(?:on|from|read|extract|edit|change|first|last)\s+(?:the\s+)?$/i.test(
        clause.slice(0, en.index),
      )
    )
      continue
    const unit = (s: string) => ({ one: 1, two: 2, three: 3 })[s.toLowerCase()] ?? count(s)
    const a = unit((range?.[1] ?? zh?.[1] ?? en?.[2])!)
    const b = unit(range?.[2] ?? en?.[3] ?? String(a))
    if (a < 1 || b < a) continue
    return {
      min: /(?:最多|不超过|以内|之内|页内|at most|up to|within|no more than)/i.test(clause) ? 1 : a,
      max: /(?:至少|at least)/i.test(clause) ? Infinity : b,
    }
  }
  return undefined
}
