import type { ParaFormat, Run } from './types'

export type DraftParagraphRole =
  'title' | 'heading' | 'subheading' | 'body' | 'list' | 'salutation' | 'signature' | 'date'

export function isNoticeDraft(title: string, instruction: string): boolean {
  return (
    /(?:通知|公告|通告)[。！!]?\s*$/.test(title) ||
    /(?:写|起草|拟|生成|撰写).{0,18}(?:通知|公告|通告)/.test(instruction)
  )
}

function isDraftDate(text: string): boolean {
  return /^(?:\d{4}年\s*\d{1,2}月\s*\d{1,2}日|\d{4}[-/.]\d{1,2}[-/.]\d{1,2})$/.test(text.trim())
}

export function draftBodyRole(
  notice: boolean,
  text: string,
  position: { opensBody: boolean; closesBody: boolean; nextText?: string | undefined },
): DraftParagraphRole {
  if (!notice) return 'body'
  if (position.closesBody && isDraftDate(text)) return 'date'
  if (
    position.nextText &&
    isDraftDate(position.nextText) &&
    /^[^。！：；\n]{1,24}(?:办公室|委员会|工作组|班委|教务处|[一二三四五六七八九十\d]班|团委|部门|部)$/.test(
      text,
    )
  )
    return 'signature'
  if (position.opensBody && /^[^。！；\n]{1,24}[：:]$/.test(text)) return 'salutation'
  return 'body'
}

/** Direct Word paragraph and run formatting for a newly drafted document. */
export function draftParagraphAppearance(
  notice: boolean,
  role: DraftParagraphRole,
): { format: ParaFormat; run: Omit<Run, 'text'> } {
  const heading = role === 'title' || role === 'heading' || role === 'subheading'
  const format: ParaFormat = { lineSpacing: heading ? 1.2 : notice ? 1.5 : 1.35 }
  const run: Omit<Run, 'text'> = {
    fontAscii: 'Calibri',
    eastAsiaFont: heading || !notice ? 'Microsoft YaHei' : 'SimSun',
    sizeHalfPoints: notice ? 24 : 22,
  }

  if (heading) {
    format.keepNext = true
    format.keepLines = true
    format.spaceBefore = role === 'title' ? 0 : 200
    format.spaceAfter = role === 'title' ? 320 : 100
    format.align = role === 'title' && notice ? 'center' : 'left'
    run.bold = true
    run.color = notice ? '202020' : '17354A'
    run.sizeHalfPoints = role === 'title' ? 44 : role === 'heading' ? 30 : 26
  } else {
    format.spaceAfter = role === 'list' ? 80 : notice ? 100 : 140
    if (notice && role === 'body') format.indentFirstLine = 480
    if (role === 'date' || role === 'signature') {
      format.align = 'right'
      format.spaceBefore = role === 'signature' ? 160 : 0
      if (role === 'signature') format.keepNext = true
    }
  }
  return { format, run }
}
