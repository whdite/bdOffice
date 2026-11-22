import type { PageSpec, SpecElement } from './page-spec'

const themes = {
  classroom: {
    paper: '#F7F8F2',
    ink: '#163D37',
    body: '#243C37',
    accent: '#176E5C',
    pale: '#E2EFE6',
    warm: '#F3E8D0',
    bright: '#ACE0BB',
  },
  report: {
    paper: '#F8F7F3',
    ink: '#202E47',
    body: '#314157',
    accent: '#3B5683',
    pale: '#E6EBF3',
    warm: '#F4E8D7',
    bright: '#F0C784',
  },
}

export type DraftSlideTheme = keyof typeof themes

export function draftSlideTheme(topic: string): DraftSlideTheme {
  return /教学|学习|课堂|课件|课程|作业|学生|\b(?:lesson|teaching|learning|classroom)\b/i.test(
    topic,
  )
    ? 'classroom'
    : 'report'
}

export function draftSlideStyleGuide(theme: DraftSlideTheme): string {
  const c = themes[theme]
  return `Use Microsoft YaHei with ink ${c.ink} and accent ${c.accent}. Use warm white ${c.paper} for content and dark ${c.ink} with white text and ${c.bright} accent on the cover or closing. Use one clearly prioritized title per slide and large body text. Choose layouts based on the content: explanation, side-by-side comparison, ordered steps, a focused exercise, or a photo with commentary. Reserve charts and tables for briefs that contain data. Avoid decorative badges and repetitive card grids.`
}

type DraftLayout = 'cover' | 'explain' | 'compare' | 'steps' | 'practice' | 'summary' | 'photo'

interface Section {
  heading: string
  text: string
}

function contentString(value: unknown): string {
  return typeof value === 'string' ? value.trim() : ''
}

function draftLayout(value: unknown): DraftLayout | undefined {
  if (
    value === 'cover' ||
    value === 'explain' ||
    value === 'compare' ||
    value === 'steps' ||
    value === 'practice' ||
    value === 'summary' ||
    value === 'photo'
  )
    return value
  return
}

// CJK glyphs break individually. Keep Latin words together to leave room for
// PowerPoint's actual line breaks, and leave extra height for font fallback.
function linesNeeded(text: string, widthPx: number, sizePt: number): number {
  const limit = (widthPx - 12) / ((sizePt * 96) / 72)
  let lines = 0
  for (const paragraph of text.split('\n')) {
    let line = 0
    lines++
    for (const token of paragraph.match(/[A-Za-z0-9]{1,}|\s+|./gu) ?? []) {
      const width = [...token].reduce(
        (sum, ch) =>
          sum +
          (/[MW@%]/.test(ch)
            ? 1
            : /[A-Z]/.test(ch)
              ? 0.77
              : /[a-z\d]/.test(ch)
                ? 0.65
                : /[\s.,;:'"!|]/.test(ch)
                  ? 0.38
                  : 1.06),
        0,
      )
      if (width > limit) return Infinity
      if (line + width > limit) {
        lines++
        line = 0
      }
      line += width
    }
  }
  return lines
}

/** Turn the draft's text and layout role into editable native slide elements. */
export function renderDraftSlide(
  input: unknown,
  theme: DraftSlideTheme,
  allowedImages: readonly string[] = [],
): PageSpec | undefined {
  if (!input || typeof input !== 'object') return
  const raw = input as Record<string, unknown>
  const title = contentString(raw.title)
  const layout = draftLayout(raw.layout)
  if (!layout || !title || !Array.isArray(raw.sections) || raw.sections.length > 4) return
  const sections: Section[] = []
  for (const section of raw.sections) {
    if (!section || typeof section !== 'object') return
    const heading = contentString(section.heading)
    const text = contentString(section.text)
    if (!heading && !text) return
    sections.push({ heading, text })
  }
  if (
    (layout === 'compare' && sections.length !== 2) ||
    (layout === 'steps' && sections.length < 2) ||
    (layout === 'cover' && sections.length > 2) ||
    (layout !== 'cover' && sections.length === 0)
  )
    return

  const c = themes[theme]
  const elements: SpecElement[] = []
  let cannotFit = false
  const rect = (x: number, y: number, w: number, h: number, fill: string) => {
    elements.push({ type: 'shape', shape: 'rect', x, y, w, h, fill })
  }
  const textBox = (
    text: string,
    x: number,
    y: number,
    w: number,
    h: number,
    sizePt: number,
    minPt: number,
    color: string,
    bold = false,
  ) => {
    if (!text) return
    while (sizePt >= minPt && linesNeeded(text, w, sizePt) * sizePt * 1.92 + 6 > h) sizePt--
    if (sizePt < minPt) {
      cannotFit = true
      return
    }
    elements.push({
      type: 'text',
      x,
      y,
      w,
      h,
      valign: 'top',
      paragraphs: text.split('\n').map((line) => ({
        lineSpacingPct: 110,
        runs: [{ text: line, sizePt, bold, color, font: 'Microsoft YaHei' }],
      })),
    })
  }
  const joined = ({ heading, text }: Section) => [heading, text].filter(Boolean).join('\n')

  if (layout === 'cover') {
    rect(920, 0, 360, 720, c.accent)
    rect(970, 212, 310, 380, c.bright)
    rect(1038, 292, 242, 300, c.ink)
    rect(84, 132, 108, 5, c.bright)
    textBox(title, 84, 183, 778, 218, 56, 35, '#FFFFFF', true)
    sections.forEach((section, i) =>
      textBox(joined(section), 84, 432 + 91 * i, 764, 84, 23, 17, '#F1F5F0'),
    )
    return cannotFit ? undefined : { background: c.ink, elements }
  }

  const dark = layout === 'summary'
  const bodyColor = dark ? '#F1F5F0' : c.body
  const headingColor = dark ? c.bright : c.accent
  rect(76, 58, 68, 5, dark ? c.bright : c.accent)
  textBox(title, 76, 88, 1128, 106, 39, 28, dark ? '#FFFFFF' : c.ink, true)

  if (layout === 'compare' || layout === 'steps') {
    const steps = layout === 'steps'
    const gap = steps ? 24 : 32
    const width = Math.floor((1128 - gap * (sections.length - 1)) / sections.length)
    sections.forEach((section, index) => {
      const x = 76 + index * (width + gap)
      if (steps) {
        rect(x, 246, width, 3, c.accent)
        textBox(String(index + 1), x, 267, width, 66, 30, 27, c.accent, true)
      } else rect(x, 234, width, 397, index === 0 ? c.pale : c.warm)
      const inset = steps ? 0 : 38
      const headingY = steps ? 347 : 282
      const textY = section.heading ? headingY + 96 : headingY
      textBox(section.heading, x + inset, headingY, width - 2 * inset, 88, 27, 19, c.ink, true)
      textBox(section.text, x + inset, textY, width - 2 * inset, 599 - textY, 22, 17, c.body)
    })
  } else if (layout === 'photo') {
    const url = contentString(raw.image)
    if (!allowedImages.includes(url) || sections.length > 2) return
    elements.push({ type: 'image', x: 622, y: 222, w: 582, h: 426, url })
    sections.forEach((section, i) => {
      const y = 244 + 201 * i
      textBox(section.heading, 76, y, 484, 78, 25, 19, c.ink, true)
      const offset = section.heading ? 81 : 0
      textBox(section.text, 76, y + offset, 484, 191 - offset, 22, 17, c.body)
    })
  } else if (layout === 'practice') {
    const [task, ...details] = sections
    rect(76, 224, 1128, 218, c.pale)
    textBox(joined(task!), 116, 267, 1048, 128, 29, 20, c.ink, true)
    details.forEach((section, i) =>
      textBox(joined(section), 116, 466 + 64 * i, 1048, 61, 21, 16, c.body),
    )
  } else {
    const rowHeight = Math.min(146, Math.floor(435 / sections.length))
    sections.forEach((section, i) => {
      const y = 220 + i * rowHeight
      rect(76, y, 1128, 2, dark ? '#728282' : '#CDD9D3')
      textBox(section.heading, 76, y + 16, 285, rowHeight - 24, 25, 19, headingColor, true)
      const bodyX = section.heading ? 397 : 76
      const bodyW = section.heading ? 807 : 1128
      textBox(section.text, bodyX, y + 16, bodyW, rowHeight - 24, 23, 17, bodyColor)
    })
  }
  return cannotFit ? undefined : { background: dark ? c.ink : c.paper, elements }
}
