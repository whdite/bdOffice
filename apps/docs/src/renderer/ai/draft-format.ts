import type { Editor } from '@tiptap/core'
import {
  draftParagraphAppearance,
  draftBodyRole,
  isNoticeDraft,
  type DraftParagraphRole,
} from '@genoffice/docx-engine'
import { hasExplicitOfficeFormatting } from '@genoffice/agent-core'
import type { PmNode } from '../editor/convert'
import { isBlankDocument } from './protocol'

export function shouldFormatDraft(editor: Editor, instruction: string | undefined): boolean {
  if (!instruction || hasExplicitOfficeFormatting(instruction) || !isBlankDocument(editor))
    return false
  const attrs = editor.state.doc.firstChild?.attrs
  if (!attrs) return false
  return (
    (!attrs.styleId || attrs.styleId === 'Normal') &&
    ['align', 'lineSpacing', 'spaceBefore', 'spaceAfter', 'indentFirstLine', 'shadingFill'].every(
      (key) => attrs[key] == null,
    )
  )
}

export function formatDraft(nodes: PmNode[], instruction: string): PmNode[] {
  const text = (node: PmNode | undefined) =>
    node?.content
      ?.map((child) => child.text ?? '')
      .join('')
      .trim() ?? ''
  const title = nodes[0]?.type === 'docHeading' && nodes[0].attrs?.level === 1
  const notice = isNoticeDraft(title ? text(nodes[0]) : '', instruction)

  return nodes.map((node, index) => {
    let role: DraftParagraphRole
    if (node.type === 'docHeading')
      role =
        title && index === 0 ? 'title' : Number(node.attrs?.level) <= 2 ? 'heading' : 'subheading'
    else if (node.type === 'docListItem') role = 'list'
    else if (node.type === 'docParagraph')
      role = draftBodyRole(notice, text(node), {
        opensBody: index === (title ? 1 : 0),
        closesBody: index === nodes.length - 1,
        nextText: index === nodes.length - 2 ? text(nodes[index + 1]) : undefined,
      })
    else return node

    const { format, run } = draftParagraphAppearance(notice, role)
    const color = run.color
    return {
      ...node,
      attrs: { ...format, ...node.attrs },
      content: node.content?.map((child) => {
        if (child.type !== 'text' || child.marks?.some((mark) => mark.type === 'docTextStyle'))
          return child
        const marks = [...(child.marks ?? [])]
        if (run.bold && !marks.some((mark) => mark.type === 'bold')) marks.push({ type: 'bold' })
        marks.push({
          type: 'docTextStyle',
          attrs: {
            fontAscii: run.fontAscii,
            eastAsiaFont: run.eastAsiaFont,
            sizeHalfPoints: run.sizeHalfPoints,
            ...(color && !marks.some((mark) => mark.type === 'link') ? { color } : {}),
          },
        })
        return { ...child, marks }
      }),
    }
  })
}
