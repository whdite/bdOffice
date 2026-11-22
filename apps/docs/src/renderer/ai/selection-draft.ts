import { t } from '../i18n/locale'

const EXPAND = 'Expand the selected passage with useful explanations and details. Preserve its meaning, facts, and language.'
const SHORTEN = 'Make the selected passage more concise, following any requested length. Keep its main point and essential facts.'
const POLISH = 'Polish the selected passage for clarity, flow, and the requested tone. Preserve its meaning and facts.'
const REWRITE = 'Rewrite only the selected passage according to the user’s instructions. Preserve facts and all details the user wants kept.'

const REWRITE_ACTIONS = [
  { verb: /^(?:扩写|擴寫|扩充|擴充|扩展|擴展|expand\b|elaborate\b)/iu, plan: EXPAND },
  {
    verb: /^(?:缩写|縮寫|缩短|縮短|简写|簡寫|精简|精簡|精炼|精煉|压缩|壓縮|shorten\b|condense\b)/iu,
    plan: SHORTEN,
  },
  { verb: /^(?:润色|潤色|潤飾|polish\b|refine\b)/iu, plan: POLISH },
  { verb: /^(?:改写|改寫|重写|重寫|rewrite\b|rephrase\b)/iu, plan: REWRITE },
]

const trimCommand = (text: string): string => text.trim().replace(/[。！？.!?]+$/u, '').trim()

export const isSelectionDraftTrial = (text: string): boolean =>
  /^(?:试选区扩写|試選區擴寫|preview selection expansion)$/iu.test(trimCommand(text))

/** Straightforward requests can start the writer immediately. The planner handles additional instructions. */
export function selectionDraftRequest(instruction: string): { plan: string; standalone: boolean } | undefined {
  const text = trimCommand(instruction)
  if (isSelectionDraftTrial(text) || text === trimCommand(t('aiChipExpand')))
    return { plan: EXPAND, standalone: true }
  if (text === trimCommand(t('aiChipShorten'))) return { plan: SHORTEN, standalone: true }
  if (text === trimCommand(t('aiChipPolish')) || text === trimCommand(t('aiPolishSelectionPrompt')))
    return { plan: POLISH, standalone: true }

  const command = text
    .replace(/^(?:(?:[请請]|麻[烦煩])你?\s*|[帮幫]我\s*|直接\s*|再\s*|(?:please|could you|can you)\b\s*)+/iu, '')
    .replace(/^(?:[把将將给給对對]\s*)?(?:(?:当前|當前)?[选選][区區]|(?:选中|選取)的?(?:文字|内容|內容|文本|部分)|[这這]段(?:文字|内容|內容|[话話])?|[这這]句[话話]?|原文)\s*/u, '')
  const action = REWRITE_ACTIONS.find(({ verb }) => verb.test(command))
  if (!action) return undefined
  const details = command.replace(action.verb, '').trim()
  if (/^(?:是什么|是什麼|是什么意思|是什麼意思|的意思|的区别|的區別|怎么|怎麼|如何)/u.test(details))
    return undefined
  const standalone =
    /^(?:(?:一[下点點些]|[这這]段(?:文字|内容|內容|[话話])?|原文|(?:选中|選取)的(?:文字|内容|內容)|吧)\s*)*$/u.test(details) ||
    /^(?:(?:this|the selected|the) (?:text|passage|selection|paragraph))$/iu.test(details)
  return { plan: action.plan, standalone }
}
