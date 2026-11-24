import type { Lang } from '@genoffice/i18n'

/** The Workspace Agent has its own surface because its event stream is not part
 * of the home-screen dictionary. Keep the copy small and deterministic: model
 * text, file names, paths, and tool/API identifiers are never translated here. */
export interface WorkspaceChatCopy {
  panelTitle: string
  workspaceFallback: string
  workspacePanel: string
  changeWorkspace: string
  openWindow: string
  focusWindow: string
  windowFailed: string
  dockWindow: string
  toggleSidebar: string
  hideKeepsRunning: string
  statusDetails: string
  viewApproval: string
  approvalMode: string
  approvalAuto: string
  approvalPerAction: string
  approvalAutoApplied: string
  approvalPerActionApplied: string
  approvalLoading: string
  cancel: string
  save: string
  saving: string
  retry: string
  workspacePickerTitle: string
  projectWorkspaces: string
  customFolder: string
  projectSpace: string
  projectFileCount(count: number): string
  noProjectWorkspaces: string
  resizeChat: string
  resizeWorkspace: string
  resizeWorkflow: string
  close: string
  loadingWorkspace: string
  chooseWorkspace: string
  chooseWorkspaceDescription: string
  chooseFolder: string
  sessions: string
  newSession: string
  noSessions: string
  messageCount(count: number): string
  rename: string
  delete: string
  renameSession(title: string): string
  deleteSession(title: string): string
  files: string
  refreshFiles: string
  noFiles: string
  activeSessionFallback: string
  running: string
  ready: string
  workflow: string
  received: string
  inProgress: string
  complete: string
  needsAttention: string
  welcomeTitle: string
  welcomeDescription: string
  you: string
  agent: string
  reasoningSummary: string
  fileRead: string
  fileCreated: string
  fileUpdated: string
  plan: string
  deny: string
  allowOnce: string
  allowedOnce: string
  allowRound: string
  allowedRound: string
  roundApprovalScope: string
  denied: string
  expired: string
  agentRunUndone: string
  agentRunRollbackAttention: string
  updatedFiles(count: number): string
  undoing: string
  undoRun: string
  reviewChanges: string
  toolRunning: string
  workStopped: string
  workingSteps(count: number): string
  finishedSteps(count: number): string
  incompleteSteps(count: number, incomplete: number): string
  toolResult(index: number): string
  askPlaceholder: string
  messageAria: string
  stop: string
  steer: string
  send: string
  enterHint: string
  chooseWorkspaceError: string
  agentFailed: string
  closeChat: string
  hideChat: string
  showChat: string
  runStat(key: string): string
  operation(value: string): string
}

const EN: WorkspaceChatCopy = {
  panelTitle: 'AI BigDog',
  workspaceFallback: 'No workspace selected',
  workspacePanel: 'Workspace',
  changeWorkspace: 'Change workspace',
  openWindow: 'Open in separate window',
  focusWindow: 'Show Agent window',
  windowFailed: 'Agent unavailable · Retry',
  dockWindow: 'Dock beside editor',
  toggleSidebar: 'Sessions and files',
  hideKeepsRunning: 'Tasks continue when hidden, until you use Stop or quit the app.',
  statusDetails: 'Details',
  viewApproval: 'View approval',
  approvalMode: 'Tool approval mode',
  approvalAuto: 'Auto execute',
  approvalPerAction: 'Approve each',
  approvalAutoApplied: 'Auto execute enabled. You can switch back at any time.',
  approvalPerActionApplied:
    'Approval required for subsequent protected actions. Running commands are not stopped.',
  approvalLoading: 'Waiting for the saved approval setting…',
  cancel: 'Cancel',
  save: 'Save',
  saving: 'Saving…',
  retry: 'Retry',
  workspacePickerTitle: 'Choose Agent workspace',
  projectWorkspaces: 'Office project spaces',
  customFolder: 'Custom folder',
  projectSpace: 'Project space',
  projectFileCount: (count) => `${count} file${count === 1 ? '' : 's'}`,
  noProjectWorkspaces: 'No Office project spaces yet',
  resizeChat: 'Resize BigDog Agent',
  resizeWorkspace: 'Resize workspace panel',
  resizeWorkflow: 'Resize workflow panel',
  close: 'Close',
  loadingWorkspace: 'Preparing your conversation...',
  chooseWorkspace: 'Choose a workspace',
  chooseWorkspaceDescription:
    'Use a project folder as the shared context for files and Agent sessions.',
  chooseFolder: 'Choose folder',
  sessions: 'Sessions',
  newSession: 'New session',
  noSessions: 'No sessions',
  messageCount: (count) => `${count} message${count === 1 ? '' : 's'}`,
  rename: 'Rename',
  delete: 'Delete',
  renameSession: () => 'Session name',
  deleteSession: (title) => `Delete session "${title}"?`,
  files: 'Files',
  refreshFiles: 'Refresh files',
  noFiles: 'No files in this workspace',
  activeSessionFallback: 'New session',
  running: 'Running',
  ready: 'Ready',
  workflow: 'Agent workflow',
  received: 'Received',
  inProgress: 'In progress',
  complete: 'Response finished',
  needsAttention: 'Needs attention',
  welcomeTitle: 'Ask AI BigDog',
  welcomeDescription: 'Ask, create documents, or choose files and memory for this conversation.',
  you: 'You',
  agent: 'AI BigDog',
  reasoningSummary: 'Reasoning summary',
  fileRead: 'Read',
  fileCreated: 'Created',
  fileUpdated: 'Updated',
  plan: 'Plan',
  deny: 'Deny',
  allowOnce: 'Allow once',
  allowedOnce: 'Allowed once',
  allowRound: 'Allow this file for the round',
  allowedRound: 'Approved for this round',
  roundApprovalScope: 'Edits and save for this file in this round. Base version',
  denied: 'Denied',
  expired: 'Expired',
  agentRunUndone: 'Agent run undone',
  agentRunRollbackAttention: 'Agent run rollback needs attention',
  updatedFiles: (count) => `Updated ${count} file${count === 1 ? '' : 's'}`,
  undoing: 'Undoing...',
  undoRun: 'Undo run',
  reviewChanges: 'Review changes',
  toolRunning: 'Running',
  workStopped: 'Interrupted',
  workingSteps: (count) => 'Working · ' + count + ' steps',
  finishedSteps: (count) => 'Worked · ' + count + ' steps',
  incompleteSteps: (count, incomplete) => count + ' steps · ' + incomplete + ' incomplete',
  toolResult: (index) => `Tool result ${index}`,
  askPlaceholder: 'Ask, create a document, or continue with files and memory...',
  messageAria: 'Message AI BigDog',
  stop: 'Stop',
  steer: 'Steer',
  send: 'Send',
  enterHint: 'Enter to send - Shift+Enter for a new line',
  chooseWorkspaceError: 'Choose a workspace first.',
  agentFailed: 'Agent failed.',
  closeChat: 'Close BigDog Agent',
  hideChat: 'Hide BigDog Agent',
  showChat: 'AI BigDog',
  runStat: (key) =>
    ({
      textLength: 'Text',
      pageEstimate: 'Pages',
      slideCount: 'Slides',
      elementCount: 'Elements',
      cellCount: 'Cells',
      pageCount: 'Pages',
      sections: 'Sections',
      comments: 'Comments',
    })[key] ?? key,
  operation: (value) => value,
}

const ZH: WorkspaceChatCopy = {
  panelTitle: 'AI \u5927\u72d7',
  workspaceFallback: '未选择工作区',
  workspacePanel: '工作区',
  changeWorkspace: '切换工作区',
  openWindow: '在独立窗口中打开',
  focusWindow: '显示 Agent 独立窗口',
  windowFailed: 'Agent 打开失败 · 重试',
  dockWindow: '返回编辑器侧栏',
  toggleSidebar: '会话与文件',
  hideKeepsRunning: '收起后任务继续；点击“停止”或退出应用才会中止。',
  statusDetails: '详情',
  viewApproval: '查看待批准操作',
  approvalMode: '工具批准方式',
  approvalAuto: '自动执行',
  approvalPerAction: '每次审批',
  approvalAutoApplied: '已启用自动执行，可随时切回每次审批。',
  approvalPerActionApplied: '后续受保护操作将逐次审批，已运行的命令不会被中止。',
  approvalLoading: '正在等待已保存的审批设置…',
  cancel: '取消',
  save: '保存',
  saving: '正在保存…',
  retry: '重试',
  workspacePickerTitle: '选择 Agent 工作区',
  projectWorkspaces: 'Office 项目空间',
  customFolder: '自定义文件夹',
  projectSpace: '项目空间',
  projectFileCount: (count) => `${count} 个文件`,
  noProjectWorkspaces: '暂无 Office 项目空间',
  resizeChat: '调整大狗 Agent 大小',
  resizeWorkspace: '调整工作区面板大小',
  resizeWorkflow: '调整工作流面板大小',
  close: '关闭',
  loadingWorkspace: '正在准备对话…',
  chooseWorkspace: '选择工作区',
  chooseWorkspaceDescription: '选择一个项目文件夹，作为文件和 Agent 会话的共享上下文。',
  chooseFolder: '选择文件夹',
  sessions: '会话',
  newSession: '新建会话',
  noSessions: '暂无会话',
  messageCount: (count) => `${count} 条消息`,
  rename: '重命名',
  delete: '删除',
  renameSession: () => '会话名称',
  deleteSession: (title) => `删除会话“${title}”？`,
  files: '文件',
  refreshFiles: '刷新文件',
  noFiles: '工作区暂无文件',
  activeSessionFallback: '新会话',
  running: '运行中',
  ready: '就绪',
  workflow: 'Agent 工作流',
  received: '已收到',
  inProgress: '处理中',
  complete: '已回复',
  needsAttention: '需要处理',
  welcomeTitle: '\u95ee\u95ee AI \u5927\u72d7',
  welcomeDescription:
    '\u53ef\u4ee5\u63d0\u95ee\u3001\u521b\u5efa\u6587\u6863\uff0c\u6216\u9009\u62e9\u6587\u4ef6\u4e0e\u8bb0\u5fc6\u4f5c\u4e3a\u4e0a\u4e0b\u6587\u3002',
  you: '你',
  agent: 'AI \u5927\u72d7',
  reasoningSummary: '思考摘要',
  fileRead: '已读取',
  fileCreated: '已创建',
  fileUpdated: '已更新',
  plan: '计划',
  deny: '拒绝',
  allowOnce: '仅允许本次',
  allowedOnce: '已允许本次',
  allowRound: '允许本轮编辑并保存',
  allowedRound: '已批准本轮',
  roundApprovalScope: '本轮可编辑并保存这份文件，基线版本',
  denied: '已拒绝',
  expired: '已过期',
  agentRunUndone: 'Agent 操作已撤销',
  agentRunRollbackAttention: 'Agent 操作回滚需要处理',
  updatedFiles: (count) => `已更新 ${count} 个文件`,
  undoing: '正在撤销…',
  undoRun: '回退本轮',
  reviewChanges: '查看变更',
  toolRunning: '运行中',
  workStopped: '已中止',
  workingSteps: (count) => '正在处理 · ' + count + ' 步',
  finishedSteps: (count) => '已处理 · ' + count + ' 步',
  incompleteSteps: (count, incomplete) => count + ' 步 · ' + incomplete + ' 步未完成',
  toolResult: (index) => `工具结果 ${index}`,
  askPlaceholder:
    '\u63d0\u95ee\u3001\u521b\u5efa\u6587\u6863\uff0c\u6216\u7ed3\u5408\u6587\u4ef6\u4e0e\u8bb0\u5fc6\u7ee7\u7eed\u5de5\u4f5c...',
  messageAria: '\u53d1\u9001\u6d88\u606f\u7ed9 AI \u5927\u72d7',
  stop: '停止',
  steer: '调整指令',
  send: '发送',
  enterHint: '按 Enter 发送，按 Shift+Enter 换行',
  chooseWorkspaceError: '请先选择工作区。',
  agentFailed: 'Agent 执行失败。',
  closeChat: '关闭大狗 Agent',
  hideChat: '隐藏大狗 Agent',
  showChat: 'AI \u5927\u72d7',
  runStat: (key) =>
    ({
      textLength: '文本',
      pageEstimate: '页数',
      slideCount: '幻灯片',
      elementCount: '元素',
      cellCount: '单元格',
      pageCount: '页数',
      sections: '节',
      comments: '批注',
    })[key] ?? key,
  operation: (value) =>
    ({ modified: '已修改', created: '已创建', deleted: '已删除' })[value] ?? value,
}

export function isChineseWorkspaceLanguage(lang: Lang | string | undefined): boolean {
  return typeof lang === 'string' && /^(?:zh|zh[-_])/iu.test(lang)
}

export function workspaceChatCopy(lang: Lang | string | undefined): WorkspaceChatCopy {
  return isChineseWorkspaceLanguage(lang) ? ZH : EN
}

/** Localize the built-in title of legacy sessions without touching titles that
 * the user explicitly renamed. */
export function workspaceSessionTitle(title: string, lang: Lang | string | undefined): string {
  if (
    isChineseWorkspaceLanguage(lang) &&
    (title === EN.panelTitle || title === 'Workspace Chat' || title === '工作区对话')
  )
    return ZH.panelTitle
  return title
}

const TOOL_LABELS_ZH: Record<string, string> = {
  'computer.status': '检查电脑操作组件',
  'computer.list_apps': '查看运行中的应用',
  'computer.list_windows': '查找应用窗口',
  'computer.observe': '读取应用窗口',
  'computer.click': '点击应用控件',
  'computer.type_text': '在应用中输入',
  'computer.set_value': '填写应用控件',
  'computer.press_key': '向应用发送按键',
  'computer.scroll': '滚动应用窗口',
  'computer.drag': '拖动窗口内项目',
  'computer.activate': '切换到应用窗口',
  'computer.launch': '打开应用',
  'context.read_file': '读取参考文件',
  'knowledge.list': '列出知识库资料',
  'knowledge.search': '检索知识库',
  'knowledge.read': '读取知识库资料',
  'session.read_summary': '读取会话记忆',
  'session.search_memory': '检索会话记忆',
  'agent.update_plan': '更新计划',
  'workspace.list': '列出工作区文件',
  'workspace.list_files': '列出工作区文件',
  'workspace.search': '搜索工作区',
  'workspace.locations': '查看已添加位置',
  'workspace.find_files': '查找本机文件',
  'workspace.add_paths': '添加到工作区',
  'workspace.search_files': '搜索工作区',
  'workspace.run_command': '运行工作区命令',
  'web.search': '搜索网页',
  'workspace.open': '打开工作区文件',
  'workspace.open_file': '打开工作区文件',
  'workspace.file_info': '查看文件信息',
  'workspace.get_file': '读取工作区文件',
  'workspace.get_file_metadata': '查看文件信息',
  'office.inspect': '检查 Office 文档',
  'office.read': '读取 Office 文档',
  'office.save': '保存 Office 文档',
  'office.edit': '编辑 Office 文档',
  'docx.create': '创建 Word 文档',
  'docx.inspect': '检查 Word 文档',
  'docx.search': '搜索 Word 文档',
  'docx.read': '读取 Word 文档',
  'docx.read_blocks': '读取 Word 文档段落',
  'docx.edit': '编辑 Word 文档',
  'docx.replace_blocks': '替换 Word 文档内容',
  'docx.insert_content': '插入 Word 文档内容',
  'docx.save': '保存 Word 文档',
  'docx.read_comments': '读取 Word 批注',
  'docx.reply_comment': '回复 Word 批注',
  'docx.resolve_comment': '处理 Word 批注',
  'docx.read_revisions': '读取 Word 修订',
  'pptx.inspect': '检查 PowerPoint',
  'pptx.search': '搜索 PowerPoint',
  'pptx.read': '读取 PowerPoint',
  'pptx.read_slides': '读取 PowerPoint 幻灯片',
  'pptx.edit': '编辑 PowerPoint',
  'pptx.edit_slide': '编辑 PowerPoint 幻灯片',
  'pptx.create': '创建 PowerPoint',
  'pptx.create_deck': '创建 PowerPoint 演示文稿',
  'pptx.render': '渲染 PowerPoint 幻灯片',
  'pptx.render_slide': '渲染 PowerPoint 幻灯片',
  'pptx.regenerate_slide': '重新生成 PowerPoint 幻灯片',
  'pptx.save': '保存 PowerPoint',
  'xlsx.inspect': '检查 Excel 工作簿',
  'xlsx.list_sheets': '列出 Excel 工作表',
  'xlsx.create': '创建 Excel 工作簿',
  'xlsx.read': '读取 Excel 工作簿',
  'xlsx.read_range': '读取 Excel 单元格区域',
  'xlsx.search': '搜索 Excel 工作簿',
  'xlsx.edit': '编辑 Excel 工作簿',
  'xlsx.write_range': '写入 Excel 单元格区域',
  'xlsx.set_formula': '设置 Excel 公式',
  'xlsx.save': '保存 Excel 工作簿',
  'xlsx.create_chart': '创建 Excel 图表',
  'xlsx.edit_chart': '编辑 Excel 图表',
  'xlsx.format_range': '设置 Excel 区域格式',
  'xlsx.insert_sheet': '插入 Excel 工作表',
  'xlsx.delete_sheet': '删除 Excel 工作表',
  'xlsx.sort': '排序 Excel 区域',
  'xlsx.filter': '筛选 Excel 区域',
  'pdf.inspect': '检查 PDF',
  'pdf.search': '搜索 PDF',
  'pdf.read_pages': '读取 PDF 页面',
  'pdf.extract_text': '提取 PDF 文本',
  'pdf.read': '读取 PDF',
  'pdf.annotate': '添加 PDF 批注',
  'pdf.extract_images': '提取 PDF 图片',
  'pdf.export_pages': '导出 PDF 页面',
  'md.create': '创建 Markdown',
  'md.inspect': '检查 Markdown',
  'md.search': '搜索 Markdown',
  'md.read': '读取 Markdown',
  'md.edit': '编辑 Markdown',
  'md.insert': '插入 Markdown 内容',
  'md.save': '保存 Markdown',
  'environment.check': '检查运行环境',
  'environment.plan_repair': '准备修复方案',
  'environment.repair': '修复运行环境',
}

/** Translate known runtime summaries while preserving paths and model output. */
export function workspaceToolLabel(name: string, lang: Lang | string | undefined): string {
  if (isChineseWorkspaceLanguage(lang)) return TOOL_LABELS_ZH[name] ?? name
  const environmentLabels: Record<string, string> = {
    'environment.check': 'Check runtime environment',
    'environment.plan_repair': 'Prepare repair plan',
    'environment.repair': 'Repair runtime environment',
  }
  return environmentLabels[name] ?? name
}

export function workspaceToolSummary(
  summary: string | undefined,
  lang: Lang | string | undefined,
): string | undefined {
  if (!summary || !isChineseWorkspaceLanguage(lang)) return summary
  const environmentSummaries: Record<string, string> = {
    'Previous effect verified; no write repeated.': '已核对上次结果，本次未重复写入',
    'Runtime environment checked': '运行环境已检查',
    'Repair plan prepared': '修复方案已准备',
    'Environment repair verified': '环境修复已完成并复查',
    'Environment repair needs attention': '环境修复需要处理',
  }
  if (environmentSummaries[summary]) return environmentSummaries[summary]
  let value = summary
  value = value.replace(/^Updated plan \((\d+) steps\)$/iu, '已更新计划（$1 步）')
  value = value.replace(/^Listed (\d+) file\(s\)$/iu, '已列出 $1 个文件')
  value = value.replace(/^Found (\d+) match\(es\)$/iu, '找到 $1 个匹配项')
  value = value.replace(/^Web search · (\d+) result\(s\)$/iu, '网页搜索 · $1 个结果')
  value = value.replace(/^Rendered Slide (\d+)$/iu, '已渲染第 $1 张幻灯片')
  value = value.replace(/^Updated (\d+) spreadsheet cell\(s\)\.$/iu, '已更新 $1 个电子表格单元格')
  value = value.replace(/^Updated (\d+) document block\(s\)\.$/iu, '已更新 $1 个文档段落')
  value = value.replace(/^Updated (\d+) slide text run\(s\)\.$/iu, '已更新 $1 个幻灯片文本')
  value = value.replace(/^Inserted content into the document\.$/iu, '已向文档插入内容')
  value = value.replace(/^Inserted Markdown content\.$/iu, '已插入 Markdown 内容')
  value = value.replace(
    /^Updated Markdown \((\d+) replacements?\)\.$/iu,
    '已更新 Markdown（$1 处）',
  )
  value = value.replace(
    /^Added PDF (.+) annotation on page (\d+)\.$/iu,
    '已在 PDF 第 $2 页添加 $1 批注',
  )
  value = value.replace(/^No matching (?:document blocks|slide text)\.?$/iu, '没有找到匹配内容')
  value = value.replace(/^([A-Za-z ]+) (.+)$/u, (match, action: string, rest: string) => {
    const actions: Record<string, string> = {
      Open: '打开',
      Read: '读取',
      Search: '搜索',
      Inspect: '检查',
      Save: '保存',
      Create: '创建',
      Updated: '已更新',
      Inserted: '已插入',
      Regenerated: '已重新生成',
      Resolve: '处理',
      Reply: '回复',
      'List sheets in': '列出工作表：',
    }
    const translated = actions[action.trim()]
    return translated ? `${translated}${translated.endsWith('：') ? '' : ' '}${rest}` : match
  })
  return value
}

export function workspaceToolDisplayText(
  value: string | undefined,
  lang: Lang | string | undefined,
): string | undefined {
  if (!value || !isChineseWorkspaceLanguage(lang)) return value
  return value
    .replace(/^(\d+) chars$/iu, '$1 个字符')
    .replace(/^\[binary content\]$/iu, '[二进制内容]')
}

/** Translate only scalar status lines; structured tool output and document
 * content must remain byte-for-byte recognizable to the user and the model. */
export function workspaceToolOutput(
  value: string | undefined,
  lang: Lang | string | undefined,
): string | undefined {
  if (!value || !isChineseWorkspaceLanguage(lang)) return value
  return value
    .replace(/^Opened (.+)$/iu, '已打开 $1')
    .replace(/^Created (.+)$/iu, '已创建 $1')
    .replace(/^Saved (.+)$/iu, '已保存 $1')
    .replace(/^Unknown tool: (.+)$/iu, '未知工具：$1')
}

export function workspaceDisplayTitle(
  value: string | undefined,
  lang: Lang | string | undefined,
): string | undefined {
  if (!value || !isChineseWorkspaceLanguage(lang)) return value
  return value.replace(/\s-\sSlide (\d+)$/iu, ' · 第 $1 张幻灯片')
}

export function workspacePermissionText(value: string, lang: Lang | string | undefined): string {
  if (!isChineseWorkspaceLanguage(lang)) return value
  const exact: Record<string, string> = {
    'Run Windows shell command': '运行 Windows Shell 命令',
    'Run workspace command': '运行工作区命令',
    'Confirm destructive action': '确认破坏性操作',
    'Confirm document change': '确认文档变更',
  }
  if (exact[value]) return exact[value]
  if (/^Command: /u.test(value)) {
    return value.replace(
      /^Command: ([\s\S]*?)\nWorking directory: ([\s\S]*?)\nWarning: this command may modify files outside the workspace and cannot be undone automatically\.?$/iu,
      '命令：$1\n工作目录：$2\n警告：此命令可能修改工作区外的文件，且无法自动撤销。',
    )
  }
  const separator = value.indexOf(' · ')
  const toolName = separator >= 0 ? value.slice(0, separator) : value
  const localizedTool = TOOL_LABELS_ZH[toolName]
  if (localizedTool) return `${localizedTool}${separator >= 0 ? value.slice(separator) : ''}`
  return value
}

export function workspaceErrorText(value: string, lang: Lang | string | undefined): string {
  value = value.replace(/^Error invoking remote method '[^']+': (?:Error: )?/u, '')
  if (value === 'task_budget_exhausted' || value === 'budget_exhausted')
    return isChineseWorkspaceLanguage(lang)
      ? '本轮用量已达上限。已完成的步骤仍可查看；请先核对未保存修改，再继续剩余事项。'
      : 'This run reached its execution limit. Completed steps remain available; review unsaved changes before continuing the remaining work.'
  if (value === 'task_effect_requires_review')
    return isChineseWorkspaceLanguage(lang)
      ? '请核对本对话中上一步操作的结果，然后继续。'
      : 'Review the previous action in this conversation to continue.'
  if (/changed on disk after this Agent run|has newer document edits/u.test(value))
    return isChineseWorkspaceLanguage(lang)
      ? '文件已有后续修改，已阻止旧版本覆盖。请先检查新修改。'
      : 'The file has newer changes. Restoring the old version was blocked. Review the newer edits first.'
  if (!isChineseWorkspaceLanguage(lang)) return value
  const exact: Record<string, string> = {
    'The document has unsaved Agent changes. Call office.save before opening it in the editor.':
      '这份文件有尚未保存的更改，请先保存再在编辑器中打开。',
    'The document context expired. Open AI BigDog from the document again.':
      '文档上下文已过期，请从原文档重新打开 AI 大狗。',
    'The original editor closed. Reopen the document before continuing.':
      '原文档已关闭，请重新打开后继续。',
    'The document changed after it was selected. Refresh its context before editing.':
      '文档内容已变化，请从原文档重新打开 AI 大狗，更新上下文后继续。',
    'The document moved or was saved as another file. Refresh its context before continuing.':
      '文档已移动或另存为其他文件，请重新确认当前文档。',
    'The document changed before the edit. Refresh its context.':
      '编辑前文档已变化，请更新上下文后继续。',
    'Use Session, memory or Settings tools to manage application data.':
      '请通过会话、记忆或设置功能管理应用内部数据。',
    'Agent failed.': ZH.agentFailed,
    'Choose a workspace first.': ZH.chooseWorkspaceError,
    'Workspace Agent is unavailable.': '工作区 Agent 当前不可用。',
    'The file could not be opened in BigDogOffice.': '无法在大狗Office 中打开该文件。',
    'AI provider is not configured. Add an API key in Settings.':
      '尚未配置 AI 服务商，请在设置中添加 API Key。',
    'AI model is not configured.': '尚未配置 AI 模型。',
  }
  return exact[value] ?? value
}
