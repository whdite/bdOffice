import type { CodexModelCatalog } from '@genoffice/ai-provider'
import { contextBridge, ipcRenderer, webUtils } from 'electron'
import { COMPUTER_USE_CHANNELS, type ComputerUseApi } from '../shared/computer-use-api'

contextBridge.exposeInMainWorld('aiOfficeComputer', {
  status: () => ipcRenderer.invoke(COMPUTER_USE_CHANNELS.status),
  restart: () => ipcRenderer.invoke(COMPUTER_USE_CHANNELS.restart),
  stop: () => ipcRenderer.invoke(COMPUTER_USE_CHANNELS.stop),
} satisfies ComputerUseApi)
import type { IpcRendererEvent } from 'electron'
import {
  AI_PROVIDERS,
  AI_MEDIA_PROVIDERS,
  AI_SEARCH_PROVIDERS,
  getProviderAdapter,
} from '@genoffice/ai-provider/browser'
import type { AiModelListResponse, AiSettings } from '@genoffice/ai-provider'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import type { UpdateUiState } from '../shared/update-api'
import { normalizeAiPanelPrefs } from '@genoffice/ui/ai-panel-prefs'
import type {
  AccountLoginEvent,
  AccountStatus,
  DefaultAppStatus,
  FolderListing,
  FolderRoot,
  MoveResult,
  AuthorNameSetting,
  HomeApi,
  RecentEntry,
  RecentPage,
  RenameResult,
  UiLanguage,
  FileSearchPage,
  FileSearchRerank,
  FileSearchSettings,
} from '../shared/home-api'
import { HOME_CHANNELS, PROJECT_CHANNELS } from '../shared/home-api'
import { CONNECTION_SETUP_CHANNELS } from '../shared/connection-setup-api'
import { ENVIRONMENT_CHANNELS, type EnvironmentApi } from '../shared/environment-api'
import { REMOTE_ADMIN_CHANNELS, type RemoteAdminApi } from '../shared/mobile-host-api'
import type { ProjectHomeApi, ProjectSummaryEntry, TimelineEntryItem } from '../shared/home-api'
import type { ProjectApi } from '@genoffice/project-store'
import { INTEGRATIONS_CHANNELS } from '../shared/integrations-api'
import type {
  IntegrationsApi,
  IntegrationsStatus,
  LegacySweepResult,
  SkillInstallState,
} from '../shared/integrations-api'
import type { FoundSkill, RecommendedSkill } from '../shared/found-skill'
import type { TabsApi, TabSummary } from '../shared/tabs-api'
import { TABS_CHANNELS } from '../shared/tabs-api'
import type {
  WorkspaceAgentApi,
  WorkspaceAgentEvent,
  WorkspaceFilesChangedEvent,
} from '../shared/workspace-agent-api'
import { WORKSPACE_AGENT_CHANNELS } from '../shared/workspace-agent-api'
import {
  RUNTIME_PROVIDER_CHANNELS,
  isAgentRuntimeSelection,
  isRuntimeProviderId,
  type RuntimeProvidersSnapshot,
} from '../shared/runtime-provider-api'

const UI_LANGUAGES: readonly UiLanguage[] = [
  'zh',
  'en',
  'ja',
  'ko',
  'fr',
  'de',
  'es',
  'th',
  'id',
  'ru',
  'ar',
  'pt',
  'it',
  'pl',
  'cs',
  'nl',
  'ms',
  'he',
  'hi',
  'zh-TW',
  'vi',
  'tr',
]

function isUiLanguage(value: unknown): value is UiLanguage {
  return UI_LANGUAGES.includes(value as UiLanguage)
}

const EMPTY_PAGE: RecentPage = { entries: [], total: 0, totalAll: 0 }

function asRecentPage(result: unknown): RecentPage {
  if (result && typeof result === 'object' && Array.isArray((result as RecentPage).entries)) {
    return result as RecentPage
  }
  return EMPTY_PAGE
}

const EMPTY_SEARCH: FileSearchPage = {
  hits: [],
  total: 0,
  index: { indexed: 0, pending: 0, scanning: false },
}

function asSearchPage(result: unknown): FileSearchPage {
  if (result && typeof result === 'object' && Array.isArray((result as FileSearchPage).hits)) {
    return result as FileSearchPage
  }
  return EMPTY_SEARCH
}

function normalizeDefaultAppStatus(result: unknown): DefaultAppStatus {
  const r = (result ?? {}) as Partial<DefaultAppStatus>
  const state = r.state
  return {
    state: state === 'default' || state === 'other' || state === 'unknown' ? state : 'unsupported',
    others: Array.isArray(r.others) ? r.others.filter((x) => typeof x === 'string') : [],
    manualOnly: r.manualOnly === true,
  }
}

function normalizeAuthorNameSetting(result: unknown): AuthorNameSetting {
  const r = (result ?? {}) as Partial<AuthorNameSetting>
  return {
    value: typeof r.value === 'string' ? r.value : '',
    fallback: typeof r.fallback === 'string' ? r.fallback : '',
  }
}

// a Settings request (menu click, Cmd+,) can land before Home has subscribed
const pendingOpenSettings: Array<{ section: string }> = []
ipcRenderer.on(HOME_CHANNELS.openSettings, (_event, target: { section: string }) => {
  if (ipcRenderer.listenerCount(HOME_CHANNELS.openSettings) === 1) pendingOpenSettings.push(target)
})

const homeApi: HomeApi = {
  async recents(query) {
    return asRecentPage(await ipcRenderer.invoke(HOME_CHANNELS.recents, query))
  },
  async searchFiles(query) {
    return asSearchPage(await ipcRenderer.invoke(HOME_CHANNELS.searchFiles, query))
  },
  async rerankSearch(query) {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.rerankSearch, query)
    return result && typeof result === 'object' && Array.isArray((result as FileSearchRerank).order)
      ? (result as FileSearchRerank)
      : null
  },
  async getFileSearchSettings() {
    return (await ipcRenderer.invoke(HOME_CHANNELS.getFileSearchSettings)) as FileSearchSettings
  },
  async setFileSearchSettings(patch) {
    return (await ipcRenderer.invoke(
      HOME_CHANNELS.setFileSearchSettings,
      patch,
    )) as FileSearchSettings
  },
  async testFileSearchRerank(input) {
    const raw = ((await ipcRenderer.invoke(HOME_CHANNELS.testFileSearchRerank, input)) ?? {}) as {
      ok?: unknown
      error?: unknown
    }
    return raw.ok === true
      ? { ok: true }
      : { ok: false, error: typeof raw.error === 'string' ? raw.error : 'Connection failed' }
  },
  async starred(query) {
    return asRecentPage(await ipcRenderer.invoke(HOME_CHANNELS.starred, query))
  },
  async starredGroups() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.starredGroups)
    return Array.isArray(result) ? result.filter((g): g is string => typeof g === 'string') : []
  },
  async setStarredGroup(paths, group) {
    await ipcRenderer.invoke(HOME_CHANNELS.setStarredGroup, paths, group)
  },
  async statPaths(paths) {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.statPaths, paths)
    return Array.isArray(result) ? (result as RecentEntry[]) : []
  },
  async toggleStar(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.toggleStar, path)
  },
  async openHelp() {
    await ipcRenderer.invoke(HOME_CHANNELS.openHelp)
  },
  // knowledge-base ask panel (handlers live in the shell main, same process
  // that owns the index and the shared ai:stream registration)
  async kbList() {
    return ipcRenderer.invoke('kb:list')
  },
  async skillsList() {
    return ipcRenderer.invoke('ai:skills-list')
  },
  async skillText(name) {
    return ipcRenderer.invoke('ai:skill-text', name)
  },
  async openSkillsSettings() {
    await ipcRenderer.invoke('ai:open-skills-settings')
  },
  async openModelSettings() {
    await ipcRenderer.invoke('ai:open-model-settings')
  },
  async kbSearch(query) {
    return ipcRenderer.invoke('kb:search', query)
  },
  async kbRead(path, offset) {
    return ipcRenderer.invoke('kb:read', { path, offset })
  },
  async kbCreateDocument(request) {
    return ipcRenderer.invoke('kb:create-document', request)
  },
  async kbCliContext(sessionId) {
    return ipcRenderer.invoke('kb:cli-context', sessionId)
  },
  async kbRunCli(sessionId, request) {
    return ipcRenderer.invoke('kb:run-cli', { sessionId, ...request })
  },
  async kbCancelCli(runId) {
    await ipcRenderer.invoke('kb:cancel-cli', runId)
  },
  async pickAttachments() {
    return ipcRenderer.invoke('files:pick')
  },
  async addAttachmentPaths(paths) {
    return ipcRenderer.invoke('files:add', paths)
  },
  async addPastedImage(data, ext) {
    return ipcRenderer.invoke('files:add-pasted-image', data, ext)
  },
  async readAttachment(path, offset, maxChars) {
    return ipcRenderer.invoke('files:read', path, offset, maxChars)
  },
  async readAttachmentImage(path) {
    return ipcRenderer.invoke('files:read-image', path)
  },
  async aiGetSettings() {
    return ipcRenderer.invoke('ai:get-settings')
  },
  async aiStream(request) {
    await ipcRenderer.invoke('ai:stream', request)
  },
  async aiStreamToolResult(reply) {
    await ipcRenderer.invoke('ai:stream-tool-result', reply)
  },
  async aiStreamCancel(requestId) {
    await ipcRenderer.invoke('ai:stream-cancel', requestId)
  },
  onAiStream(handler) {
    const listener = (_event: Electron.IpcRendererEvent, chunk: unknown) => handler(chunk as never)
    ipcRenderer.on('ai:stream-chunk', listener)
    return () => ipcRenderer.removeListener('ai:stream-chunk', listener)
  },
  async aiGskStatus() {
    return ipcRenderer.invoke('ai:gsk-status')
  },
  async openPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.openPath, path)
  },
  async browse() {
    await ipcRenderer.invoke(HOME_CHANNELS.browse)
  },
  async newDoc(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newDoc, opts)
  },
  async newSheet(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newSheet, opts)
  },
  async newSlide(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newSlide, opts)
  },
  async newMarkdown(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newMarkdown, opts)
  },
  async newHtml(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newHtml, opts)
  },
  async newPdf(opts) {
    await ipcRenderer.invoke(HOME_CHANNELS.newPdf, opts)
  },
  async removeRecent(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.removeRecent, paths)
  },
  async unstarPaths(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.unstarPaths, paths)
  },
  async revealPath(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.revealPath, path)
  },
  async renameFile(path, newName) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.renameFile, path, newName)
    return (result ?? { ok: false, error: 'Rename failed' }) as RenameResult
  },
  async duplicateFile(path) {
    if (typeof path !== 'string' || !path) throw new Error('Invalid path.')
    await ipcRenderer.invoke(HOME_CHANNELS.duplicateFile, path)
  },
  async deleteFiles(paths) {
    await ipcRenderer.invoke(HOME_CHANNELS.deleteFiles, paths)
  },
  async folderRoots() {
    return (await ipcRenderer.invoke(HOME_CHANNELS.folderRoots)) as FolderRoot[]
  },
  async addFolderRoot() {
    return (await ipcRenderer.invoke(HOME_CHANNELS.addFolderRoot)) as FolderRoot | null
  },
  async dropFolderRoots(paths) {
    return (await ipcRenderer.invoke(HOME_CHANNELS.dropFolderRoots, paths)) as FolderRoot[]
  },
  registeredWorkspaceFiles: () => ipcRenderer.invoke(HOME_CHANNELS.registeredWorkspaceFiles),
  removeWorkspaceReference: (path) =>
    ipcRenderer.invoke(HOME_CHANNELS.removeWorkspaceReference, path),
  async removeFolderRoot(path) {
    await ipcRenderer.invoke(HOME_CHANNELS.removeFolderRoot, path)
  },
  pathForFile(file) {
    return webUtils.getPathForFile(file)
  },
  async listFolder(dir) {
    return (await ipcRenderer.invoke(HOME_CHANNELS.listFolder, dir)) as FolderListing
  },
  async createFolder(parent, name) {
    return (await ipcRenderer.invoke(HOME_CHANNELS.createFolder, parent, name)) as RenameResult
  },
  async renameFolder(dir, newName) {
    return (await ipcRenderer.invoke(HOME_CHANNELS.renameFolder, dir, newName)) as RenameResult
  },
  async movePaths(paths, targetDir, onConflict) {
    return (await ipcRenderer.invoke(
      HOME_CHANNELS.movePaths,
      paths,
      targetDir,
      onConflict,
    )) as MoveResult
  },
  async deleteFolder(dir) {
    await ipcRenderer.invoke(HOME_CHANNELS.deleteFolder, dir)
  },
  onFolderChanged(handler) {
    const listener = (_event: IpcRendererEvent, dirs: unknown) => {
      if (Array.isArray(dirs)) handler(dirs.filter((d): d is string => typeof d === 'string'))
    }
    ipcRenderer.on(HOME_CHANNELS.folderChanged, listener)
    return () => ipcRenderer.removeListener(HOME_CHANNELS.folderChanged, listener)
  },
  async openTrash() {
    await ipcRenderer.invoke(HOME_CHANNELS.openTrash)
  },
  async getLanguage() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getLanguage)
    return isUiLanguage(result) ? result : 'zh'
  },
  async setLanguage(lang) {
    if (!isUiLanguage(lang)) throw new Error('Invalid language.')
    await ipcRenderer.invoke(HOME_CHANNELS.setLanguage, lang)
  },
  async getUpdateChannel() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getUpdateChannel)
    return result === 'beta' ? 'beta' : 'stable'
  },
  async setUpdateChannel(channel) {
    // validated inline: a runtime import from ../shared/update-api would be
    // shared with the update.ts preload entry and get split into a chunk,
    // which sandboxed preload scripts cannot load (window.aiOffice would
    // silently disappear). Preload entries must stay single-file bundles.
    if (channel !== 'stable' && channel !== 'beta') throw new Error('Invalid update channel.')
    await ipcRenderer.invoke(HOME_CHANNELS.setUpdateChannel, channel)
  },
  async getAutoUpdateCheck() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getAutoUpdateCheck)
    return result !== false
  },
  async setAutoUpdateCheck(on) {
    if (typeof on !== 'boolean') throw new Error('Invalid update check setting.')
    await ipcRenderer.invoke(HOME_CHANNELS.setAutoUpdateCheck, on)
  },
  async getSilentInstallUpdate() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getSilentInstallUpdate)
    return result === true
  },
  async setSilentInstallUpdate(on) {
    if (typeof on !== 'boolean') throw new Error('Invalid silent install setting.')
    await ipcRenderer.invoke(HOME_CHANNELS.setSilentInstallUpdate, on)
  },
  async accountStatus() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.accountStatus)
    return (result ?? { loggedIn: false }) as AccountStatus
  },
  async accountLogin() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.accountLogin)
    return result === true
  },
  onOpenSettings(handler) {
    const listener = (_event: IpcRendererEvent, target: { section: string }) => handler(target)
    ipcRenderer.on(HOME_CHANNELS.openSettings, listener)
    for (const target of pendingOpenSettings.splice(0)) handler(target)
    return () => ipcRenderer.removeListener(HOME_CHANNELS.openSettings, listener)
  },
  onAccountLogin(handler) {
    const listener = (_event: IpcRendererEvent, ev: AccountLoginEvent) => handler(ev)
    ipcRenderer.on(HOME_CHANNELS.accountLoginEvent, listener)
    return () => ipcRenderer.removeListener(HOME_CHANNELS.accountLoginEvent, listener)
  },
  async openLoginUrl() {
    await ipcRenderer.invoke(HOME_CHANNELS.accountLoginOpenUrl)
  },
  async accountLogout() {
    await ipcRenderer.invoke(HOME_CHANNELS.accountLogout)
  },
  async getAppVersion() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getAppVersion)
    return typeof result === 'string' ? result : ''
  },
  async getUpdateState() {
    const result: unknown = await ipcRenderer.invoke('update:get-state')
    return (result as UpdateUiState | null) ?? null
  },
  async openUpdateDialog() {
    const result: unknown = await ipcRenderer.invoke('update:open-for-update')
    return result === true
  },
  onUpdateStateChanged(handler: (state: UpdateUiState) => void) {
    const listener = (_e: IpcRendererEvent, state: UpdateUiState) => handler(state)
    ipcRenderer.on('update:state-changed', listener)
    return () => ipcRenderer.removeListener('update:state-changed', listener)
  },
  async onboardingSeen() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.onboardingSeen)
    return result === true
  },
  async setOnboardingSeen(choice) {
    const result: unknown = await ipcRenderer.invoke(
      HOME_CHANNELS.setOnboardingSeen,
      ...(choice === undefined ? [] : [choice]),
    )
    return result === true
  },
  connectionSetup: {
    inspect: () => ipcRenderer.invoke(CONNECTION_SETUP_CHANNELS.inspect),
    test: (draft) => ipcRenderer.invoke(CONNECTION_SETUP_CHANNELS.test, draft),
    commit: (input) => ipcRenderer.invoke(CONNECTION_SETUP_CHANNELS.commit, input),
    cancel: (requestId) => ipcRenderer.invoke(CONNECTION_SETUP_CHANNELS.cancel, requestId),
    openHelp: (page) => ipcRenderer.invoke(CONNECTION_SETUP_CHANNELS.openHelp, page),
  },
  async getTheme() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getTheme)
    return result === 'dark' || result === 'light' ? result : 'system'
  },
  async setTheme(theme) {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'system')
      throw new Error('Invalid theme.')
    await ipcRenderer.invoke(HOME_CHANNELS.setTheme, theme)
  },
  async getDocumentTheme() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getDocumentTheme)
    return result === 'dark' || result === 'light' ? result : 'follow'
  },
  async setDocumentTheme(theme) {
    if (theme !== 'light' && theme !== 'dark' && theme !== 'follow')
      throw new Error('Invalid document theme.')
    await ipcRenderer.invoke(HOME_CHANNELS.setDocumentTheme, theme)
  },
  async getAutoSaveDefault() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getAutoSaveDefault)
    const r = result as { on?: unknown; updatedAt?: unknown } | null
    return {
      on: r?.on === true,
      updatedAt: typeof r?.updatedAt === 'number' ? r.updatedAt : 0,
    }
  },
  async getTabSleepEnabled() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getTabSleepEnabled)
    return result !== false
  },
  async setTabSleepEnabled(on) {
    if (typeof on !== 'boolean') throw new Error('Invalid tab sleep setting.')
    await ipcRenderer.invoke(HOME_CHANNELS.setTabSleepEnabled, on)
  },
  async getRestoreLastSession() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getRestoreLastSession)
    return result === true
  },
  async setRestoreLastSession(on) {
    if (typeof on !== 'boolean') throw new Error('Invalid restore session setting.')
    await ipcRenderer.invoke(HOME_CHANNELS.setRestoreLastSession, on)
  },
  async setAutoSaveDefault(on) {
    if (typeof on !== 'boolean') throw new Error('Invalid AutoSave default.')
    await ipcRenderer.invoke(HOME_CHANNELS.setAutoSaveDefault, on)
  },
  async getMcpStatus() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getMcpStatus)
    const r = result as {
      running?: unknown
      enabled?: unknown
      port?: unknown
      background?: unknown
      logging?: unknown
      url?: unknown
      capabilities?: unknown
      error?: unknown
    } | null
    return {
      running: r?.running === true,
      enabled: r?.enabled === true,
      port: typeof r?.port === 'number' ? r.port : 3093,
      background: r?.background === true,
      logging: r?.logging === true,
      url: typeof r?.url === 'string' ? r.url : null,
      capabilities: Array.isArray(r?.capabilities)
        ? r.capabilities.filter((c): c is string => typeof c === 'string')
        : ['docs'],
      ...(typeof r?.error === 'string' ? { error: r.error } : {}),
    }
  },
  async setMcpSettings(patch: {
    enabled?: boolean
    port?: number
    background?: boolean
    logging?: boolean
  }) {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.setMcpSettings, patch)
    const r = result as {
      running?: unknown
      enabled?: unknown
      port?: unknown
      background?: unknown
      logging?: unknown
      url?: unknown
      capabilities?: unknown
      error?: unknown
    } | null
    return {
      running: r?.running === true,
      enabled: r?.enabled === true,
      port: typeof r?.port === 'number' ? r.port : 3093,
      background: r?.background === true,
      logging: r?.logging === true,
      url: typeof r?.url === 'string' ? r.url : null,
      capabilities: Array.isArray(r?.capabilities)
        ? r.capabilities.filter((c): c is string => typeof c === 'string')
        : ['docs'],
      ...(typeof r?.error === 'string' ? { error: r.error } : {}),
    }
  },
  async getMcpLogs() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getMcpLogs)
    return Array.isArray(result) ? result.filter((l): l is string => typeof l === 'string') : []
  },
  async clearMcpLogs() {
    await ipcRenderer.invoke(HOME_CHANNELS.clearMcpLogs)
  },
  async openMcpLogFile() {
    await ipcRenderer.invoke(HOME_CHANNELS.openMcpLogFile)
  },
  async getAnalyticsEnabled() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getAnalyticsEnabled)
    return result !== false
  },
  async setAnalyticsEnabled(enabled) {
    if (typeof enabled !== 'boolean') throw new Error('Invalid analytics consent.')
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.setAnalyticsEnabled, enabled)
    return result === true
  },
  async getLegacyImportPrompt() {
    return (await ipcRenderer.invoke(HOME_CHANNELS.getLegacyImportPrompt)) !== false
  },
  async setLegacyImportPrompt(prompt) {
    return (await ipcRenderer.invoke(HOME_CHANNELS.setLegacyImportPrompt, prompt)) !== false
  },
  async getAuthorName() {
    return normalizeAuthorNameSetting(await ipcRenderer.invoke(HOME_CHANNELS.getAuthorName))
  },
  async setAuthorName(value) {
    return normalizeAuthorNameSetting(await ipcRenderer.invoke(HOME_CHANNELS.setAuthorName, value))
  },
  async getAiPanelPrefs() {
    return normalizeAiPanelPrefs(await ipcRenderer.invoke(HOME_CHANNELS.getAiPanelPrefs))
  },
  async setAiPanelPrefs(patch) {
    return normalizeAiPanelPrefs(await ipcRenderer.invoke(HOME_CHANNELS.setAiPanelPrefs, patch))
  },
  async getDefaultSaveDir() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.getDefaultSaveDir)
    return typeof result === 'string' ? result : ''
  },
  async getDefaultAppStatus() {
    return normalizeDefaultAppStatus(await ipcRenderer.invoke(HOME_CHANNELS.getDefaultAppStatus))
  },
  async setDefaultApp() {
    return normalizeDefaultAppStatus(await ipcRenderer.invoke(HOME_CHANNELS.setDefaultApp))
  },
  async pickDefaultSaveDir() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.pickDefaultSaveDir)
    return typeof result === 'string' && result ? result : null
  },
  onThemeChanged(handler) {
    const listener = (_event: Electron.IpcRendererEvent, theme: unknown) => {
      if (theme === 'light' || theme === 'dark' || theme === 'system') handler(theme)
    }
    ipcRenderer.on('app:theme-changed', listener)
    return () => ipcRenderer.removeListener('app:theme-changed', listener)
  },
  onLanguageChanged(handler) {
    const listener = (_event: Electron.IpcRendererEvent, language: unknown) => {
      if (isUiLanguage(language)) handler(language)
    }
    ipcRenderer.on('app:language-changed', listener)
    return () => ipcRenderer.removeListener('app:language-changed', listener)
  },
  onDocumentThemeChanged(handler) {
    const listener = (_event: Electron.IpcRendererEvent, theme: unknown) => {
      if (theme === 'light' || theme === 'dark' || theme === 'follow') handler(theme)
    }
    ipcRenderer.on('app:document-theme-changed', listener)
    return () => ipcRenderer.removeListener('app:document-theme-changed', listener)
  },
  async openGenTeam() {
    await ipcRenderer.invoke(HOME_CHANNELS.openGenTeam)
  },
  async checkBigDogRelease() {
    return ipcRenderer.invoke(HOME_CHANNELS.checkBigDogRelease)
  },
  async getBigDogReleaseStatus() {
    return ipcRenderer.invoke(HOME_CHANNELS.getBigDogReleaseStatus)
  },
  onBigDogReleaseChanged(handler) {
    const listener = (_event: Electron.IpcRendererEvent, status: Parameters<typeof handler>[0]) =>
      handler(status)
    ipcRenderer.on(HOME_CHANNELS.bigDogReleaseChanged, listener)
    return () => ipcRenderer.removeListener(HOME_CHANNELS.bigDogReleaseChanged, listener)
  },
  async dismissBigDogRelease(version) {
    await ipcRenderer.invoke(HOME_CHANNELS.dismissBigDogRelease, version)
  },
  async openBigDogReleases(version) {
    await ipcRenderer.invoke(HOME_CHANNELS.openBigDogReleases, version)
  },
  async openBigDogWebsite() {
    await ipcRenderer.invoke(HOME_CHANNELS.openBigDogWebsite)
  },
  async openGitHubRepo() {
    await ipcRenderer.invoke(HOME_CHANNELS.openGitHubRepo)
  },
  async openRemoteGuide() {
    await ipcRenderer.invoke(HOME_CHANNELS.openRemoteGuide)
  },
  async githubStars() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.githubStars)
    return typeof result === 'number' && Number.isFinite(result) ? result : null
  },
  async starPromptShouldShow() {
    const result: unknown = await ipcRenderer.invoke(HOME_CHANNELS.starPromptShouldShow)
    const raw = (result ?? {}) as { show?: unknown; docOpens?: unknown }
    return {
      show: raw.show === true,
      docOpens:
        typeof raw.docOpens === 'number' && Number.isFinite(raw.docOpens) ? raw.docOpens : 0,
    }
  },
  async starPromptAction(action) {
    if (action !== 'starred' && action !== 'later') throw new Error('Invalid star prompt action.')
    await ipcRenderer.invoke(HOME_CHANNELS.starPromptAction, action)
  },
  // AI settings channels are registered once by the shell's aggregated docs handlers
  onAiSettingsChanged(handler) {
    const listener = () => handler()
    ipcRenderer.on('ai:settings-changed', listener)
    ipcRenderer.on('ai:connection-changed', listener)
    return () => {
      ipcRenderer.removeListener('ai:settings-changed', listener)
      ipcRenderer.removeListener('ai:connection-changed', listener)
    }
  },
  async getAiSettings() {
    return (await ipcRenderer.invoke('ai:get-settings')) as AiSettings
  },
  async setAiSettings(settings) {
    await ipcRenderer.invoke('ai:set-settings', settings)
  },
  async listAiModels(request) {
    return (await ipcRenderer.invoke('ai:list-models', request)) as AiModelListResponse
  },
  async cancelAiModelList(requestId) {
    await ipcRenderer.invoke('ai:list-models-cancel', requestId)
  },
  getAiProviders() {
    return AI_PROVIDERS.map((meta) => {
      const defaultBaseUrl = meta.needsBaseUrl
        ? ''
        : getProviderAdapter(meta.id).resolveEndpoint({
            apiKey: '',
            model: meta.defaultModel,
          }).baseUrl
      return { ...meta, defaultBaseUrl }
    })
  },
  async getCodexModels(cliPath) {
    return (await ipcRenderer.invoke('ai:codex-models', cliPath)) as CodexModelCatalog
  },
  async getCustomModels(baseUrl, apiKey, api) {
    return (await ipcRenderer.invoke('ai:custom-models', {
      baseUrl,
      apiKey,
      ...(api ? { api } : {}),
    })) as CodexModelCatalog
  },
  async testAiSettings(settings) {
    const result: unknown = await ipcRenderer.invoke('ai:chat', {
      settings,
      system: 'You are a connectivity test. Reply with the single word OK.',
      user: 'ping',
    })
    const raw = (result ?? {}) as { ok?: unknown; error?: unknown }
    return raw.ok === true
      ? { ok: true }
      : { ok: false, error: typeof raw.error === 'string' ? raw.error : 'Connection failed' }
  },
  async getRuntimeProviders() {
    return (await ipcRenderer.invoke(
      RUNTIME_PROVIDER_CHANNELS.snapshot,
    )) as RuntimeProvidersSnapshot
  },
  async connectRuntimeProvider(id) {
    if (!isRuntimeProviderId(id)) throw new Error('Invalid Runtime Provider.')
    return (await ipcRenderer.invoke(
      RUNTIME_PROVIDER_CHANNELS.connect,
      id,
    )) as RuntimeProvidersSnapshot
  },
  async installRuntimeProvider(id) {
    if (!isRuntimeProviderId(id)) throw new Error('Invalid Runtime Provider.')
    return (await ipcRenderer.invoke(
      RUNTIME_PROVIDER_CHANNELS.install,
      id,
    )) as RuntimeProvidersSnapshot
  },
  async loginRuntimeProvider(id) {
    if (!isRuntimeProviderId(id)) throw new Error('Invalid Runtime Provider.')
    return (await ipcRenderer.invoke(
      RUNTIME_PROVIDER_CHANNELS.login,
      id,
    )) as RuntimeProvidersSnapshot
  },
  async selectAgentRuntime(selection) {
    if (!isAgentRuntimeSelection(selection)) throw new Error('Invalid Agent runtime selection.')
    return (await ipcRenderer.invoke(
      RUNTIME_PROVIDER_CHANNELS.select,
      selection,
    )) as RuntimeProvidersSnapshot
  },
  async logoutRuntimeProvider(id) {
    if (!isRuntimeProviderId(id)) throw new Error('Invalid Runtime Provider.')
    return (await ipcRenderer.invoke(
      RUNTIME_PROVIDER_CHANNELS.logout,
      id,
    )) as RuntimeProvidersSnapshot
  },
  getAiMediaProviders() {
    return AI_MEDIA_PROVIDERS
  },
  getAiSearchProviders() {
    return AI_SEARCH_PROVIDERS
  },
  async testAiSearchSettings(input) {
    const raw = ((await ipcRenderer.invoke('ai:search-test', input)) ?? {}) as {
      ok?: unknown
      error?: unknown
    }
    return raw.ok === true
      ? { ok: true }
      : { ok: false, error: typeof raw.error === 'string' ? raw.error : 'Connection failed' }
  },
  async testAiMediaSettings(input) {
    const raw = ((await ipcRenderer.invoke('ai:media-test', input)) ?? {}) as {
      ok?: unknown
      error?: unknown
    }
    return raw.ok === true
      ? { ok: true }
      : { ok: false, error: typeof raw.error === 'string' ? raw.error : 'Connection failed' }
  },
}

contextBridge.exposeInMainWorld('aiOffice', homeApi)
const environmentApi: EnvironmentApi = {
  check: () => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.check),
  plan: (reportId) => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.plan, reportId),
  repair: (input) => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.repair, input),
  task: (id) => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.task, id),
  cancelTask: (id) => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.cancelTask, id),
  recentRepairs: () => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.recentRepairs),
  askBigDog: () => ipcRenderer.invoke(ENVIRONMENT_CHANNELS.askBigDog),
}
contextBridge.exposeInMainWorld('aiOfficeEnvironment', environmentApi)
const remoteAdminApi: RemoteAdminApi = {
  status: () => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.status),
  save: (input) => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.save, input),
  start: (input) => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.start, input),
  stop: () => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.stop),
  invite: () => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.invite),
  saveQr: () => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.saveQr),
  revoke: (id) => ipcRenderer.invoke(REMOTE_ADMIN_CHANNELS.revoke, id),
}
contextBridge.exposeInMainWorld('aiOfficeRemote', remoteAdminApi)
const assistantTasksApi: AssistantTasksApi = {
  feed: (sessionId, afterCursor) =>
    ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.feed, sessionId, afterCursor),
  evidence: (id) => ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.evidence, id),
  control: (input) => ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.control, input),
  answer: (input) => ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.answer, input),
  acknowledge: (id) => ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.acknowledge, id),
  openSession: (id) => ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.openSession, id),
  openArtifact: (sessionId, artifactId) =>
    ipcRenderer.invoke(ASSISTANT_TASK_CHANNELS.openArtifact, sessionId, artifactId),
}
contextBridge.exposeInMainWorld('aiOfficeTasks', assistantTasksApi)

const workspaceAgentApi: WorkspaceAgentApi = {
  prepareContext: () => ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.prepareContext),
  openHome: (id) => ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.openHome, id),
  clearDocumentContext: () => ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.clearDocumentContext),
  getContextSelection: (id) => ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.contextSelection, id),
  setContextSelection: (id, selection) =>
    ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.setContextSelection, id, selection),
  pickContextFiles: (id) => ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.pickContextFiles, id),
  attachContextFiles: (id, paths) =>
    ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.attachContextFiles, id, paths),
  previewContext: (id) => ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.previewContext, id),
  async getWorkspace() {
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.getWorkspace)) as Awaited<
      ReturnType<WorkspaceAgentApi['getWorkspace']>
    >
  },
  async getContext() {
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.getContext)) as Awaited<
      ReturnType<WorkspaceAgentApi['getContext']>
    >
  },
  async openWorkspace() {
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.openWorkspace)) as Awaited<
      ReturnType<WorkspaceAgentApi['openWorkspace']>
    >
  },
  async setWorkspace(path) {
    if (typeof path !== 'string' || !path.trim()) throw new Error('Workspace path is required.')
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.setWorkspace, path)) as Awaited<
      ReturnType<WorkspaceAgentApi['setWorkspace']>
    >
  },
  async listProjectWorkspaces() {
    const result: unknown = await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.listProjectWorkspaces)
    return Array.isArray(result) ? result : []
  },
  async setProjectWorkspace(projectId) {
    if (typeof projectId !== 'string' || !projectId.trim())
      throw new Error('Project workspace id is required.')
    return (await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.setProjectWorkspace,
      projectId,
    )) as Awaited<ReturnType<WorkspaceAgentApi['setProjectWorkspace']>>
  },
  async listFiles(request) {
    const result: unknown = await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.listFiles, request)
    return Array.isArray(result) ? result : []
  },
  async getFile(path, maxChars) {
    if (typeof path !== 'string' || !path.trim()) throw new Error('File path is required.')
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.getFile, path, maxChars)) as Awaited<
      ReturnType<WorkspaceAgentApi['getFile']>
    >
  },
  async searchFiles(request) {
    const result: unknown = await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.searchFiles, request)
    return Array.isArray(result) ? result : []
  },
  async getFileMetadata(path) {
    if (typeof path !== 'string' || !path.trim()) throw new Error('File path is required.')
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.fileMetadata, path)) as Awaited<
      ReturnType<WorkspaceAgentApi['getFileMetadata']>
    >
  },
  async openFile(path) {
    if (typeof path !== 'string' || !path.trim()) throw new Error('File path is required.')
    await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.openFile, path)
  },
  async createSession(workspaceId) {
    return (await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.createSession,
      workspaceId,
    )) as Awaited<ReturnType<WorkspaceAgentApi['createSession']>>
  },
  async resumeSession(sessionId) {
    if (!sessionId) throw new Error('Invalid session id.')
    return (await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.resumeSession, sessionId)) as Awaited<
      ReturnType<WorkspaceAgentApi['resumeSession']>
    >
  },
  async listSessions(workspaceId) {
    const result: unknown = await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.listSessions,
      workspaceId,
    )
    return Array.isArray(result) ? result : []
  },
  async getSessionMessages(sessionId) {
    if (!sessionId) throw new Error('Invalid session id.')
    const result: unknown = await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.sessionMessages,
      sessionId,
    )
    return Array.isArray(result) ? result : []
  },
  async getSessionHistoryPrefix(sessionId) {
    if (!sessionId) throw new Error('Invalid session id.')
    const result: unknown = await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.sessionHistoryPrefix,
      sessionId,
    )
    return Array.isArray(result) ? result : []
  },
  async getSessionEvents(sessionId, afterSequence) {
    if (!sessionId) throw new Error('Invalid session id.')
    const result: unknown = await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.sessionEvents,
      sessionId,
      afterSequence,
    )
    return Array.isArray(result) ? (result as WorkspaceAgentEvent[]) : []
  },
  async renameSession(sessionId, title) {
    if (!sessionId || typeof title !== 'string' || !title.trim())
      throw new Error('Invalid session rename.')
    return (await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.renameSession,
      sessionId,
      title,
    )) as Awaited<ReturnType<WorkspaceAgentApi['renameSession']>>
  },
  async deleteSession(sessionId) {
    if (!sessionId) throw new Error('Invalid session id.')
    await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.deleteSession, sessionId)
  },
  async send(sessionId, message) {
    if (!sessionId) return
    if (!message || typeof message !== 'object' || typeof message.text !== 'string')
      throw new Error('Invalid Agent input.')
    if (!message.text.trim()) return
    await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.send, sessionId, message)
  },
  async steer(sessionId, message) {
    if (!sessionId) return
    if (!message || typeof message !== 'object' || typeof message.text !== 'string')
      throw new Error('Invalid Agent input.')
    if (!message.text.trim()) return
    await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.steer, sessionId, message)
  },
  async cancel(sessionId) {
    if (!sessionId) return
    await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.cancel, sessionId)
  },
  async getApprovalMode() {
    const value: unknown = await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.getApprovalMode)
    return value === 'auto' ? 'auto' : 'per-action'
  },
  async setApprovalMode(mode) {
    if (mode !== 'auto' && mode !== 'per-action') throw new Error('Invalid Agent approval mode.')
    const value: unknown = await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.setApprovalMode, mode)
    return value === 'auto' ? 'auto' : 'per-action'
  },
  onApprovalModeChanged(handler) {
    const listener = (_event: unknown, mode: unknown) => {
      if (mode === 'auto' || mode === 'per-action') handler(mode)
    }
    ipcRenderer.on('agent:editor-approval-mode-changed', listener)
    return () => ipcRenderer.removeListener('agent:editor-approval-mode-changed', listener)
  },
  async resolvePermission(sessionId, requestId, decision) {
    if (!sessionId || !requestId) throw new Error('Invalid permission request.')
    if (decision !== 'allow-once' && decision !== 'allow-round' && decision !== 'deny')
      throw new Error('Invalid permission decision.')
    await ipcRenderer.invoke(
      WORKSPACE_AGENT_CHANNELS.resolvePermission,
      sessionId,
      requestId,
      decision,
    )
  },
  async listRunReviews(sessionId, cursor) {
    if (!sessionId) throw new Error('Invalid session id.')
    return ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.runReviews, sessionId, cursor)
  },
  async undoRun(sessionId, runId) {
    if (!sessionId || !runId) throw new Error('Invalid Agent run rollback request.')
    return ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.undoRun, sessionId, runId)
  },
  async openWordReview(sessionId, runId, path) {
    await ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.openWordReview, sessionId, runId, path)
  },
  subscribe(sessionId, handler) {
    let disposed = false
    const listener = (_event: IpcRendererEvent, value: unknown) => {
      if (!value || typeof value !== 'object') return
      const event = value as WorkspaceAgentEvent
      if (event.sessionId === sessionId) handler(event)
    }
    ipcRenderer.on(WORKSPACE_AGENT_CHANNELS.event, listener)
    void ipcRenderer
      .invoke(WORKSPACE_AGENT_CHANNELS.subscribe, sessionId)
      .then(() => {
        if (disposed)
          void ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.unsubscribe, sessionId).catch(() => {})
      })
      .catch(() => {
        ipcRenderer.removeListener(WORKSPACE_AGENT_CHANNELS.event, listener)
      })
    return () => {
      disposed = true
      ipcRenderer.removeListener(WORKSPACE_AGENT_CHANNELS.event, listener)
      void ipcRenderer.invoke(WORKSPACE_AGENT_CHANNELS.unsubscribe, sessionId).catch(() => {})
    }
  },
  subscribeFiles(workspaceId, handler) {
    let disposed = false
    const listener = (_event: IpcRendererEvent, value: unknown) => {
      if (!value || typeof value !== 'object') return
      const changed = value as WorkspaceFilesChangedEvent
      if (changed.workspaceId === workspaceId) handler(changed)
    }
    ipcRenderer.on(WORKSPACE_AGENT_CHANNELS.filesChanged, listener)
    void ipcRenderer
      .invoke(WORKSPACE_AGENT_CHANNELS.subscribeFiles, workspaceId)
      .then(() => {
        if (disposed)
          void ipcRenderer
            .invoke(WORKSPACE_AGENT_CHANNELS.unsubscribeFiles, workspaceId)
            .catch(() => {})
      })
      .catch(() => {
        ipcRenderer.removeListener(WORKSPACE_AGENT_CHANNELS.filesChanged, listener)
      })
    return () => {
      disposed = true
      ipcRenderer.removeListener(WORKSPACE_AGENT_CHANNELS.filesChanged, listener)
      void ipcRenderer
        .invoke(WORKSPACE_AGENT_CHANNELS.unsubscribeFiles, workspaceId)
        .catch(() => {})
    }
  },
}

contextBridge.exposeInMainWorld('aiOfficeWorkspaceAgent', workspaceAgentApi)
// Keep the neutral API name from the Workspace Agent contract as a small
// compatibility alias; the shell's existing naming convention remains the
// primary `aiOfficeWorkspaceAgent` surface used by the renderer.
contextBridge.exposeInMainWorld('workspaceAgent', workspaceAgentApi)

const projectHomeApi: ProjectHomeApi = {
  async listProjects() {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.list)
    return Array.isArray(result) ? (result as ProjectSummaryEntry[]) : []
  },
  async listFiles(projectId) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.files, { projectId })
    return Array.isArray(result)
      ? result.filter((path): path is string => typeof path === 'string')
      : []
  },
  async createProject(name) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.create, { name })
    return result as ProjectSummaryEntry
  },
  async renameProject(id, name) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.rename, { id, name })
  },
  async deleteProject(id) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.delete, { id })
  },
  async moveFile(filePath, projectId) {
    await ipcRenderer.invoke(PROJECT_CHANNELS.moveFile, { filePath, projectId })
  },
  async getTimeline(projectId, limit) {
    const result: unknown = await ipcRenderer.invoke(PROJECT_CHANNELS.timeline, {
      projectId,
      limit,
    })
    return Array.isArray(result) ? (result as TimelineEntryItem[]) : []
  },
}

contextBridge.exposeInMainWorld('aiOfficeProject', projectHomeApi)

/**
 * Model file naming.
 *
 * The same channel the editor apps expose under this name, so the shared
 * client (`@genoffice/ui`) finds one bridge wherever it runs.
 *
 * A global of its own rather than part of `aiOffice`: `aiOffice` is the
 * shell's home API, and an editor window that reused the name would shadow it
 * there — `window.aiOffice.setTheme` would vanish with it.
 */
const aiNamingApi = {
  suggestFileName(input: { content: string }) {
    return ipcRenderer.invoke('ai:suggest-file-name', input)
  },
}

contextBridge.exposeInMainWorld('aiNamingApi', aiNamingApi)

// The same project:* channels the editor preloads use; the home screen's
// knowledge-base ask panel keeps its thread through them. The handlers are
// registered once by the docs main process in shell mode (registerProjectIpc).
const projectApi: ProjectApi = {
  pendingChat: (consume?: boolean) => ipcRenderer.invoke('project:pendingChat', consume),
  onChatCancelled: (handler: (source: { projectId?: string; chatId?: string }) => void) => {
    const listener = (
      _event: Electron.IpcRendererEvent,
      source: { projectId?: string; chatId?: string },
    ) => handler(source)
    ipcRenderer.on('project:cancel-chat', listener)
    return () => ipcRenderer.removeListener('project:cancel-chat', listener)
  },
  loadContext: (args) => ipcRenderer.invoke('project:loadContext', args),
  saveContext: (args) => ipcRenderer.invoke('project:saveContext', args),
  newConversation: (args) => ipcRenderer.invoke('project:newConversation', args),
  onChatActivated: (handler) => {
    const listener = (
      _event: Electron.IpcRendererEvent,
      ids: import('@genoffice/project-store').ResolveChatResult,
    ) => handler(ids)
    ipcRenderer.on('project:chat-activated', listener)
    return () => ipcRenderer.removeListener('project:chat-activated', listener)
  },
  resolveChat: (args) => ipcRenderer.invoke('project:resolveChat', args),
  appendChat: (args) => ipcRenderer.invoke('project:appendChat', args),
  loadChat: (args) => ipcRenderer.invoke('project:loadChat', args),
  rebindChat: (args) => ipcRenderer.invoke('project:rebindChat', args),
  listChats: (args) => ipcRenderer.invoke('project:listChats', args),
  deleteChat: (args) => ipcRenderer.invoke('project:deleteChat', args),
}

contextBridge.exposeInMainWorld('projectApi', projectApi)

const integrationsApi: IntegrationsApi = {
  installCli: () => ipcRenderer.invoke(INTEGRATIONS_CHANNELS.installCli),
  async status() {
    return (await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.status)) as IntegrationsStatus
  },
  async installSkill(target) {
    return (await ipcRenderer.invoke(
      INTEGRATIONS_CHANNELS.installSkill,
      target,
    )) as SkillInstallState
  },
  async uninstallSkill(agentId) {
    return (await ipcRenderer.invoke(
      INTEGRATIONS_CHANNELS.uninstallSkill,
      agentId,
    )) as SkillInstallState
  },
  async sweepLegacySkillCopies() {
    return (await ipcRenderer.invoke(
      INTEGRATIONS_CHANNELS.sweepLegacySkillCopies,
    )) as LegacySweepResult
  },
  async pickSkillDir(title) {
    const r: unknown = await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.pickSkillDir, title)
    return typeof r === 'string' ? r : null
  },
  async saveSkillZip(title) {
    const r: unknown = await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.saveSkillZip, title)
    return typeof r === 'string' ? r : null
  },
  async copyText(text) {
    await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.copyText, text)
  },
  async listSkills() {
    return (await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.listSkills)) as FoundSkill[]
  },
  async skillBody(path) {
    return (await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.skillBody, path)) as string
  },
  async importSkillFolder(title) {
    const r: unknown = await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.importSkillFolder, title)
    return (r as FoundSkill | null) ?? null
  },
  async importSkillZip(title) {
    const r: unknown = await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.importSkillZip, title)
    return (r as FoundSkill | null) ?? null
  },
  async listRecommendedSkills() {
    return (await ipcRenderer.invoke(
      INTEGRATIONS_CHANNELS.listRecommendedSkills,
    )) as RecommendedSkill[]
  },
  async removeSkill(name) {
    await ipcRenderer.invoke(INTEGRATIONS_CHANNELS.removeSkill, name)
  },
  async installRecommendedSkill(name) {
    return (await ipcRenderer.invoke(
      INTEGRATIONS_CHANNELS.installRecommendedSkill,
      name,
    )) as FoundSkill
  },
}
contextBridge.exposeInMainWorld('aiOfficeIntegrations', integrationsApi)

const tabsApi: TabsApi = {
  async list() {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.list)
    return Array.isArray(result) ? (result as TabSummary[]) : []
  },
  async activate(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.activate, id)
  },
  async close(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.close, id)
  },
  async showMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showMenu, x, y)
  },
  async showNewMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showNewMenu, x, y)
  },
  async showTabMenu(id, x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showTabMenu, id, x, y)
  },
  async detach(id) {
    await ipcRenderer.invoke(TABS_CHANNELS.detach, id)
  },
  async tearOff(id, screenX, screenY) {
    const result: unknown = await ipcRenderer.invoke(TABS_CHANNELS.tearOff, id, screenX, screenY)
    return result === true
  },
  dragTornWindow(screenX, screenY) {
    ipcRenderer.send(TABS_CHANNELS.dragTornWindow, screenX, screenY)
  },
  async dockTornWindow(index) {
    await ipcRenderer.invoke(TABS_CHANNELS.dockTornWindow, index)
  },
  async endTornDrag() {
    await ipcRenderer.invoke(TABS_CHANNELS.endTornDrag)
  },
  onDockPreview(handler) {
    const listener = (_event: IpcRendererEvent, preview: { x: number } | null) =>
      handler(preview && typeof preview.x === 'number' ? { x: preview.x } : null)
    ipcRenderer.on(TABS_CHANNELS.dockPreview, listener)
    return () => ipcRenderer.removeListener(TABS_CHANNELS.dockPreview, listener)
  },
  reportDockIndex(index) {
    ipcRenderer.send(TABS_CHANNELS.dockIndex, index)
  },
  async showAppMenu(x, y) {
    await ipcRenderer.invoke(TABS_CHANNELS.showAppMenu, x, y)
  },
  async reorder(id, toIndex) {
    await ipcRenderer.invoke(TABS_CHANNELS.reorder, id, toIndex)
  },
  onChanged(handler) {
    const listener = (_event: IpcRendererEvent, tabs: TabSummary[]) => handler(tabs)
    ipcRenderer.on(TABS_CHANNELS.changed, listener)
    return () => ipcRenderer.removeListener(TABS_CHANNELS.changed, listener)
  },
  notifyChromePressed() {
    ipcRenderer.send(TABS_CHANNELS.chromePressed)
  },
  onChromePressed(handler) {
    const listener = () => handler()
    ipcRenderer.on('app:chrome-pressed', listener)
    return () => ipcRenderer.removeListener('app:chrome-pressed', listener)
  },
  async setWorkspaceChat(open, width) {
    await ipcRenderer.invoke(TABS_CHANNELS.workspaceChat, open, width)
  },
  onStartRename(handler) {
    const listener = (_event: IpcRendererEvent, id: unknown) => {
      if (typeof id === 'string') handler(id)
    }
    ipcRenderer.on(TABS_CHANNELS.startRename, listener)
    return () => ipcRenderer.removeListener(TABS_CHANNELS.startRename, listener)
  },
}

contextBridge.exposeInMainWorld('aiOfficeTabs', tabsApi)

// open documents dragged from the OS anywhere over Home or the tab strip
installDropOpenBridge()

contextBridge.exposeInMainWorld('editorApproval', {
  getMode: () => ipcRenderer.invoke('agent:editor-approval-mode'),
  setMode: (mode: 'auto' | 'per-action') => {
    if (mode !== 'auto' && mode !== 'per-action') throw new Error('Invalid Agent approval mode.')
    return ipcRenderer.invoke('agent:set-editor-approval-mode', mode)
  },
  onModeChanged: (handler: (mode: 'auto' | 'per-action') => void) => {
    const listener = (_event: unknown, mode: 'auto' | 'per-action') => handler(mode)
    ipcRenderer.on('agent:editor-approval-mode-changed', listener)
    return () => ipcRenderer.removeListener('agent:editor-approval-mode-changed', listener)
  },
})

const managedSessions: import('@genoffice/project-store').ManagedSessionsApi = {
  projects: () => ipcRenderer.invoke('managed-session:call', 'projects'),
  createProject: (name) => ipcRenderer.invoke('managed-session:call', 'create-project', { name }),
  renameProject: (id, name) =>
    ipcRenderer.invoke('managed-session:call', 'rename-project', { id, name }),
  bindProjectFolder: (id) => ipcRenderer.invoke('managed-session:call', 'bind-project', { id }),
  list: (filter) => ipcRenderer.invoke('managed-session:call', 'list', filter),
  get: (id) => ipcRenderer.invoke('managed-session:call', 'get', { id }),
  create: (projectId, parentId) =>
    ipcRenderer.invoke('managed-session:call', 'create', { projectId, parentId }),
  rename: (id, title, options) =>
    ipcRenderer.invoke('managed-session:call', 'rename', { id, title, ...options }),
  move: (id, projectId, parentId, options) =>
    ipcRenderer.invoke('managed-session:call', 'move', { id, projectId, parentId, ...options }),
  archive: (id, archived) =>
    ipcRenderer.invoke('managed-session:call', 'archive', { id, archived }),
  trash: (id) => ipcRenderer.invoke('managed-session:call', 'trash', { id }),
  restore: (id) => ipcRenderer.invoke('managed-session:call', 'restore', { id }),
  destroy: (id) => ipcRenderer.invoke('managed-session:call', 'destroy', { id }),
  updateSummary: (id, markdown, options) =>
    ipcRenderer.invoke('managed-session:call', 'summary', { id, markdown, ...options }),
  setMemoryEnabled: (id, enabled, options) =>
    ipcRenderer.invoke('managed-session:call', 'memory-enabled', { id, enabled, ...options }),
  command: (command) => ipcRenderer.invoke('managed-session:call', 'command', command),
  resume: (id) => ipcRenderer.invoke('managed-session:call', 'resume', { id }),
  cancel: (id) => ipcRenderer.invoke('managed-session:call', 'cancel', { id }),
  exportSession: (id) => ipcRenderer.invoke('managed-session:call', 'export', { id }),
  importSession: () => ipcRenderer.invoke('managed-session:call', 'import'),
  onChanged: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('managed-session:changed', listener)
    return () => ipcRenderer.removeListener('managed-session:changed', listener)
  },
}
contextBridge.exposeInMainWorld('managedSessions', managedSessions)
contextBridge.exposeInMainWorld('managedHomeChat', {
  onOpen: (handler: (id: string) => void) => {
    const listener = (_event: Electron.IpcRendererEvent, id: string) => handler(id)
    ipcRenderer.on('home:open-conversation', listener)
    return () => ipcRenderer.removeListener('home:open-conversation', listener)
  },
  pending: () => ipcRenderer.invoke('home:pending-conversation'),
})
import { ASSISTANT_TASK_CHANNELS, type AssistantTasksApi } from '../shared/assistant-task-api'
