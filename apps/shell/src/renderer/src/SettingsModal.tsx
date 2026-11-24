import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import {
  AI_CUSTOM_FONT_MAX_PX,
  AI_CUSTOM_FONT_MIN_PX,
  DEFAULT_AI_PANEL_PREFS,
  Dropdown,
  aiPanelFontPx,
  clampAiCustomFontSize,
} from '@genoffice/ui'
import type { AiFontSize, AiPanelPrefs, AiPanelSide } from '@genoffice/ui'
import {
  DEFAULT_MAX_OUTPUT_TOKENS,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  clampMaxOutputTokens,
  AI_IMAGE_PROVIDERS,
  AI_SEARCH_PROVIDERS,
  resolveAiSearchSettings,
  imageConfigurationError,
  imageProviderMeta,
  resolveImageSettings,
} from '@genoffice/ai-provider/browser'
import type { AiImageProviderConfig, AiImageProviderId, AiSettings } from '@genoffice/ai-provider'
import { aiModelCopy } from './ai-model-copy'
import { settingsCopy } from './settings-copy'
import { ApiModelPicker } from './ApiModelPicker'
import { useI18n } from './locale'
import type { StringKey, TFunc } from './locale'
import type { AiCatalogEntry, UiTheme } from '../../shared/home-api'
import type { ProviderStatus, RuntimeProvidersSnapshot } from '../../shared/runtime-provider-api'
import { accountCopy, accountErrorText } from '../../shared/account-copy'
import { ProviderLogo } from './provider-logos'
import { bigDogCopy } from './bigdog-copy'
import { BigDogWebsiteLink } from './BigDogWebsiteLink'
import { onboardingCopy, openOnboarding } from './onboarding-copy'
import { EnvironmentPanel } from './settings/EnvironmentPanel'
import { RemotePanel } from './settings/RemotePanel'
import { environmentCopy } from '../../shared/environment-copy'
import type { BigDogReleaseStatus } from '../../shared/bigdog-release'
import type { UpdateChannel } from '../../shared/update-api'
import appIcon from './assets/app-icon.svg'
import { IntegrationsPane, skillUpdateDue } from './IntegrationsPane'
import { FoundSkills } from './FoundSkills'
import { DefaultAppSetting } from './DefaultAppSetting'
import './settings.css'

// ── Settings modal (opened from the account menu) ─────────
// Two-pane dialog: section nav on the left, fields on the right.
// All values go through the existing home IPC; nothing is stored locally.

// sorted by ISO 639 language code — native-script labels have no natural
// shared alphabet, so the code is the ordering key
const LANG_OPTIONS = [
  { value: 'ar', label: 'العربية' },
  { value: 'cs', label: 'Čeština' },
  { value: 'de', label: 'Deutsch' },
  { value: 'en', label: 'English' },
  { value: 'es', label: 'Español' },
  { value: 'fr', label: 'Français' },
  { value: 'he', label: 'עברית' },
  { value: 'hi', label: 'हिन्दी' },
  { value: 'id', label: 'Bahasa Indonesia' },
  { value: 'it', label: 'Italiano' },
  { value: 'ja', label: '日本語' },
  { value: 'ko', label: '한국어' },
  { value: 'ms', label: 'Bahasa Melayu' },
  { value: 'nl', label: 'Nederlands' },
  { value: 'pl', label: 'Polski' },
  { value: 'pt', label: 'Português' },
  { value: 'ru', label: 'Русский' },
  { value: 'th', label: 'ไทย' },
  { value: 'zh', label: '简体中文' },
  { value: 'zh-TW', label: '繁體中文' },
] as const

// GenMail's option order: follow-system first, then the manual picks
const THEME_OPTIONS = [
  { value: 'system', labelKey: 'themeSystem' },
  { value: 'light', labelKey: 'themeLight' },
  { value: 'dark', labelKey: 'themeDark' },
] as const satisfies readonly { value: UiTheme; labelKey: StringKey }[]

const AI_FONT_SIZE_OPTIONS = [
  { value: 'default', labelKey: 'aiFontSizeDefault' },
  { value: 'large', labelKey: 'aiFontSizeLarge' },
  { value: 'xlarge', labelKey: 'aiFontSizeXLarge' },
  { value: 'custom', labelKey: 'aiFontSizeCustom' },
] as const satisfies readonly { value: AiFontSize; labelKey: StringKey }[]

const CHANNEL_OPTIONS = [
  { value: 'stable', labelKey: 'channelStable' },
  { value: 'beta', labelKey: 'channelBeta' },
] as const satisfies readonly { value: 'stable' | 'beta'; labelKey: StringKey }[]

/** px stepper for the custom AI panel text size; in-range values apply live,
 * out-of-range or partial input is clamped on blur */
function CustomFontSizeInput({
  value,
  label,
  onCommit,
}: {
  value: number
  label: string
  onCommit: (px: number) => void
}) {
  const [draft, setDraft] = useState(String(value))
  const [editing, setEditing] = useState(false)
  const shown = editing ? draft : String(value)
  const commit = (raw: string) => {
    const px = clampAiCustomFontSize(raw)
    if (px !== null && px !== value) onCommit(px)
  }
  return (
    <label className="set-num">
      <input
        type="number"
        className="set-input set-num-input"
        aria-label={label}
        min={AI_CUSTOM_FONT_MIN_PX}
        max={AI_CUSTOM_FONT_MAX_PX}
        step={1}
        value={shown}
        onFocus={() => {
          setDraft(String(value))
          setEditing(true)
        }}
        onChange={(e) => {
          setDraft(e.target.value)
          const n = Number(e.target.value)
          if (Number.isInteger(n) && n >= AI_CUSTOM_FONT_MIN_PX && n <= AI_CUSTOM_FONT_MAX_PX) {
            onCommit(n)
          }
        }}
        onBlur={() => {
          commit(draft)
          setEditing(false)
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') e.currentTarget.blur()
        }}
      />
      <span className="set-num-unit">px</span>
    </label>
  )
}

type SectionId =
  'aiModel' | 'general' | 'environment' | 'remote' | 'integrations' | 'skills' | 'about'
export interface SettingsTarget {
  section: string
  anchor?: string
  block?: string
}

const SECTIONS: readonly { id: SectionId; labelKey: StringKey | 'environment' | 'remote' }[] = [
  { id: 'aiModel', labelKey: 'setSecAiModel' },
  { id: 'general', labelKey: 'setSecGeneral' },
  { id: 'environment', labelKey: 'environment' },
  { id: 'remote', labelKey: 'remote' },
  { id: 'integrations', labelKey: 'setSecIntegrations' },
  { id: 'skills', labelKey: 'setSecSkills' },
  { id: 'about', labelKey: 'setSecAbout' },
]

function SectionIcon({ id }: { id: SectionId }) {
  if (id === 'remote')
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <rect
          x="4.5"
          y="1.5"
          width="7"
          height="13"
          rx="1.6"
          stroke="currentColor"
          strokeWidth="1.3"
        />
        <path d="M7 12h2" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      </svg>
    )
  if (id === 'aiModel') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M8 1.8 9.5 6l4.2 1.5L9.5 9 8 13.2 6.5 9 2.3 7.5 6.5 6 8 1.8Z"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinejoin="round"
        />
        <path
          d="M12.8 11.2v3M11.3 12.7h3"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
      </svg>
    )
  }
  if (id === 'integrations') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M5.5 2v3M10.5 2v3M4 5h8v2.5a4 4 0 0 1-8 0V5ZM8 11.5V14"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    )
  }
  if (id === 'general') {
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <path
          d="M2 5h8M13 5h1M2 11h1M6 11h8"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
        <circle cx="11.5" cy="5" r="1.7" stroke="currentColor" strokeWidth="1.3" />
        <circle cx="4.5" cy="11" r="1.7" stroke="currentColor" strokeWidth="1.3" />
      </svg>
    )
  }
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <circle cx="8" cy="8" r="6.3" stroke="currentColor" strokeWidth="1.3" />
      <path d="M8 7.4v3.4" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      <circle cx="8" cy="5.1" r="0.8" fill="currentColor" />
    </svg>
  )
}

/** label-over-value field row with an optional right-aligned action */
function Field({
  label,
  value,
  valueTitle,
  action,
}: {
  label: string
  value: ReactNode
  valueTitle?: string
  action?: ReactNode
}) {
  return (
    <div className="set-field">
      <div className="set-field-text">
        <div className="set-field-label">{label}</div>
        <div className="set-field-value" data-tip={valueTitle}>
          {value}
        </div>
      </div>
      {action}
    </div>
  )
}

function AccountConnections({ lang, page }: { lang: string; page: 'api' | 'accounts' }) {
  const copy = accountCopy(lang)
  const [snapshot, setSnapshot] = useState<RuntimeProvidersSnapshot | null>(null)
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let alive = true
    const getRuntimeProviders = window.aiOffice.getRuntimeProviders
    if (!getRuntimeProviders) {
      setSnapshot({ selection: 'api', providers: [] })
      return () => {
        alive = false
      }
    }
    void getRuntimeProviders()
      .then((value) => {
        if (alive) setSnapshot(value)
      })
      .catch((reason) => {
        if (alive) setError(accountErrorText(reason, lang))
      })
    return () => {
      alive = false
    }
  }, [lang])

  const run = (key: string, action: () => Promise<RuntimeProvidersSnapshot>) => {
    setBusy(key)
    setError(null)
    void action()
      .then(setSnapshot)
      .catch((reason) => setError(accountErrorText(reason, lang)))
      .finally(() => setBusy(null))
  }

  const runtimeCard = (provider: ProviderStatus) => {
    const selected = snapshot?.selection === provider.id
    const detail = [copy.phase[provider.phase], provider.plan, provider.accountLabel]
      .filter(Boolean)
      .join(' · ')
    const working = busy?.endsWith(`:${provider.id}`) === true
    return (
      <div className={`set-runtime-card${selected ? ' selected' : ''}`} key={provider.id}>
        <div className={`set-runtime-mark ${provider.id}`} aria-hidden="true">
          {provider.id === 'codex' ? 'C' : 'A'}
        </div>
        <div className="set-runtime-copy">
          <div className="set-runtime-name">{provider.label}</div>
          <div className="set-runtime-meta">
            {provider.phase === 'blocked'
              ? copy.claudeUnavailable
              : provider.error
                ? accountErrorText(provider.error, lang)
                : detail}
          </div>
        </div>
        <div className="set-runtime-actions">
          {provider.installed && (
            <button
              className="set-runtime-link"
              disabled={busy !== null}
              onClick={() =>
                run(`update:${provider.id}`, () =>
                  window.aiOffice.installRuntimeProvider(provider.id),
                )
              }
            >
              {busy === `update:${provider.id}` ? copy.busy : copy.update}
            </button>
          )}
          <button
            className={selected ? 'set-btn' : 'set-btn primary'}
            disabled={busy !== null || (selected && provider.ready) || provider.phase === 'blocked'}
            onClick={() =>
              provider.ready
                ? run(`select:${provider.id}`, () =>
                    window.aiOffice.selectAgentRuntime(provider.id),
                  )
                : run(`connect:${provider.id}`, () =>
                    window.aiOffice.connectRuntimeProvider(provider.id),
                  )
            }
          >
            {working
              ? copy.busy
              : selected && provider.ready
                ? copy.selected
                : provider.phase === 'blocked'
                  ? copy.phase.blocked
                  : provider.ready
                    ? copy.use
                    : copy.connect}
          </button>
          {provider.authenticated && (
            <button
              className="set-runtime-link"
              disabled={busy !== null}
              onClick={() =>
                run(`logout:${provider.id}`, () =>
                  window.aiOffice.logoutRuntimeProvider(provider.id),
                )
              }
            >
              {copy.logout}
            </button>
          )}
        </div>
      </div>
    )
  }

  return (
    <section className="set-runtime-section">
      {page === 'accounts' && (
        <div>
          <h4 id="set-runtime-heading" className="set-runtime-heading">
            {copy.title}
          </h4>
          <p className="set-runtime-description">{copy.description}</p>
        </div>
      )}
      {page === 'api' && (
        <div className={`set-runtime-card${snapshot?.selection === 'api' ? ' selected' : ''}`}>
          <div className="set-runtime-mark api" aria-hidden="true">
            API
          </div>
          <div className="set-runtime-copy">
            <div className="set-runtime-name">{copy.api}</div>
            <div className="set-runtime-meta">{copy.apiDetail}</div>
          </div>
          <button
            className={snapshot?.selection === 'api' ? 'set-btn' : 'set-btn primary'}
            disabled={!snapshot || snapshot.selection === 'api' || busy !== null}
            onClick={() => run('select:api', () => window.aiOffice.selectAgentRuntime('api'))}
          >
            {snapshot?.selection === 'api' ? copy.selected : copy.use}
          </button>
        </div>
      )}
      {page === 'accounts' && snapshot?.providers.map(runtimeCard)}
      {!snapshot && !error && <div className="set-runtime-loading">{copy.busy}</div>}
      {error && (
        <div className="set-runtime-error" role="status">
          {error}
        </div>
      )}
      {page === 'accounts' && <p className="set-runtime-boundary">{copy.boundary}</p>}
    </section>
  )
}

/** Chat API, official account login and image generation are independent pages. */
function AiModelPane({ t, lang }: { t: TFunc; lang: string }) {
  const copy = aiModelCopy(lang)
  const [page, setPage] = useState<'api' | 'accounts' | 'images' | 'search'>('api')
  const [catalog] = useState<AiCatalogEntry[]>(() => window.aiOffice.getAiProviders?.() ?? [])
  const [settings, setSettings] = useState<AiSettings | null>(null)
  const [dirty, setDirty] = useState(false)
  const [saved, setSaved] = useState(false)
  const [saving, setSaving] = useState(false)
  const [testing, setTesting] = useState(false)
  const [feedback, setFeedback] = useState<AiStatus | null>(null)
  const [loadError, setLoadError] = useState<string | null>(null)
  const [maxTokensDraft, setMaxTokensDraft] = useState<string | null>(null)
  const [modelDiscoveryRevision, setModelDiscoveryRevision] = useState(0)
  const revision = useRef(0)
  const feedbackRevision = useRef(0)
  const alive = useRef(true)

  useEffect(() => {
    alive.current = true
    let current = true
    void window.aiOffice
      .getAiSettings?.()
      .then((s) => {
        if (current && s) {
          setSettings({ ...s, imageGeneration: resolveImageSettings(s.imageGeneration, s) })
        }
      })
      .catch((error) => {
        if (current) setLoadError(error instanceof Error ? error.message : String(error))
      })
    return () => {
      alive.current = false
      current = false
    }
  }, [])

  const touch = () => {
    revision.current++
    feedbackRevision.current++
    setDirty(true)
    setSaved(false)
    setFeedback(null)
  }
  const updateConfig = (patch: Partial<AiSettings['providers']['openai']>, userEdit = true) => {
    if (
      !settings ||
      Object.entries(patch).every(
        ([key, value]) =>
          JSON.stringify(settings.providers[settings.provider][key as keyof typeof patch]) ===
          JSON.stringify(value),
      )
    )
      return
    setSettings((previous) =>
      previous
        ? {
            ...previous,
            providers: {
              ...previous.providers,
              [previous.provider]: {
                ...previous.providers[previous.provider],
                ...patch,
                ...((patch.apiKey !== undefined &&
                  patch.apiKey !== previous.providers[previous.provider].apiKey) ||
                (patch.baseUrl !== undefined &&
                  patch.baseUrl !== previous.providers[previous.provider].baseUrl)
                  ? { modelCatalog: undefined }
                  : {}),
              },
            },
          }
        : previous,
    )
    if (userEdit) touch()
  }
  const updateImageConfig = (patch: Partial<AiImageProviderConfig>) => {
    if (
      !settings?.imageGeneration ||
      Object.entries(patch).every(
        ([key, value]) =>
          settings.imageGeneration!.providers[settings.imageGeneration!.provider][
            key as keyof AiImageProviderConfig
          ] === value,
      )
    )
      return
    setSettings((previous) => {
      if (!previous) return previous
      const images = resolveImageSettings(previous.imageGeneration, previous)
      return {
        ...previous,
        imageGeneration: {
          ...images,
          providers: {
            ...images.providers,
            [images.provider]: { ...images.providers[images.provider], ...patch },
          },
        },
      }
    })
    touch()
  }
  const commitMaxTokens = () => {
    if (maxTokensDraft === null || !settings) return
    setMaxTokensDraft(null)
    const next = clampMaxOutputTokens(Number.parseInt(maxTokensDraft, 10))
    if (next === (settings.maxOutputTokens ?? DEFAULT_MAX_OUTPUT_TOKENS)) return
    setSettings((previous) => (previous ? { ...previous, maxOutputTokens: next } : previous))
    touch()
  }
  const save = async () => {
    if (!settings || saving) return
    const next =
      maxTokensDraft === null
        ? settings
        : {
            ...settings,
            maxOutputTokens: clampMaxOutputTokens(Number.parseInt(maxTokensDraft, 10)),
          }
    const version = revision.current
    setSaving(true)
    setFeedback(null)
    try {
      await window.aiOffice.setAiSettings(next)
      if (alive.current && revision.current === version) {
        setSettings(next)
        setMaxTokensDraft(null)
        setDirty(false)
        setSaved(true)
      }
    } catch (error) {
      if (alive.current && revision.current === version)
        setFeedback({ kind: 'err', text: error instanceof Error ? error.message : copy.saveFailed })
    } finally {
      if (alive.current) setSaving(false)
    }
  }
  const test = async () => {
    if (!settings || testing) return
    const version = feedbackRevision.current
    setTesting(true)
    setFeedback(null)
    try {
      const result =
        page === 'search'
          ? await window.aiOffice.testAiSearchSettings({
              provider: search.provider,
              apiKey:
                search.provider === 'genspark' ? '' : search.providers[search.provider].apiKey,
            })
          : await window.aiOffice.testAiSettings(settings)
      if (alive.current && feedbackRevision.current === version)
        setFeedback(
          result.ok
            ? { kind: 'ok', text: t('setAiTestOk') }
            : { kind: 'err', text: result.error || t('setAiTestFail') },
        )
    } catch (error) {
      if (alive.current && feedbackRevision.current === version)
        setFeedback({ kind: 'err', text: error instanceof Error ? error.message : String(error) })
    } finally {
      if (alive.current) setTesting(false)
    }
  }
  const meta = catalog.find((entry) => entry.id === settings?.provider)
  const config = settings?.providers[settings.provider]
  const images = settings?.imageGeneration
  const imageMeta = images && imageProviderMeta(images.provider)
  const imageConfig = images?.providers[images.provider]
  const pages = ['api', 'accounts', 'images', 'search'] as const
  const search = resolveAiSearchSettings(settings?.search)
  const switchPage = (next: typeof page) => {
    setPage(next)
    setFeedback(null)
    // Prevent late chat-test responses from appearing on a different page.
    feedbackRevision.current++
  }

  return (
    <>
      <h3 className="set-pane-title">{t('setSecAiModel')}</h3>
      <div className="set-ai-tabs" role="tablist" aria-label={copy.tabsLabel}>
        {pages.map((id, index) => (
          <button
            key={id}
            id={`set-ai-tab-${id}`}
            type="button"
            role="tab"
            aria-selected={page === id}
            aria-controls={`set-ai-panel-${id}`}
            tabIndex={page === id ? 0 : -1}
            onClick={() => switchPage(id)}
            onKeyDown={(event) => {
              const next =
                event.key === 'ArrowRight'
                  ? (index + 1) % pages.length
                  : event.key === 'ArrowLeft'
                    ? (index + pages.length - 1) % pages.length
                    : event.key === 'Home'
                      ? 0
                      : event.key === 'End'
                        ? pages.length - 1
                        : -1
              if (next < 0) return
              event.preventDefault()
              const nextPage = pages[next]!
              switchPage(nextPage)
              document.getElementById(`set-ai-tab-${nextPage}`)?.focus()
            }}
          >
            {copy.tabs[id]}
          </button>
        ))}
      </div>
      {dirty && (
        <p className="set-ai-draft" role="status">
          {copy.unsaved}
        </p>
      )}
      <div
        className="set-ai-page"
        role="tabpanel"
        id={`set-ai-panel-${page}`}
        aria-labelledby={`set-ai-tab-${page}`}
      >
        <div hidden={page === 'images' || page === 'search'}>
          <AccountConnections lang={lang} page={page === 'accounts' ? 'accounts' : 'api'} />
        </div>
        {page === 'search' && settings && (
          <>
            <p className="set-field-desc set-ai-note">{copy.searchHint}</p>
            <div className="set-field">
              <label className="set-field-label" htmlFor="set-ai-search-provider">
                {copy.searchProvider}
              </label>
              <select
                id="set-ai-search-provider"
                className="set-input"
                value={search.provider}
                onChange={(event) => {
                  if (event.target.value === search.provider) return
                  setSettings({
                    ...settings,
                    search: { ...search, provider: event.target.value as typeof search.provider },
                  })
                  touch()
                }}
              >
                {AI_SEARCH_PROVIDERS.map((provider) => (
                  <option key={provider.id} value={provider.id}>
                    {provider.id === 'genspark' ? copy.searchDefault : provider.label}
                  </option>
                ))}
              </select>
            </div>
            {search.provider !== 'genspark' && (
              <>
                <p className="set-field-desc">
                  {search.provider === 'parallel' ? copy.parallelKeyHint : copy.searchKeyHint}
                </p>
                <div className="set-field">
                  <label className="set-field-label" htmlFor="set-ai-search-key">
                    API Key
                  </label>
                  <input
                    id="set-ai-search-key"
                    className="set-input"
                    type="password"
                    autoComplete="off"
                    value={search.providers[search.provider].apiKey}
                    onChange={(event) => {
                      setSettings({
                        ...settings,
                        search: {
                          ...search,
                          providers: {
                            ...search.providers,
                            [search.provider]: { apiKey: event.target.value.trim() },
                          },
                        },
                      })
                      touch()
                    }}
                  />
                </div>
              </>
            )}
          </>
        )}
        {page !== 'accounts' && loadError && (
          <p className="set-runtime-error" role="alert">
            {copy.loadFailed}: {loadError}
          </p>
        )}
        {page !== 'accounts' && !settings && !loadError && <p>{copy.loading}</p>}
        {page === 'api' && settings && config && (
          <>
            <p className="set-field-desc set-ai-note">{copy.chatHint}</p>
            <div className="set-field">
              <div className="set-field-text">
                <label className="set-field-label">{t('setAiProvider')}</label>
              </div>
              <Dropdown
                className="set-dd"
                value={settings.provider}
                ariaLabel={t('setAiProvider')}
                options={catalog.map((entry) => ({
                  value: entry.id,
                  label: entry.label,
                  render: (
                    <>
                      <ProviderLogo id={entry.id} />
                      {entry.label}
                    </>
                  ),
                }))}
                onPick={(value) => {
                  if (value === settings.provider) return
                  setSettings((previous) =>
                    previous
                      ? {
                          ...previous,
                          provider: value as AiSettings['provider'],
                          gskToolsEnabled: false,
                        }
                      : previous,
                  )
                  touch()
                }}
              />
            </div>
            {meta && (
              <ApiModelPicker
                key={settings.provider}
                provider={settings.provider}
                config={config}
                meta={meta}
                label={t('setAiModelId')}
                lang={lang}
                settledRevision={modelDiscoveryRevision}
                onChange={updateConfig}
              />
            )}
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <label className="set-field-label" htmlFor="set-ai-max-tokens">
                    {t('setAiMaxTokens')}
                  </label>
                  <div className="set-field-desc">{t('setAiMaxTokensDesc')}</div>
                </div>
              </div>
              <input
                id="set-ai-max-tokens"
                className="set-input"
                type="number"
                min={MIN_MAX_OUTPUT_TOKENS}
                max={MAX_MAX_OUTPUT_TOKENS}
                step={1024}
                value={
                  maxTokensDraft ?? String(settings.maxOutputTokens ?? DEFAULT_MAX_OUTPUT_TOKENS)
                }
                onChange={(event) => {
                  setMaxTokensDraft(event.target.value)
                  touch()
                }}
                onBlur={commitMaxTokens}
              />
            </div>
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <label className="set-field-label" htmlFor="set-ai-key">
                    {t('setAiApiKey')}
                  </label>
                  <div className="set-field-desc">{t('setAiKeyHint')}</div>
                </div>
              </div>
              <input
                id="set-ai-key"
                className="set-input"
                type="password"
                value={config.apiKey}
                placeholder={meta?.keyPlaceholder ?? 'API Key'}
                spellCheck={false}
                autoComplete="off"
                onChange={(event) => updateConfig({ apiKey: event.target.value.trim() })}
                onBlur={() => setModelDiscoveryRevision((value) => value + 1)}
              />
            </div>
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <label className="set-field-label" htmlFor="set-ai-base-url">
                    {t('setAiBaseUrl')}
                  </label>
                  {!meta?.needsBaseUrl && (
                    <div className="set-field-desc">{t('setAiBaseUrlHint')}</div>
                  )}
                </div>
              </div>
              <input
                id="set-ai-base-url"
                className="set-input"
                type="text"
                value={config.baseUrl ?? ''}
                placeholder={meta?.needsBaseUrl ? 'https://…/v1' : meta?.defaultBaseUrl}
                spellCheck={false}
                onChange={(event) => updateConfig({ baseUrl: event.target.value.trim() })}
                onBlur={() => setModelDiscoveryRevision((value) => value + 1)}
              />
            </div>
          </>
        )}
        {page === 'images' && images && imageMeta && imageConfig && (
          <>
            <p className="set-field-desc set-ai-note">{copy.imageHint}</p>
            <div className="set-field">
              <div className="set-field-text">
                <label className="set-field-label">{copy.imageProvider}</label>
              </div>
              <Dropdown
                className="set-dd"
                value={images.provider}
                ariaLabel={copy.imageProvider}
                options={AI_IMAGE_PROVIDERS.map((entry) => ({
                  value: entry.id,
                  label: entry.label,
                  render: (
                    <>
                      <ProviderLogo id={entry.id} />
                      {entry.label}
                    </>
                  ),
                }))}
                onPick={(value) => {
                  if (value === images.provider) return
                  setSettings((previous) =>
                    previous
                      ? {
                          ...previous,
                          imageGeneration: {
                            ...resolveImageSettings(previous.imageGeneration, previous),
                            provider: value as AiImageProviderId,
                          },
                        }
                      : previous,
                  )
                  touch()
                }}
              />
            </div>
            {(images.provider === 'custom' || images.provider === 'qwen') && (
              <div className="set-ai-protocol">
                <span>{copy.protocol}</span>
                <code>{imageMeta.protocol}</code>
                <p>
                  {images.provider === 'custom'
                    ? copy.customHint
                    : images.provider === 'qwen'
                      ? copy.qwenHint
                      : copy.imageBaseHint}
                </p>
              </div>
            )}
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <label className="set-field-label" htmlFor="set-ai-image-model">
                    {copy.imageModel}
                  </label>
                  <div className="set-field-desc">{copy.imageModelHint}</div>
                </div>
              </div>
              <input
                id="set-ai-image-model"
                className="set-input"
                type="text"
                list="set-ai-image-model-options"
                value={imageConfig.model}
                placeholder={imageMeta.defaultModel || 'image-model-id'}
                spellCheck={false}
                onChange={(event) => updateImageConfig({ model: event.target.value.trim() })}
              />
              <datalist id="set-ai-image-model-options">
                {imageMeta.models.map((model) => (
                  <option key={model} value={model} />
                ))}
              </datalist>
            </div>
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <label className="set-field-label" htmlFor="set-ai-image-key">
                    {copy.imageKey}
                  </label>
                  <div className="set-field-desc">{copy.imageKeyHint}</div>
                </div>
              </div>
              <input
                id="set-ai-image-key"
                className="set-input"
                type="password"
                value={imageConfig.apiKey}
                placeholder={imageMeta.keyPlaceholder}
                spellCheck={false}
                autoComplete="off"
                onChange={(event) => updateImageConfig({ apiKey: event.target.value.trim() })}
              />
            </div>
            <div className="set-field">
              <div className="set-field-text">
                <div className="set-field-stack">
                  <label className="set-field-label" htmlFor="set-ai-image-base-url">
                    {copy.imageBase}
                  </label>
                  <div className="set-field-desc">{copy.imageBaseHint}</div>
                </div>
              </div>
              <input
                id="set-ai-image-base-url"
                className="set-input"
                type="text"
                value={imageConfig.baseUrl ?? ''}
                placeholder={imageMeta.defaultBaseUrl || 'https://…/v1'}
                spellCheck={false}
                onChange={(event) => updateImageConfig({ baseUrl: event.target.value.trim() })}
              />
            </div>
            <p className="set-field-desc set-ai-note">{copy.imageCheckHint}</p>
          </>
        )}
      </div>
      {page !== 'accounts' && settings && (
        <div className="set-pane-footer">
          <AiStatusPill
            status={
              feedback ??
              (saving
                ? { kind: 'testing', text: copy.saving }
                : testing && (page === 'api' || page === 'search')
                  ? { kind: 'testing', text: t('setAiTesting') }
                  : saved
                    ? { kind: 'ok', text: t('setAiSaved') }
                    : null)
            }
          />
          {page === 'api' || page === 'search' ? (
            <button
              className="set-btn"
              disabled={testing || saving}
              aria-busy={testing}
              onClick={() => void test()}
            >
              {testing ? t('setAiTesting') : t('setAiTest')}
            </button>
          ) : (
            <button
              className="set-btn"
              onClick={() => {
                if (!images || !imageConfig) return
                const error = imageConfigurationError(images.provider, imageConfig, lang)
                setFeedback(
                  error ? { kind: 'err', text: error } : { kind: 'ok', text: copy.imageChecked },
                )
              }}
            >
              {copy.checkImage}
            </button>
          )}
          <button
            className="set-btn primary"
            disabled={!dirty || saving}
            aria-busy={saving}
            onClick={() => void save()}
          >
            {saving ? copy.saving : t('setAiSave')}
          </button>
        </div>
      )}
    </>
  )
}

interface AiStatus {
  kind: 'testing' | 'ok' | 'err'
  text: string
}

/** colored feedback pill in the AI pane footer: spinner while testing, then success/error */
function AiStatusPill({ status }: { status: AiStatus | null }) {
  if (!status) return null
  return (
    <span
      className={`set-ai-status ${status.kind}`}
      role="status"
      // error text (HTTP body, network message) can be long — full text via native tooltip
      title={status.kind === 'err' ? status.text : undefined}
    >
      {status.kind === 'testing' ? (
        <span className="set-ai-spin" aria-hidden="true" />
      ) : status.kind === 'ok' ? (
        <svg
          className="set-ai-status-icon"
          width="14"
          height="14"
          viewBox="0 0 14 14"
          aria-hidden="true"
        >
          <circle cx="7" cy="7" r="6.3" fill="currentColor" opacity="0.16" />
          <path
            d="M4.2 7.3l1.9 1.9 3.7-4.3"
            stroke="currentColor"
            strokeWidth="1.5"
            strokeLinecap="round"
            strokeLinejoin="round"
            fill="none"
          />
        </svg>
      ) : (
        <svg
          className="set-ai-status-icon"
          width="14"
          height="14"
          viewBox="0 0 14 14"
          aria-hidden="true"
        >
          <circle cx="7" cy="7" r="6.3" fill="currentColor" opacity="0.16" />
          <path d="M7 3.8v3.9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
          <circle cx="7" cy="10.1" r="1" fill="currentColor" />
        </svg>
      )}
      <span className="set-ai-status-text">{status.text}</span>
    </span>
  )
}

export interface SettingsModalProps {
  onClose: () => void
  initialSection?: string | undefined
  /** an installed skill is older than the bundled one: dot on the Integrations entry */
  skillUpdateDue?: boolean
  onSkillUpdateDue?: (due: boolean) => void
}

export function SettingsModal({
  onClose,
  initialSection,
  skillUpdateDue: updateDue = false,
  onSkillUpdateDue,
}: SettingsModalProps) {
  const { lang, setLang, t } = useI18n()
  const uiCopy = settingsCopy(lang)
  const [section, setSection] = useState<SectionId>(
    SECTIONS.some((entry) => entry.id === initialSection)
      ? (initialSection as SectionId)
      : 'aiModel',
  )
  useEffect(() => {
    if (SECTIONS.some((entry) => entry.id === initialSection))
      setSection(initialSection as SectionId)
  }, [initialSection])
  const [theme, setTheme] = useState<UiTheme>('system')
  const [saveDir, setSaveDir] = useState('')
  const [analyticsOn, setAnalyticsOn] = useState(true)
  const [analyticsSaving, setAnalyticsSaving] = useState(false)
  const [autoSaveOn, setAutoSaveOn] = useState(false)
  const [aiPrefs, setAiPrefs] = useState<AiPanelPrefs>(DEFAULT_AI_PANEL_PREFS)
  const [channel, setChannel] = useState<UpdateChannel>('stable')
  const [appVersion, setAppVersion] = useState('')
  const brand = bigDogCopy(lang)
  const [releaseStatus, setReleaseStatus] = useState<BigDogReleaseStatus | null>(null)
  const [releaseBusy, setReleaseBusy] = useState(false)
  const [releaseError, setReleaseError] = useState('')
  const [generalFeedback, setGeneralFeedback] = useState<AiStatus | null>(null)
  const [generalBusy, setGeneralBusy] = useState(false)
  const generalRevision = useRef(0)
  const generalQueue = useRef(Promise.resolve())
  const mounted = useRef(true)

  useEffect(() => {
    let alive = true
    mounted.current = true
    void window.aiOffice.getTheme?.().then((th) => {
      if (alive) setTheme(th)
    })
    void window.aiOffice.getDefaultSaveDir?.().then((dir) => {
      if (alive && dir) setSaveDir(dir)
    })
    void window.aiOffice.getAnalyticsEnabled?.().then((on) => {
      if (alive) setAnalyticsOn(on !== false)
    })
    void window.aiOffice.getAutoSaveDefault?.().then((v) => {
      if (alive) setAutoSaveOn(v.on)
    })
    void window.aiOffice.getAiPanelPrefs?.().then((prefs) => {
      if (alive) setAiPrefs(prefs)
    })
    void window.aiOffice.getUpdateChannel?.().then((ch) => {
      if (alive) setChannel(ch)
    })
    void window.aiOffice.getAppVersion?.().then((v) => {
      if (alive && v) setAppVersion(v)
    })
    return () => {
      alive = false
      mounted.current = false
    }
  }, [])

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [onClose])

  const releaseRequest = useRef(false)
  const checkRelease = async (next?: UpdateChannel) => {
    if (releaseRequest.current) return
    releaseRequest.current = true
    setReleaseBusy(true)
    setReleaseError('')
    setReleaseStatus(null)
    try {
      if (next) {
        await window.aiOffice.setUpdateChannel(next)
        if (mounted.current) setChannel(next)
      }
      const status = await window.aiOffice.checkBigDogRelease()
      if (mounted.current) setReleaseStatus(status)
    } catch {
      if (mounted.current) setReleaseStatus({ phase: 'error' })
    } finally {
      releaseRequest.current = false
      if (mounted.current) setReleaseBusy(false)
    }
  }

  const generalAction = async (action: () => Promise<void | false>) => {
    const version = ++generalRevision.current
    setGeneralBusy(true)
    setGeneralFeedback({ kind: 'testing', text: uiCopy.saving })
    // Keep rapid custom-font edits in order; a late response cannot overwrite a newer patch.
    const operation = generalQueue.current.then(action)
    generalQueue.current = operation.then(
      () => {},
      () => {},
    )
    try {
      const result = await operation
      if (mounted.current && generalRevision.current === version)
        setGeneralFeedback(result === false ? null : { kind: 'ok', text: uiCopy.saved })
    } catch {
      if (mounted.current && generalRevision.current === version)
        setGeneralFeedback({ kind: 'err', text: uiCopy.saveFailed })
    } finally {
      if (mounted.current && generalRevision.current === version) setGeneralBusy(false)
    }
  }
  const applyTheme = (next: UiTheme) =>
    void generalAction(async () => {
      await window.aiOffice.setTheme(next)
      if (!mounted.current) return
      setTheme(next)
      if (next === 'system') document.documentElement.removeAttribute('data-theme')
      else document.documentElement.setAttribute('data-theme', next)
    })
  const updateAiPrefs = (patch: Partial<AiPanelPrefs>) => {
    void generalAction(async () => {
      const next = await window.aiOffice.setAiPanelPrefs(patch)
      if (mounted.current) setAiPrefs(next)
    })
  }

  const changeSaveDir = () => {
    void generalAction(async () => {
      const dir = await window.aiOffice.pickDefaultSaveDir()
      if (!dir) return false
      if (mounted.current) setSaveDir(dir)
    })
  }

  return (
    <div
      className="set-overlay"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div className="set-dialog" role="dialog" aria-modal="true" aria-label={t('settings')}>
        <div className="set-header">
          <h2 className="set-title">{t('settings')}</h2>
          <button className="set-close" onClick={onClose} aria-label={t('cancel')}>
            <svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true">
              <path
                d="M2 2l10 10M12 2L2 12"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
              />
            </svg>
          </button>
        </div>
        <div className="set-body">
          <nav className="set-nav" aria-label={t('settings')}>
            {SECTIONS.map((s) => (
              <button
                key={s.id}
                className={`set-nav-item${section === s.id ? ' active' : ''}`}
                aria-current={section === s.id}
                onClick={() => setSection(s.id)}
              >
                <SectionIcon id={s.id} />
                {s.labelKey === 'remote'
                  ? lang.startsWith('zh')
                    ? '手机协作'
                    : 'Phone collaboration'
                  : s.labelKey === 'environment'
                    ? environmentCopy(lang).section
                    : t(s.labelKey)}
                {s.id === 'integrations' && updateDue && (
                  <span className="set-nav-dot" role="img" aria-label={t('intgUpdateDue')} />
                )}
              </button>
            ))}
          </nav>
          <div className={`set-pane set-pane-${section}`}>
            {section === 'aiModel' && <AiModelPane t={t} lang={lang} />}
            {section === 'skills' && <FoundSkills t={t} />}
            {section === 'remote' && <RemotePanel />}
            {section === 'environment' && (
              <EnvironmentPanel
                onModels={() => setSection('aiModel')}
                onIntegrations={() => setSection('integrations')}
                onUpdates={() => setSection('about')}
                onClose={onClose}
              />
            )}
            {section === 'general' && (
              <>
                <h3 className="set-pane-title">{t('setSecGeneral')}</h3>
                <Field
                  label={onboardingCopy(lang).reopen}
                  value={onboardingCopy(lang).reopenHint}
                  action={
                    <button
                      className="set-btn"
                      onClick={() => {
                        onClose()
                        openOnboarding()
                      }}
                    >
                      {uiCopy.reopen}
                    </button>
                  }
                />
                <DefaultAppSetting t={t} />
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('language')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={lang}
                    ariaLabel={t('language')}
                    options={LANG_OPTIONS.map((opt) => ({ value: opt.value, label: opt.label }))}
                    onPick={(v) => setLang(v as typeof lang)}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('theme')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={theme}
                    disabled={generalBusy}
                    ariaLabel={t('theme')}
                    options={THEME_OPTIONS.map((opt) => ({
                      value: opt.value,
                      label: t(opt.labelKey),
                    }))}
                    onPick={(v) => applyTheme(v as UiTheme)}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('setAiPanelSide')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={aiPrefs.side}
                    disabled={generalBusy}
                    ariaLabel={t('setAiPanelSide')}
                    options={[
                      { value: 'left', label: t('aiPanelSideLeft') },
                      { value: 'right', label: t('aiPanelSideRight') },
                    ]}
                    onPick={(side) => updateAiPrefs({ side: side as AiPanelSide })}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('setAiFontSize')}</label>
                  </div>
                  {aiPrefs.fontSize === 'custom' && (
                    <CustomFontSizeInput
                      value={aiPrefs.customFontSize}
                      label={t('aiFontSizeCustom')}
                      onCommit={(px) => updateAiPrefs({ customFontSize: px })}
                    />
                  )}
                  <Dropdown
                    className="set-dd"
                    value={aiPrefs.fontSize}
                    disabled={generalBusy}
                    ariaLabel={t('setAiFontSize')}
                    options={AI_FONT_SIZE_OPTIONS.map((opt) => ({
                      value: opt.value,
                      label: t(opt.labelKey),
                    }))}
                    onPick={(v) => {
                      const fontSize = v as AiFontSize
                      // start the custom size from the preset being left so nothing jumps
                      updateAiPrefs(
                        fontSize === 'custom' && aiPrefs.fontSize !== 'custom'
                          ? { fontSize, customFontSize: aiPanelFontPx(aiPrefs) }
                          : { fontSize },
                      )
                    }}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-stack">
                      <div className="set-field-label">{t('setAiSpellcheck')}</div>
                      <div className="set-field-desc">{t('setAiSpellcheckDesc')}</div>
                    </div>
                  </div>
                  <button
                    className="set-switch"
                    role="switch"
                    aria-checked={aiPrefs.spellcheck}
                    disabled={generalBusy}
                    aria-label={t('setAiSpellcheck')}
                    onClick={() => updateAiPrefs({ spellcheck: !aiPrefs.spellcheck })}
                  />
                </div>
                <Field
                  label={t('saveLocation')}
                  value={saveDir || '—'}
                  valueTitle={saveDir}
                  action={
                    <button className="set-btn" disabled={generalBusy} onClick={changeSaveDir}>
                      {t('setChange')}
                    </button>
                  }
                />
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-stack">
                      <div className="set-field-label">{t('setAutoSave')}</div>
                      <div className="set-field-desc">{t('setAutoSaveDesc')}</div>
                    </div>
                  </div>
                  <button
                    className="set-switch"
                    role="switch"
                    aria-checked={autoSaveOn}
                    disabled={generalBusy}
                    aria-label={t('setAutoSave')}
                    onClick={() => {
                      const next = !autoSaveOn
                      void generalAction(async () => {
                        await window.aiOffice.setAutoSaveDefault(next)
                        if (mounted.current) setAutoSaveOn(next)
                      })
                    }}
                  />
                </div>
                <div className="set-field">
                  <div className="set-field-text">
                    <div className="set-field-stack">
                      <div className="set-field-label">{t('setAnalytics')}</div>
                      <div className="set-field-desc">{t('setAnalyticsDesc')}</div>
                    </div>
                  </div>
                  <button
                    className="set-switch"
                    role="switch"
                    aria-checked={analyticsOn}
                    aria-label={t('setAnalytics')}
                    disabled={analyticsSaving}
                    onClick={() => {
                      const next = !analyticsOn
                      setAnalyticsSaving(true)
                      void window.aiOffice
                        .setAnalyticsEnabled(next)
                        .then((persisted) => {
                          if (persisted) setAnalyticsOn(next)
                        })
                        .catch(() => setGeneralFeedback({ kind: 'err', text: uiCopy.saveFailed }))
                        .finally(() => setAnalyticsSaving(false))
                    }}
                  />
                </div>
                <AiStatusPill status={generalFeedback} />
              </>
            )}
            {section === 'integrations' && (
              <IntegrationsPane t={t} onStatus={(st) => onSkillUpdateDue?.(skillUpdateDue(st))} />
            )}
            {section === 'about' && (
              <>
                <h3 className="set-pane-title">{t('setSecAbout')}</h3>
                <div className="set-about-brand">
                  <img src={appIcon} alt="" />
                  <div>
                    <strong>{brand.product}</strong>
                  </div>
                </div>
                <Field label={t('versionLabel')} value={appVersion || '—'} />
                <Field
                  label={brand.website}
                  value={<BigDogWebsiteLink className="set-about-website" showAddress />}
                />
                <p className="set-about-copy">{brand.productDescription}</p>
                <p className="set-field-desc">{brand.licenses}</p>
                <div className="set-field">
                  <div className="set-field-text">
                    <label className="set-field-label">{t('updateChannel')}</label>
                  </div>
                  <Dropdown
                    className="set-dd"
                    value={channel}
                    ariaLabel={t('updateChannel')}
                    options={CHANNEL_OPTIONS.map((opt) => ({
                      value: opt.value,
                      label: t(opt.labelKey),
                    }))}
                    disabled={releaseBusy}
                    onPick={(v) => void checkRelease(v === 'beta' ? 'beta' : 'stable')}
                  />
                </div>
                <div className="set-about-releases">
                  <strong>{brand.releases}</strong>
                  <p>{brand.releaseDescription}</p>
                  <div className="set-about-actions">
                    <button
                      className="set-btn"
                      disabled={releaseBusy}
                      aria-busy={releaseBusy}
                      onClick={() => void checkRelease()}
                    >
                      {releaseBusy ? brand.checking : brand.checkUpdates}
                    </button>
                    <button
                      className="set-btn"
                      onClick={() => {
                        setReleaseError('')
                        void window.aiOffice
                          .openBigDogReleases(
                            releaseStatus?.phase === 'available'
                              ? releaseStatus.version
                              : undefined,
                          )
                          .catch(() => setReleaseError(brand.checkFailed))
                      }}
                    >
                      {brand.openReleases}
                    </button>
                    <button
                      className="set-btn"
                      onClick={() => {
                        setReleaseError('')
                        void window.aiOffice
                          .openGitHubRepo()
                          .catch(() => setReleaseError(brand.checkFailed))
                      }}
                    >
                      {brand.repository}
                    </button>
                  </div>
                  <p role={releaseError || releaseStatus?.phase === 'error' ? 'alert' : 'status'}>
                    {releaseError ||
                      (releaseStatus?.phase === 'available'
                        ? brand.available(releaseStatus.version)
                        : releaseStatus?.phase === 'current'
                          ? brand.upToDate
                          : releaseStatus?.phase === 'error'
                            ? brand.checkFailed
                            : '')}
                  </p>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
