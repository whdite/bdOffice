import { randomBytes } from 'node:crypto'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { pairingServer } from '@bigdogoffice/remote-core/pairing'
import {
  LOCAL_REMOTE_PORT,
  MOBILE_HOST_DEFAULT_PORT,
  type RemoteAdminError,
  type RemoteAdminStatus,
  type RemoteLaunchSettings,
  type RemoteRoute,
} from '../../shared/mobile-host-api'
import { RemoteIdentityStore } from './identity-store'
import { startLocalRemote } from './local-remote'
import {
  readPreviousRemoteOrigin,
  readRemoteConfig,
  RemoteVault,
  type connectDesktopRemote,
  type RemoteConfig,
} from './remote'
import { MobileDevHost, type MobileHostDependencies } from './server'

type SavedSettings = Omit<RemoteLaunchSettings, 'hasPassword'> & { password?: string }
type DesktopConnection = Awaited<ReturnType<typeof connectDesktopRemote>>
type LocalServer = Awaited<ReturnType<typeof startLocalRemote>>
type Runtime = {
  gateway?: MobileDevHost
  remote?: DesktopConnection
  localServer?: LocalServer
  settings?: SavedSettings
  requests: AbortController
}

class SetupError extends Error {
  constructor(readonly code: RemoteAdminError) {
    super(code)
  }
}

export function remoteAdminError(error: unknown): RemoteAdminError {
  if (error instanceof SetupError) return error.code
  const cause = error as { code?: string; status?: number; message?: string } | null
  if (cause?.code === 'EADDRINUSE' || cause?.code === 'REMOTE_PORT_IN_USE')
    return 'REMOTE_PORT_IN_USE'
  if (cause?.code === 'INVALID_SERVER') return 'REMOTE_CONFIG_INVALID'
  if (cause?.code === 'ACCOUNT_AUTH_REQUIRED') return 'REMOTE_AUTH_REQUIRED'
  if (cause?.status === 401 || cause?.status === 403) return 'REMOTE_AUTH_FAILED'
  if (
    cause?.message === 'OS credential protection unavailable' ||
    cause?.message === 'An OS keyring is required'
  )
    return 'REMOTE_SETTINGS_UNAVAILABLE'
  return 'REMOTE_UNAVAILABLE'
}

function normalizeInput(input: unknown, previous: SavedSettings): SavedSettings {
  const value = input as Partial<SavedSettings> | null
  if (!value || typeof value.autoStart !== 'boolean') throw new SetupError('REMOTE_CONFIG_INVALID')
  const transport = value.transport ?? previous.transport
  if (transport !== 'auto' && transport !== 'relay') throw new SetupError('REMOTE_CONFIG_INVALID')
  if (value.mode === 'local')
    return { ...previous, mode: 'local', autoStart: value.autoStart, transport }
  if (
    value.mode !== 'external' ||
    typeof value.serverUrl !== 'string' ||
    value.serverUrl.length > 256 ||
    (value.account !== undefined &&
      (typeof value.account !== 'string' || value.account.length > 254)) ||
    (value.password !== undefined &&
      (typeof value.password !== 'string' || value.password.length > 1024))
  )
    throw new SetupError('REMOTE_CONFIG_INVALID')
  let serverUrl: string
  try {
    serverUrl = pairingServer(value.serverUrl.trim())
  } catch {
    throw new SetupError('REMOTE_CONFIG_INVALID')
  }
  const account = value.account?.trim() ?? ''
  const password =
    value.password ||
    (serverUrl === previous.serverUrl && account === previous.account
      ? previous.password
      : undefined)
  if (password && (!account || password.length < 12)) throw new SetupError('REMOTE_CONFIG_INVALID')
  return { mode: 'external', autoStart: value.autoStart, transport, serverUrl, account, password }
}

export class DesktopRemoteService {
  private readonly settingsPath: string
  private vault?: RemoteVault
  private settingsRead?: Promise<SavedSettings>
  private identities?: RemoteIdentityStore
  private tail = Promise.resolve()
  private runtime?: Runtime
  private starting?: Runtime
  private phase: RemoteAdminStatus['phase'] = 'stopped'
  private error?: RemoteAdminError
  private closing = false

  constructor(
    private readonly userData: string,
    private readonly dependencies: () => MobileHostDependencies,
    private readonly connect: (
      config: RemoteConfig,
      port: number,
      gateway: MobileDevHost,
      options: { signal: AbortSignal; authorize: boolean; local: boolean },
    ) => Promise<DesktopConnection>,
    private readonly environment: NodeJS.ProcessEnv = process.env,
  ) {
    this.settingsPath = join(userData, 'mobile', 'remote-settings.vault')
  }

  get mounted() {
    return this.runtime
  }

  private get managedByEnvironment() {
    return Boolean(
      this.environment.BIGDOG_REMOTE_CONFIG || this.environment.BIGDOG_MOBILE_DEV === '1',
    )
  }

  private defaults(): SavedSettings {
    const previous = readPreviousRemoteOrigin(this.userData)
    return {
      mode: 'external',
      autoStart: false,
      transport: 'auto',
      serverUrl: previous?.startsWith('https:') ? previous : '',
      account: '',
    }
  }

  private secureSettings() {
    try {
      return (this.vault ??= new RemoteVault(this.settingsPath))
    } catch {
      throw new SetupError('REMOTE_SETTINGS_UNAVAILABLE')
    }
  }

  private readSettings() {
    return (this.settingsRead ??= (async () => {
      const defaults = this.defaults()
      if (!existsSync(this.settingsPath)) return defaults
      try {
        const raw = await this.secureSettings().get('launch')
        if (!raw) return defaults
        const record = JSON.parse(raw) as SavedSettings
        const external = record.serverUrl
          ? normalizeInput({ ...record, mode: 'external' }, defaults)
          : defaults
        return normalizeInput(record, external)
      } catch {
        this.settingsRead = undefined
        throw new SetupError('REMOTE_SETTINGS_UNAVAILABLE')
      }
    })())
  }

  async status(): Promise<RemoteAdminStatus> {
    let error: RemoteAdminError | undefined
    const { password, ...settings } = await this.readSettings().catch(() => {
      if (!this.managedByEnvironment) error = 'REMOTE_SETTINGS_UNAVAILABLE'
      return this.defaults()
    })
    const runtime = this.runtime
    const localAvailable = !runtime?.localServer || runtime.localServer.running()
    const online = localAvailable && (runtime?.remote?.host.isOnline() ?? false)
    const routes = online
      ? (runtime?.remote?.host.connectionModes() ?? [])
          .map((item) => item.mode?.toLowerCase())
          .filter(
            (mode): mode is RemoteRoute =>
              mode === 'lan' || mode === 'p2p' || mode === 'turn' || mode === 'relay',
          )
      : []
    return {
      enabled: Boolean(runtime),
      online,
      canInvite: Boolean(runtime && localAvailable && (!runtime.remote || online)),
      phase: this.phase,
      settings: { ...settings, hasPassword: Boolean(password) },
      managedByEnvironment: this.managedByEnvironment,
      error: this.error ?? (!localAvailable ? 'REMOTE_UNAVAILABLE' : error),
      serverUrl: runtime?.remote?.api.baseUrl,
      name: runtime?.remote?.identity.name,
      routes: [...new Set(routes)],
      invitationState: runtime?.gateway?.invitationState(),
      devices: (runtime?.gateway?.identities?.list() ?? []).map((row) => ({
        id: String(row.id),
        name: String(row.name),
        revoked: Boolean(row.revoked),
        lastSeen: Number(row.lastSeen),
      })),
    }
  }

  private serialize(action: () => Promise<void>) {
    const operation = this.tail.then(action)
    this.tail = operation.catch(() => {})
    return operation
  }

  startOnLaunch() {
    return this.serialize(async () => {
      try {
        if (this.closing) return
        if (this.managedByEnvironment) await this.launch(undefined, false)
        else {
          const settings = await this.readSettings()
          if (settings.autoStart) await this.launch(settings, false)
        }
      } catch (error) {
        this.error = remoteAdminError(error)
      }
    })
  }

  private async saveSettings(input: unknown) {
    if (this.managedByEnvironment) throw new SetupError('REMOTE_ENV_MANAGED')
    const settings = normalizeInput(input, await this.readSettings())
    try {
      await this.secureSettings().set('launch', JSON.stringify(settings))
    } catch {
      throw new SetupError('REMOTE_SETTINGS_UNAVAILABLE')
    }
    this.settingsRead = Promise.resolve(settings)
    return settings
  }

  async save(input: unknown): Promise<RemoteAdminStatus> {
    await this.serialize(async () => {
      this.error = undefined
      try {
        if (this.closing) throw new SetupError('REMOTE_DISABLED')
        await this.saveSettings(input)
      } catch (error) {
        this.error = remoteAdminError(error)
      }
    })
    return this.status()
  }

  async start(input?: unknown): Promise<RemoteAdminStatus> {
    await this.serialize(async () => {
      this.error = undefined
      try {
        if (this.closing) throw new SetupError('REMOTE_DISABLED')
        if (this.managedByEnvironment) {
          if (input !== undefined) throw new SetupError('REMOTE_ENV_MANAGED')
          await this.launch(undefined, false)
          return
        }
        const settings =
          input === undefined ? await this.readSettings() : await this.saveSettings(input)
        const authorize =
          settings.mode === 'external' &&
          Boolean((input as { password?: string } | undefined)?.password)
        await this.launch(settings, authorize)
      } catch (error) {
        this.error = remoteAdminError(error)
      }
    })
    return this.status()
  }

  async stop(): Promise<RemoteAdminStatus> {
    this.starting?.requests.abort()
    await this.serialize(async () => {
      this.error = undefined
      await this.stopRuntime()
    })
    return this.status()
  }

  dispose() {
    this.closing = true
    this.starting?.requests.abort()
    this.runtime?.requests.abort()
    return this.serialize(async () => {
      await this.stopRuntime()
      this.identities?.close()
      this.identities = undefined
    })
  }

  private async release(runtime: Runtime) {
    runtime.requests.abort()
    await runtime.remote?.stop().catch(() => {})
    await runtime.gateway?.close().catch(() => {})
    await runtime.localServer?.close().catch(() => {})
  }

  private async stopRuntime() {
    const runtime = this.runtime
    if (!runtime) return
    this.runtime = undefined
    this.phase = 'stopping'
    await this.release(runtime)
    this.phase = 'stopped'
  }

  private async launch(settings: SavedSettings | undefined, authorize: boolean) {
    const prior = this.runtime?.settings
    if (
      !authorize &&
      this.runtime?.remote?.host.isOnline() &&
      (this.runtime.localServer?.running() ?? true) &&
      prior?.mode === settings?.mode &&
      prior?.transport === settings?.transport &&
      (settings?.mode !== 'external' ||
        (prior?.serverUrl === settings.serverUrl &&
          prior.account === settings.account &&
          prior.password === settings.password))
    )
      return
    await this.stopRuntime()
    if (this.closing) return
    const runtime: Runtime = { requests: new AbortController(), settings }
    this.starting = runtime
    this.phase = 'starting'
    try {
      const config = await this.connectionConfig(settings, runtime)
      const port = settings
        ? 0
        : Number(this.environment.BIGDOG_MOBILE_PORT ?? MOBILE_HOST_DEFAULT_PORT)
      if (!settings && (!Number.isInteger(port) || port < 1 || port > 65535))
        throw new SetupError('REMOTE_CONFIG_INVALID')
      runtime.requests.signal.throwIfAborted()
      const code = config?.pairingCode ?? this.environment.BIGDOG_MOBILE_PAIRING_CODE ?? ''
      if (!/^[^\s\x00-\x1f\x7f]{6,128}$/.test(code)) throw new SetupError('REMOTE_CONFIG_INVALID')
      const db = (this.identities ??= new RemoteIdentityStore(
        join(this.userData, 'mobile', 'identity.sqlite'),
      ))
      runtime.gateway = new MobileDevHost(this.dependencies(), code, Date.now, db)
      const boundPort = await runtime.gateway.start(port)
      runtime.requests.signal.throwIfAborted()
      if (config) {
        runtime.remote = await this.connect(config, boundPort, runtime.gateway, {
          signal: runtime.requests.signal,
          authorize,
          local: settings?.mode === 'local',
        })
      }
      runtime.requests.signal.throwIfAborted()
      this.runtime = runtime
      this.phase = 'running'
    } catch (error) {
      await this.release(runtime)
      this.phase = 'stopped'
      throw error
    } finally {
      this.starting = undefined
    }
  }

  private async connectionConfig(settings: SavedSettings | undefined, runtime: Runtime) {
    if (!settings) {
      const config = this.environment.BIGDOG_REMOTE_CONFIG
      if (!config) return undefined
      try {
        return readRemoteConfig(config)
      } catch {
        throw new SetupError('REMOTE_CONFIG_INVALID')
      }
    }
    const pairingCode = randomBytes(16).toString('hex')
    if (settings.mode === 'external') {
      return {
        serverUrl: settings.serverUrl,
        account: settings.account,
        password: settings.password,
        pairingCode,
        forceRelay: settings.transport === 'relay',
      }
    }
    const vault = this.secureSettings()
    let password: string
    try {
      const saved = await vault.get('local-remote-password')
      password = saved || randomBytes(32).toString('base64url')
      if (!saved) await vault.set('local-remote-password', password)
    } catch {
      throw new SetupError('REMOTE_SETTINGS_UNAVAILABLE')
    }
    const account = 'bigdog-desktop@localhost'
    const serverUrl = `http://127.0.0.1:${LOCAL_REMOTE_PORT}`
    runtime.localServer = await startLocalRemote(
      {
        account,
        password,
        serverUrl,
        dataFile: join(this.userData, 'mobile', 'local-remote', 'state.json'),
      },
      runtime.requests.signal,
    )
    return { serverUrl, account, password, pairingCode, forceRelay: settings.transport === 'relay' }
  }
}
