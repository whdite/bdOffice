import { deviceTokenPairSchema } from '@dsh-remote/protocol'
import type { RtcIceServer } from '@dsh-remote/webrtc'
import type { HostIdentity } from './host-ports'
export class ServerApiError extends Error {
  readonly phase?: string
  readonly status?: number
  readonly reason?: string
  constructor(
    readonly code: string,
    message: string,
    readonly retryable = false,
    status?: number,
  ) {
    super(message)
    this.status = status
  }
}
export interface AuthorizedPeerDevice {
  deviceId: string
  name: string
  platform: string
  membershipId: string
  identityKey: string
  role: 'host' | 'client'
}
export interface SecureSettings {
  get(key: string): Promise<string | undefined>
  set(key: string, value: string): Promise<void>
}
type Tokens = ReturnType<typeof deviceTokenPairSchema.parse>
/** Account authorization is for the self-hosted carrier only, independent of model-provider accounts. */
export class HostServerApi {
  readonly baseUrl: string
  private tokens?: Tokens
  private identity?: HostIdentity
  private refreshing?: Promise<Tokens>
  constructor(
    url: string,
    readonly storage: SecureSettings,
    readonly role: 'host' | 'client',
    private readonly signal?: AbortSignal,
  ) {
    const parsed = new URL(url)
    if (
      parsed.username ||
      parsed.password ||
      parsed.search ||
      parsed.hash ||
      parsed.pathname !== '/' ||
      (parsed.protocol !== 'https:' &&
        !(
          parsed.protocol === 'http:' &&
          ['localhost', '127.0.0.1', '10.0.2.2'].includes(parsed.hostname)
        ))
    )
      throw new ServerApiError('INVALID_SERVER', 'HTTPS is required for a remote server.')
    this.baseUrl = parsed.origin
  }
  private get storageKey() {
    return `remote-tokens:${this.baseUrl}:${this.role}:${this.identity?.deviceId}`
  }
  private async request(path: string, init: RequestInit = {}, token?: string): Promise<unknown> {
    const deadline = init.signal ?? AbortSignal.timeout(15_000)
    const response = await fetch(this.baseUrl + path, {
      ...init,
      redirect: 'error',
      signal: this.signal ? AbortSignal.any([deadline, this.signal]) : deadline,
      headers: {
        'content-type': 'application/json',
        ...(token ? { authorization: `Bearer ${token}` } : {}),
      },
    })
    const value = await response.json()
    if (!response.ok)
      throw new ServerApiError(
        value?.error?.code || 'SERVER_ERROR',
        'Remote server request failed.',
        response.status >= 500,
        response.status,
      )
    return value
  }
  async authorize(identity: HostIdentity, email: string, password: string) {
    this.identity = identity
    const login = (await this.request('/api/v1/auth/login', {
      method: 'POST',
      body: JSON.stringify({ email, password }),
    })) as { token: string }
    const tokens = await this.request(
      '/api/v1/devices/register',
      {
        method: 'POST',
        body: JSON.stringify({
          v: 1,
          device: {
            deviceId: identity.deviceId,
            name: identity.name,
            role: this.role,
            platform: this.role === 'host' ? 'desktop' : 'android',
            identityKey: identity.publicKey,
            clientVersion: '0.1.0',
          },
        }),
      },
      login.token,
    )
    return this.save(tokens)
  }
  private async save(value: unknown) {
    const tokens = deviceTokenPairSchema.parse(value)
    await this.storage.set(this.storageKey, JSON.stringify(tokens))
    this.tokens = tokens
    return tokens
  }
  async authenticate(identity = this.identity): Promise<Tokens> {
    if (!identity) throw new ServerApiError('ACCOUNT_AUTH_REQUIRED', 'Authorize this device first.')
    this.identity = identity
    if (!this.tokens) {
      const value = await this.storage.get(this.storageKey)
      if (value) this.tokens = deviceTokenPairSchema.parse(JSON.parse(value))
    }
    if (this.tokens && this.tokens.accessTokenExpiresAt > Date.now() + 30_000) return this.tokens
    return this.refreshCredentials()
  }
  async refreshCredentials(rejected?: string): Promise<Tokens> {
    if (this.refreshing) return this.refreshing
    if (
      rejected &&
      this.tokens &&
      this.tokens.accessToken !== rejected &&
      this.tokens.accessTokenExpiresAt > Date.now() + 30_000
    )
      return this.tokens
    const prior = this.tokens
    if (!prior || !this.identity)
      throw new ServerApiError('ACCOUNT_AUTH_REQUIRED', 'Authorize this device first.')
    this.refreshing = (async () => {
      // Persist only after a valid response. Upstream recovers the same successor if a response is lost.
      const rotate = (refreshToken: string) =>
        this.request('/api/v1/auth/refresh', {
          method: 'POST',
          body: JSON.stringify({ deviceId: this.identity!.deviceId, refreshToken }),
        })
      const next = await this.save(await rotate(prior.refreshToken))
      // A lost response can recover an already-expired access token. Persist its
      // valid successor refresh token before rotating again, so crashes are safe.
      return next.accessTokenExpiresAt <= Date.now() + 30_000
        ? this.save(await rotate(next.refreshToken))
        : next
    })().finally(() => {
      this.refreshing = undefined
    })
    return this.refreshing
  }
  async clearAuthorization() {
    this.tokens = undefined
    await this.storage.set(this.storageKey, '')
  }
  private async authenticatedRequest(path: string, init: RequestInit = {}) {
    const tokens = await this.authenticate()
    try {
      return await this.request(path, init, tokens.accessToken)
    } catch (error) {
      if (
        !(error instanceof ServerApiError) ||
        error.status !== 401 ||
        !['AUTH_INVALID', 'TOKEN_EXPIRED'].includes(error.code)
      )
        throw error
      return this.request(
        path,
        init,
        (await this.refreshCredentials(tokens.accessToken)).accessToken,
      )
    }
  }
  async deviceFor(id: string): Promise<AuthorizedPeerDevice> {
    const value = (await this.authenticatedRequest(
      `/api/v1/devices/${encodeURIComponent(id)}`,
    )) as AuthorizedPeerDevice
    if (value.deviceId !== id || !value.identityKey || !value.membershipId)
      throw new ServerApiError('INVALID_PEER', 'Invalid peer descriptor.')
    return value
  }
  async devices() {
    return this.authenticatedRequest('/api/v1/devices')
  }
  async verifyCurrentDevice(identity: HostIdentity) {
    await this.authenticate(identity)
    const device = (await this.authenticatedRequest('/api/v1/me')) as AuthorizedPeerDevice
    if (
      device.deviceId !== identity.deviceId ||
      device.identityKey !== identity.publicKey ||
      device.role !== this.role
    )
      throw new ServerApiError('AUTH_INVALID', 'Remote device authorization changed.', false, 401)
  }
  async createPairing() {
    return this.authenticatedRequest('/api/v1/pairing/create', {
      method: 'POST',
      body: '{}',
    }) as Promise<{ ticket: string; expiresAt: number }>
  }
  async redeemPairing(
    identity: HostIdentity,
    ticket: string,
    proof: { publicKey: string; signature: string; encoding: string },
  ) {
    this.identity = identity
    const result = (await this.request('/api/v1/pairing/redeem', {
      method: 'POST',
      body: JSON.stringify({
        ticket,
        device: {
          deviceId: identity.deviceId,
          identityKey: identity.publicKey,
          name: identity.name,
          role: 'client',
          platform: 'android',
          clientVersion: '0.1.3',
        },
        proof,
      }),
    })) as { tokens: unknown; host: AuthorizedPeerDevice }
    await this.save(result.tokens)
    return result.host
  }
  async courier(action: string, data?: unknown): Promise<any> {
    if (!/^[a-zA-Z0-9_/-]+$/.test(action))
      throw new ServerApiError('INVALID_ACTION', 'Invalid courier action.')
    return this.authenticatedRequest(
      '/api/v1/bigdog/' + action,
      data === undefined ? {} : { method: 'POST', body: JSON.stringify(data) },
    )
  }
  async turnCredentials(id: string): Promise<RtcIceServer[]> {
    const value = (await this.authenticatedRequest(
      `/api/v1/turn/credentials?connection_id=${encodeURIComponent(id)}`,
    )) as { iceServers: RtcIceServer[] }
    if (!Array.isArray(value.iceServers))
      throw new ServerApiError('INVALID_ICE', 'Invalid ICE configuration.')
    return value.iceServers
  }
}
