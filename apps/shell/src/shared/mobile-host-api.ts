/** Development-only loopback protocol. It is not the production encrypted Remote transport. */
export const MOBILE_HOST_PROTOCOL = 'bigdog.mobile-dev.v1'
export const MOBILE_HOST_DEFAULT_PORT = 48219
export const LOCAL_REMOTE_PORT = 48300
export type RemoteServerMode = 'local' | 'external'
export type RemoteTransportPreference = 'auto' | 'relay'
export type RemoteRoute = 'lan' | 'p2p' | 'turn' | 'relay'
export interface RemoteLaunchInput {
  mode: RemoteServerMode
  autoStart: boolean
  transport?: RemoteTransportPreference
  serverUrl?: string
  account?: string
  password?: string
}
export interface RemoteLaunchSettings {
  mode: RemoteServerMode
  autoStart: boolean
  transport: RemoteTransportPreference
  serverUrl: string
  account: string
  hasPassword: boolean
}
export type RemoteAdminError =
  | 'REMOTE_CONFIG_INVALID'
  | 'REMOTE_SETTINGS_UNAVAILABLE'
  | 'REMOTE_PORT_IN_USE'
  | 'REMOTE_AUTH_REQUIRED'
  | 'REMOTE_AUTH_FAILED'
  | 'REMOTE_UNAVAILABLE'
  | 'REMOTE_DISABLED'
  | 'REMOTE_DISCONNECTED'
  | 'REMOTE_ENV_MANAGED'
export interface RemoteAdminStatus {
  enabled: boolean
  online: boolean
  canInvite: boolean
  phase: 'stopped' | 'starting' | 'running' | 'stopping'
  settings: RemoteLaunchSettings
  managedByEnvironment: boolean
  error?: RemoteAdminError
  serverUrl?: string
  name?: string
  routes: RemoteRoute[]
  invitationState?: 'ready' | 'used' | 'expired'
  devices: Array<{ id: string; name: string; revoked: boolean; lastSeen: number }>
}
export interface RemoteAdminApi {
  status(): Promise<RemoteAdminStatus>
  save(input: RemoteLaunchInput): Promise<RemoteAdminStatus>
  start(input?: RemoteLaunchInput): Promise<RemoteAdminStatus>
  stop(): Promise<RemoteAdminStatus>
  invite(): Promise<{ code: string; expiresAt: number; qrDataUrl?: string }>
  saveQr(): Promise<boolean>
  revoke(deviceId: string): Promise<void>
}
export const REMOTE_ADMIN_CHANNELS = {
  status: 'remote-admin:status',
  save: 'remote-admin:save',
  start: 'remote-admin:start',
  stop: 'remote-admin:stop',
  invite: 'remote-admin:invite',
  saveQr: 'remote-admin:save-qr',
  revoke: 'remote-admin:revoke',
} as const
declare global {
  interface Window {
    aiOfficeRemote: RemoteAdminApi
  }
}
export interface MobilePairReply {
  token: string
  deviceId: string
  expiresAt: number
}
export interface MobileMessage {
  id: string
  role: 'user' | 'assistant'
  text: string
  at: string | null
}
export interface MobileArtifact {
  id: string
  name: string
  mime: string
  bytes: number
  digest: string
}
export interface MobileRpcEnvelope {
  version: 1
  id: string
  method: string
  params?: Record<string, unknown>
}
export type MobileRpcReply =
  | { version: 1; id: string; ok: true; value: unknown }
  | { version: 1; id: string; ok: false; error: { code: string; message: string } }
