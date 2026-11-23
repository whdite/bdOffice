import { useCallback, useEffect, useRef, useState } from 'react'
import { useI18n } from '../locale'
import { settingsCopy } from '../settings-copy'
import {
  LOCAL_REMOTE_PORT,
  type RemoteAdminError,
  type RemoteAdminStatus,
  type RemoteLaunchInput,
  type RemoteRoute,
} from '../../../shared/mobile-host-api'

type LaunchDraft = Required<RemoteLaunchInput>
const initialDraft: LaunchDraft = {
  mode: 'external',
  autoStart: false,
  transport: 'auto',
  serverUrl: '',
  account: '',
  password: '',
}

const routeLabels: Record<RemoteRoute, [string, string]> = {
  lan: ['局域网直连', 'LAN'],
  p2p: ['P2P 直连', 'P2P'],
  turn: ['TURN 中继', 'TURN'],
  relay: ['服务器转发', 'Server relay'],
}

const remoteErrors: Record<RemoteAdminError, [string, string]> = {
  REMOTE_CONFIG_INVALID: [
    '请检查 Remote 地址和配置。首次授权需要账号及至少 12 位的密码。',
    'Check the Remote address and settings. First-time authorization requires an account and a password of at least 12 characters.',
  ],
  REMOTE_SETTINGS_UNAVAILABLE: [
    '无法读取或保存凭据。请检查本机系统密钥库。',
    'The system credential store could not read or save the settings.',
  ],
  REMOTE_PORT_IN_USE: [
    `端口 ${LOCAL_REMOTE_PORT} 或启动配置中的端口已被占用。`,
    `Port ${LOCAL_REMOTE_PORT} or a port specified in the startup configuration is in use.`,
  ],
  REMOTE_AUTH_REQUIRED: [
    '此电脑尚未获得授权。请填写 Remote 账号和密码后重试。',
    'Enter the Remote account and password to authorize this desktop.',
  ],
  REMOTE_AUTH_FAILED: [
    'Remote 授权失败。请核对账号和密码。',
    'Remote authorization failed. Check the account and password.',
  ],
  REMOTE_UNAVAILABLE: [
    '无法连接 Remote。请检查服务是否运行、地址和网络是否可用。',
    'Remote could not be reached. Check the service, address, and network.',
  ],
  REMOTE_DISABLED: ['请先启动 Remote。', 'Start Remote first.'],
  REMOTE_DISCONNECTED: [
    'Remote 连接已断开，请在连接恢复后重新生成。',
    'Remote disconnected. Create a new invitation once it reconnects.',
  ],
  REMOTE_ENV_MANAGED: [
    '当前连接由启动参数管理。',
    'This connection is managed by the startup configuration.',
  ],
}

function remoteErrorText(reason: unknown, zh: boolean, fallback: string) {
  const code = /\b(REMOTE_[A-Z_]+)\b/.exec(String(reason))?.[1] as RemoteAdminError | undefined
  return (code && remoteErrors[code]?.[zh ? 0 : 1]) || fallback
}

function toDraft(settings: RemoteAdminStatus['settings']): LaunchDraft {
  return {
    mode: settings.mode,
    autoStart: settings.autoStart,
    transport: settings.transport,
    serverUrl: settings.serverUrl,
    account: settings.account,
    password: '',
  }
}

export function RemotePanel() {
  const { lang } = useI18n(),
    zh = lang.startsWith('zh'),
    copy = settingsCopy(lang)
  const [status, setStatus] = useState<RemoteAdminStatus>()
  const [draft, setDraft] = useState(initialDraft)
  const [invite, setInvite] = useState<{ code: string; expiresAt: number; qrDataUrl?: string }>()
  const [loading, setLoading] = useState(true)
  const [busy, setBusy] = useState<string | null>(null)
  const [error, setError] = useState('')
  const [notice, setNotice] = useState('')
  const [now, setNow] = useState(Date.now)
  const alive = useRef(true),
    lock = useRef(false),
    hydrated = useRef(false),
    revision = useRef(0)
  const refresh = useCallback(async () => {
    const running = ++revision.current
    setLoading(true)
    setError('')
    try {
      const next = await window.aiOfficeRemote.status()
      if (alive.current && revision.current === running) {
        setStatus(next)
        if (!hydrated.current) {
          setDraft(toDraft(next.settings))
          hydrated.current = true
        }
      }
    } catch {
      if (alive.current)
        setError(
          zh ? '无法读取连接状态，请重试。' : 'Could not read connection status. Please retry.',
        )
    } finally {
      if (alive.current) setLoading(false)
    }
  }, [zh])
  useEffect(() => {
    alive.current = true
    void refresh()
    return () => {
      alive.current = false
    }
  }, [refresh])
  useEffect(() => {
    if (!invite) return
    setNow(Date.now())
    const timer = setInterval(() => setNow(Date.now()), 1000)
    return () => clearInterval(timer)
  }, [invite])
  useEffect(() => {
    if (loading) return
    let disposed = false,
      pending = false
    const timer = setInterval(async () => {
      if (pending || lock.current) return
      pending = true
      const running = revision.current
      try {
        const next = await window.aiOfficeRemote.status()
        if (!disposed && revision.current === running) {
          setStatus(next)
          if (!hydrated.current) {
            setDraft(toDraft(next.settings))
            hydrated.current = true
          }
        }
      } catch {
        /* Keep the invite usable through a transient status failure. */
      } finally {
        pending = false
      }
    }, 3000)
    return () => {
      disposed = true
      clearInterval(timer)
    }
  }, [loading])
  const run = async (id: string, action: () => Promise<void>, failure: string) => {
    if (lock.current) return
    ++revision.current
    lock.current = true
    setBusy(id)
    setError('')
    setNotice('')
    try {
      await action()
    } catch (reason) {
      if (alive.current) setError(remoteErrorText(reason, zh, failure))
    } finally {
      lock.current = false
      if (alive.current) setBusy(null)
    }
  }
  const expired = !!invite && now >= invite.expiresAt
  const consumed = !!invite && status?.invitationState === 'used'
  const loopbackInvite =
    status?.serverUrl &&
    ['localhost', '127.0.0.1', '10.0.2.2'].includes(new URL(status.serverUrl).hostname)
  const busyWithConnection = status?.phase === 'starting' || status?.phase === 'stopping'
  const disabled = !status || loading || !!busy || busyWithConnection
  const statusText = !status
    ? ''
    : status.phase === 'starting'
      ? zh
        ? '正在启动 Remote…'
        : 'Starting Remote…'
      : status.phase === 'stopping'
        ? zh
          ? '正在停止 Remote…'
          : 'Stopping Remote…'
        : status.online
          ? zh
            ? 'Remote 已连接'
            : 'Remote connected'
          : status.enabled
            ? status.serverUrl
              ? zh
                ? '正在连接 Remote…'
                : 'Connecting to Remote…'
              : zh
                ? '本机联调网关已启动'
                : 'Local development gateway running'
            : zh
              ? 'Remote 已停止'
              : 'Remote stopped'
  const configurationError = status?.error ? remoteErrors[status.error][zh ? 0 : 1] : ''
  const start = () =>
    run(
      'start',
      async () => {
        setInvite(undefined)
        const next = await window.aiOfficeRemote.start(
          status?.managedByEnvironment ? undefined : draft,
        )
        if (!alive.current) return
        setStatus(next)
        if (!next.error) setDraft(toDraft(next.settings))
      },
      remoteErrors.REMOTE_UNAVAILABLE[zh ? 0 : 1],
    )
  return (
    <section className="set-remote">
      <h3 className="set-pane-title">{zh ? '手机协作' : 'Phone collaboration'}</h3>
      <p className="set-field-desc" role="status">
        {loading ? copy.loading : statusText}
      </p>
      {status?.serverUrl && (
        <p className="set-field-desc set-remote-address">
          {status.name} · {status.serverUrl}
        </p>
      )}
      {!!status?.routes.length && (
        <p className="set-field-desc" role="status">
          {zh ? '当前手机链路' : 'Active phone routes'} ·{' '}
          {status.routes.map((route) => routeLabels[route][zh ? 0 : 1]).join(' / ')}
        </p>
      )}
      {status && !status.managedByEnvironment && (
        <form
          className="set-remote-form"
          onSubmit={(event) => {
            event.preventDefault()
            void start()
          }}
        >
          <label className="set-field">
            <span className="set-field-label">{zh ? 'Remote 服务' : 'Remote service'}</span>
            <select
              className="set-input"
              value={draft.mode}
              disabled={disabled}
              onChange={(event) => {
                const mode = event.currentTarget.value as LaunchDraft['mode']
                setDraft((previous) => ({ ...previous, mode }))
              }}
            >
              <option value="external">
                {zh ? '连接已有 Remote' : 'Connect to an existing Remote'}
              </option>
              <option value="local">{zh ? '本机调试 Remote' : 'Local development Remote'}</option>
            </select>
          </label>
          {draft.mode === 'local' ? (
            <p className="set-field-desc">
              {zh
                ? `在本机启动 Remote，使用端口 ${LOCAL_REMOTE_PORT}，适用于模拟器或已反向转发端口的设备。`
                : `Starts Remote on this computer at port ${LOCAL_REMOTE_PORT} for emulators or devices with a reverse port forward.`}
            </p>
          ) : (
            <>
              <label className="set-field">
                <span className="set-field-label">{zh ? '服务器地址' : 'Server address'}</span>
                <input
                  className="set-input"
                  type="url"
                  placeholder="https://remote.example.com"
                  value={draft.serverUrl}
                  maxLength={256}
                  required
                  disabled={disabled}
                  onChange={(event) => setDraft({ ...draft, serverUrl: event.currentTarget.value })}
                />
              </label>
              <label className="set-field">
                <span className="set-field-label">{zh ? 'Remote 账号' : 'Remote account'}</span>
                <input
                  className="set-input"
                  value={draft.account}
                  maxLength={254}
                  disabled={disabled}
                  onChange={(event) => setDraft({ ...draft, account: event.currentTarget.value })}
                />
              </label>
              <label className="set-field">
                <span className="set-field-label">{zh ? 'Remote 密码' : 'Remote password'}</span>
                <input
                  className="set-input"
                  type="password"
                  autoComplete="new-password"
                  placeholder={
                    status.settings.hasPassword
                      ? zh
                        ? '使用已保存的密码'
                        : 'Use saved password'
                      : zh
                        ? '首次授权时填写'
                        : 'Required for first-time authorization'
                  }
                  value={draft.password}
                  maxLength={1024}
                  disabled={disabled}
                  onChange={(event) => setDraft({ ...draft, password: event.currentTarget.value })}
                />
              </label>
            </>
          )}
          <label className="set-field">
            <span className="set-field-label">{zh ? '连接路径' : 'Connection path'}</span>
            <select
              className="set-input"
              value={draft.transport}
              disabled={disabled}
              onChange={(event) => {
                const transport = event.currentTarget.value as LaunchDraft['transport']
                setDraft((previous) => ({ ...previous, transport }))
              }}
            >
              <option value="auto">{zh ? '自动选路' : 'Automatic'}</option>
              <option value="relay">{zh ? '服务器转发' : 'Server relay'}</option>
            </select>
          </label>
          <p className="set-field-desc">
            {zh
              ? '自动模式先尝试局域网或 P2P，再尝试 TURN，最后使用服务器转发。'
              : 'Automatic mode tries LAN or P2P, then TURN, then encrypted server relay.'}
          </p>
          <div className="set-field">
            <span className="set-field-label">
              {zh ? '随大狗自动启动 Remote' : 'Start Remote with BigDog'}
            </span>
            <button
              type="button"
              className="set-switch"
              role="switch"
              aria-label={zh ? '随大狗自动启动 Remote' : 'Start Remote with BigDog'}
              aria-checked={draft.autoStart}
              disabled={disabled}
              onClick={() =>
                setDraft((previous) => ({ ...previous, autoStart: !previous.autoStart }))
              }
            />
          </div>
          <div className="set-about-actions">
            <button type="submit" className="set-btn primary" disabled={disabled}>
              {busy === 'start' || busyWithConnection
                ? copy.working
                : zh
                  ? '保存并启动'
                  : 'Save and start'}
            </button>
            <button
              type="button"
              className="set-btn"
              disabled={disabled}
              onClick={() =>
                void run(
                  'save',
                  async () => {
                    const next = await window.aiOfficeRemote.save(draft)
                    if (!alive.current) return
                    setStatus(next)
                    if (!next.error) {
                      setDraft(toDraft(next.settings))
                      setNotice(zh ? '已保存 Remote 设置。' : 'Remote settings saved.')
                    }
                  },
                  zh
                    ? '无法保存 Remote 设置，请重试。'
                    : 'Could not save the Remote settings. Please retry.',
                )
              }
            >
              {zh ? '保存设置' : 'Save settings'}
            </button>
          </div>
        </form>
      )}
      {status?.managedByEnvironment && (
        <p className="set-field-desc">{remoteErrors.REMOTE_ENV_MANAGED[zh ? 0 : 1]}</p>
      )}
      <div className="set-about-actions">
        {status?.managedByEnvironment && !status.enabled && (
          <button className="set-btn" disabled={disabled} onClick={() => void start()}>
            {zh ? '重新连接' : 'Reconnect'}
          </button>
        )}
        {status?.enabled && (
          <button
            className="set-btn"
            disabled={disabled}
            onClick={() =>
              void run(
                'stop',
                async () => {
                  const next = await window.aiOfficeRemote.stop()
                  if (alive.current) {
                    setInvite(undefined)
                    setStatus(next)
                  }
                },
                zh ? 'Remote 未能停止，请重试。' : 'Remote could not stop. Please retry.',
              )
            }
          >
            {zh ? '停止 Remote' : 'Stop Remote'}
          </button>
        )}
        <button
          className="set-btn primary"
          disabled={!status?.canInvite || disabled}
          aria-busy={busy === 'invite'}
          onClick={() =>
            void run(
              'invite',
              async () => {
                const next = await window.aiOfficeRemote.invite()
                if (alive.current) {
                  setInvite(next)
                  setStatus((previous) => previous && { ...previous, invitationState: 'ready' })
                }
              },
              zh
                ? '未能生成配对码，请检查连接后重试。'
                : 'Could not create an invitation. Check the connection and retry.',
            )
          }
        >
          {busy === 'invite' ? copy.working : zh ? '生成手机配对码' : 'Create pairing invitation'}
        </button>
        <button
          className="set-btn"
          disabled={loading || !!busy}
          aria-busy={loading}
          onClick={() => void refresh()}
        >
          {loading ? copy.loading : zh ? '刷新状态' : 'Refresh status'}
        </button>
        <button
          className="set-btn"
          disabled={!!busy}
          onClick={() =>
            void run(
              'guide',
              () => window.aiOffice.openRemoteGuide(),
              zh ? '无法打开说明，请稍后重试。' : 'Could not open the guide. Please retry.',
            )
          }
        >
          {zh ? '配置说明' : 'Setup guide'}
        </button>
      </div>
      {invite && status?.enabled && (
        <div className="set-remote-invite">
          {loopbackInvite && (
            <p className="set-field-desc">
              {zh
                ? '这是本机调试邀请。真机通过 Wi‑Fi 或蜂窝网络配对时，使用手机可访问的 HTTPS Remote。'
                : 'This invitation uses loopback. For Wi-Fi or cellular pairing, use an HTTPS Remote reachable from the phone.'}
            </p>
          )}
          {invite.qrDataUrl && !expired && !consumed && (
            <>
              <img
                className="set-remote-qr"
                src={invite.qrDataUrl}
                alt={zh ? '手机配对二维码' : 'Phone pairing QR code'}
              />
              <p>
                {zh
                  ? '手机打开 AI 大狗，点击「扫描电脑二维码」。'
                  : 'Open AI BigDog on your phone and scan this code.'}
              </p>
              <p className="set-field-desc">
                {zh
                  ? '也可复制完整配对码，粘贴到手机；无需输入服务器地址或账号。仅向自己的手机提供此邀请。'
                  : 'Or paste the full pairing code on your phone. No server address or account entry is needed. Share this invitation only with your own phone.'}
              </p>
            </>
          )}
          {invite.qrDataUrl ? (
            <details>
              <summary>{zh ? '查看完整配对码' : 'Show full pairing code'}</summary>
              <code>{invite.code}</code>
            </details>
          ) : (
            <code>{invite.code}</code>
          )}
          <p className="set-field-desc" role="status">
            {consumed
              ? zh
                ? '手机已配对，可在手机继续对话、查看任务和文件。'
                : 'Phone paired. Continue conversations and view tasks and files on your phone.'
              : expired
                ? zh
                  ? '配对码已过期，请重新生成。'
                  : 'Invitation expired. Create a new one.'
                : zh
                  ? `一次有效，有效期至 ${new Date(invite.expiresAt).toLocaleTimeString(lang)}。`
                  : `Single use, expires at ${new Date(invite.expiresAt).toLocaleTimeString(lang)}.`}
          </p>
          <button
            className="set-btn"
            disabled={expired || consumed || !!busy}
            aria-busy={busy === 'copy'}
            onClick={() =>
              void run(
                'copy',
                async () => {
                  await navigator.clipboard.writeText(invite.code)
                  if (alive.current) setNotice(copy.copied)
                },
                copy.copyFailed,
              )
            }
          >
            {notice === copy.copied ? copy.copied : zh ? '复制配对码' : 'Copy invitation'}
          </button>
          {invite.qrDataUrl && (
            <button
              className="set-btn"
              disabled={expired || consumed || !!busy}
              onClick={() =>
                void run(
                  'save-qr',
                  async () => {
                    if (await window.aiOfficeRemote.saveQr())
                      setNotice(
                        zh
                          ? '二维码已保存，可在手机从图片识别。'
                          : 'QR image saved. Open it from your phone.',
                      )
                  },
                  zh ? '二维码未能保存，请重试。' : 'Could not save the QR image. Please retry.',
                )
              }
            >
              {zh ? '保存二维码' : 'Save QR image'}
            </button>
          )}
        </div>
      )}
      {(error || (!busy && configurationError)) && (
        <p className="set-runtime-error" role="alert">
          {error || configurationError}
        </p>
      )}
      {notice && (
        <p className="set-field-desc" role="status">
          {notice}
        </p>
      )}
      <h4 className="set-pane-subtitle">{zh ? '已授权手机' : 'Authorized phones'}</h4>
      {status && status.devices.length === 0 && (
        <p className="set-field-desc">{zh ? '尚未配对手机。' : 'No phones paired yet.'}</p>
      )}
      {status?.devices.map((device) => (
        <div key={device.id} className="set-field">
          <div className="set-field-stack set-remote-device">
            <span>{device.name}</span>
            <span className="set-field-desc">
              {device.revoked
                ? zh
                  ? '已撤销'
                  : 'Revoked'
                : `${zh ? '最近连接：' : 'Last seen: '}${new Date(device.lastSeen).toLocaleString(lang)}`}
            </span>
          </div>
          <button
            className="set-btn"
            disabled={device.revoked || !!busy || loading}
            aria-busy={busy === device.id}
            onClick={() => {
              if (
                lock.current ||
                !window.confirm(
                  zh
                    ? '撤销这台手机的连接授权？已接收的任务可在本机任务中停止。'
                    : 'Revoke this phone? Accepted tasks can be stopped in the local task list.',
                )
              )
                return
              void run(
                device.id,
                async () => {
                  await window.aiOfficeRemote.revoke(device.id)
                  await refresh()
                  if (alive.current)
                    setNotice(zh ? '已撤销手机授权。' : 'Phone authorization revoked.')
                },
                zh ? '撤销未完成，请重试。' : 'Revocation failed. Please retry.',
              )
            }}
          >
            {busy === device.id ? copy.working : zh ? '撤销' : 'Revoke'}
          </button>
        </div>
      ))}
    </section>
  )
}
