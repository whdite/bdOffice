import { dialog, ipcMain, type WebContents } from 'electron'
import { writeFile } from 'node:fs/promises'
import QRCode from 'qrcode'
import { encodePairingInvitation } from '@bigdogoffice/remote-core/pairing'
import { assertTrustedSettingsSender } from '../trusted-settings-sender'
import { REMOTE_ADMIN_CHANNELS } from '../../shared/mobile-host-api'
import { remoteAdminError, type DesktopRemoteService } from './desktop-service'
export function registerRemoteAdmin(options: {
  contents(): WebContents | undefined
  service(): DesktopRemoteService
}) {
  let latestQr:
    | { png: string; expiresAt: number; runtime: NonNullable<DesktopRemoteService['mounted']> }
    | undefined
  let issuing = false
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.status, (event) => {
    assertTrustedSettingsSender(event, options.contents())
    return options.service().status()
  })
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.save, (event, input: unknown) => {
    assertTrustedSettingsSender(event, options.contents())
    return options.service().save(input)
  })
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.start, (event, input: unknown) => {
    assertTrustedSettingsSender(event, options.contents())
    return options.service().start(input)
  })
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.stop, (event) => {
    assertTrustedSettingsSender(event, options.contents())
    latestQr = undefined
    return options.service().stop()
  })
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.invite, async (event) => {
    assertTrustedSettingsSender(event, options.contents())
    const service = options.service()
    const runtime = service.mounted
    const host = runtime?.gateway
    if (!runtime || !host) throw new Error('REMOTE_DISABLED')
    const remote = runtime.remote
    if (!remote) return host.createInvitation()
    if (!remote.host.isOnline()) throw new Error('REMOTE_DISCONNECTED')
    if (issuing) throw new Error('REMOTE_UNAVAILABLE')
    issuing = true
    try {
      const ticket = await remote.api.createPairing()
      if (service.mounted !== runtime) throw new Error('REMOTE_DISCONNECTED')
      const local = host.createInvitation()
      latestQr = undefined
      const expiresAt = Math.min(ticket.expiresAt, local.expiresAt)
      const code = encodePairingInvitation({
        v: 1,
        server: remote.api.baseUrl,
        host: remote.identity.deviceId,
        name: remote.identity.name,
        key: remote.identity.publicKey,
        code: local.code,
        ticket: ticket.ticket,
        expires: expiresAt,
      })
      const qrDataUrl = await QRCode.toDataURL(code, {
        width: 640,
        margin: 3,
        errorCorrectionLevel: 'M',
      })
      if (service.mounted !== runtime) throw new Error('REMOTE_DISCONNECTED')
      latestQr = { png: qrDataUrl, expiresAt, runtime }
      return { code, expiresAt, qrDataUrl }
    } catch (error) {
      throw new Error(service.mounted !== runtime ? 'REMOTE_DISCONNECTED' : remoteAdminError(error))
    } finally {
      issuing = false
    }
  })
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.saveQr, async (event) => {
    assertTrustedSettingsSender(event, options.contents())
    const current = latestQr
    if (
      !current ||
      current.runtime !== options.service().mounted ||
      current.expiresAt <= Date.now()
    )
      throw new Error('PAIRING_EXPIRED')
    const result = await dialog.showSaveDialog({
      title: '保存手机配对二维码',
      defaultPath: '大狗手机配对.png',
      filters: [{ name: 'PNG', extensions: ['png'] }],
    })
    if (result.canceled || !result.filePath) return false
    if (
      current !== latestQr ||
      current.runtime !== options.service().mounted ||
      current.expiresAt <= Date.now()
    )
      throw new Error('PAIRING_EXPIRED')
    await writeFile(result.filePath, Buffer.from(current.png.split(',')[1]!, 'base64'))
    return true
  })
  ipcMain.handle(REMOTE_ADMIN_CHANNELS.revoke, (event, id: unknown) => {
    assertTrustedSettingsSender(event, options.contents())
    if (typeof id !== 'string' || !/^[a-f0-9]{64}$/.test(id)) throw new Error('INVALID_DEVICE')
    options.service().mounted?.gateway?.identities?.revoke(id)
  })
}
