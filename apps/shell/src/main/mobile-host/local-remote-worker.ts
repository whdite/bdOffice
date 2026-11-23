import { createRemoteServer } from '@bigdogoffice/remote-server'
import { LOCAL_REMOTE_PORT } from '../../shared/mobile-host-api'

let server: ReturnType<typeof createRemoteServer> | undefined
let stopping = false
async function stop(code = 0) {
  if (stopping) return
  stopping = true
  await server?.close().catch(() => {})
  process.exit(code)
}
process.once('disconnect', () => void stop())
process.on('message', (message: unknown) => {
  if (!message || typeof message !== 'object') return
  const request = message as {
    type?: string
    account?: string
    password?: string
    dataFile?: string
    serverUrl?: string
  }
  if (request.type === 'stop') {
    void stop()
    return
  }
  if (
    stopping ||
    server ||
    request.type !== 'start' ||
    typeof request.account !== 'string' ||
    typeof request.password !== 'string' ||
    typeof request.dataFile !== 'string' ||
    request.serverUrl !== `http://127.0.0.1:${LOCAL_REMOTE_PORT}`
  )
    return
  try {
    server = createRemoteServer({
      account: request.account,
      password: request.password,
      dataFile: request.dataFile,
      publicUrl: request.serverUrl,
      logLevel: 'error',
    })
    server.server.once('error', (error: NodeJS.ErrnoException) => {
      process.send?.(
        {
          ready: false,
          code: error.code === 'EADDRINUSE' ? 'REMOTE_PORT_IN_USE' : 'REMOTE_UNAVAILABLE',
        },
        () => void stop(1),
      )
    })
    server.server.listen(LOCAL_REMOTE_PORT, '127.0.0.1', () => process.send?.({ ready: true }))
  } catch {
    process.send?.({ ready: false, code: 'REMOTE_UNAVAILABLE' }, () => void stop(1))
  }
})
