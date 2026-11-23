import { spawn } from 'node:child_process'
import { join } from 'node:path'
import type { RemoteAdminError } from '../../shared/mobile-host-api'

export async function startLocalRemote(
  config: { account: string; password: string; dataFile: string; serverUrl: string },
  signal: AbortSignal,
) {
  signal.throwIfAborted()
  const worker = spawn(process.execPath, [join(__dirname, 'local-remote-worker.js')], {
    env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
    stdio: ['ignore', 'ignore', 'ignore', 'ipc'],
    windowsHide: true,
  })
  const closed = new Promise<void>((resolve) => {
    worker.once('exit', () => resolve())
    worker.once('error', () => resolve())
  })
  let stopping: Promise<void> | undefined
  const close = () =>
    (stopping ??= (async () => {
      if (worker.connected) worker.send({ type: 'stop' }, () => {})
      const timer = setTimeout(() => {
        if (worker.exitCode === null && worker.signalCode === null) worker.kill()
      }, 2000)
      timer.unref()
      try {
        await closed
      } finally {
        clearTimeout(timer)
      }
    })())
  try {
    await new Promise<void>((resolve, reject) => {
      const fail = (code: RemoteAdminError) => reject(Object.assign(new Error(code), { code }))
      const abort = () => reject(signal.reason)
      const timer = setTimeout(() => fail('REMOTE_UNAVAILABLE'), 8000)
      const done = () => {
        clearTimeout(timer)
        signal.removeEventListener('abort', abort)
      }
      signal.addEventListener('abort', abort, { once: true })
      void closed.then(() => {
        done()
        fail('REMOTE_UNAVAILABLE')
      })
      worker.once('message', (message: { ready?: unknown; code?: unknown }) => {
        done()
        if (message.ready === true) resolve()
        else fail(message.code === 'REMOTE_PORT_IN_USE' ? message.code : 'REMOTE_UNAVAILABLE')
      })
      worker.send({ type: 'start', ...config }, (error) => {
        if (error) {
          done()
          reject(error)
        }
      })
    })
    signal.throwIfAborted()
    return { close, running: () => worker.connected && worker.exitCode === null }
  } catch (error) {
    await close()
    throw error
  }
}
