import { resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig } from 'electron-vite'

export default defineConfig({
  // Bundle everything into the shell main (same policy as apps/docs): the
  // imported docs/sheets main modules are TS source with no build artifacts,
  // so externalizing them would break Node ESM resolution at runtime.
  main: {
    // node:sqlite is a Node 22+ builtin the bundler's builtin list may predate
    build: {
      // These workspace packages export TypeScript source, including subpath exports.
      // electron-vite 5 externalizes production dependencies by default.
      externalizeDeps: {
        // Preserve runtime detection of optional native modules in ws and the RTC backend.
        include: ['werift', '@trycua/cua-driver', 'ws'],
        exclude: [
          '@bigdogoffice/computer-use',
          '@bigdogoffice/remote-core',
          '@bigdogoffice/remote-server',
          '@dsh-remote/protocol',
          '@dsh-remote/crypto',
          '@dsh-remote/webrtc',
        ],
      },
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/main/index.ts'),
          'computer-use-worker': resolve(__dirname, 'src/main/computer-use-worker.ts'),
          'local-remote-worker': resolve(__dirname, 'src/main/mobile-host/local-remote-worker.ts'),
        },
        external: ['node:sqlite', '@roamhq/wrtc', '@trycua/cua-driver'],
      },
    },
  },
  preload: {
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/preload/index.ts'),
          // dedicated preload for the auto-update window
          update: resolve(__dirname, 'src/preload/update.ts'),
          // dedicated preload for the PDF password prompt window
          'pdf-password': resolve(__dirname, 'src/preload/pdf-password.ts'),
        },
      },
    },
  },
  renderer: {
    plugins: [react()],
    build: {
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'src/renderer/index.html'),
          // strong-guidance update window (see src/main/update-window.ts)
          update: resolve(__dirname, 'src/renderer/update.html'),
          // PDF password prompt window (see src/main/pdf-password-dialog.ts)
          'pdf-password': resolve(__dirname, 'src/renderer/pdf-password.html'),
        },
      },
    },
    server: {
      port: Number(process.env.SHELL_DEV_PORT) || 5199,
      strictPort: Boolean(process.env.SHELL_DEV_PORT),
    },
  },
})
