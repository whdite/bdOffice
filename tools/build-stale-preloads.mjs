/**
 * Rebuild app preloads whose sources are newer than the built artifact.
 *
 * In dev mode the shell loads each app's preload from apps/<app>/out/preload/
 * straight off disk; `npm run dev` only starts renderer vite servers and never
 * rebuilds preloads, so after a pull the artifact can silently miss newly added
 * preload APIs (calls fail with "... is not a function" in the renderer console
 * only). The root `predev` hook builds just the stale preload bundles.
 */
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { join, resolve } from 'node:path'

const APPS = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']

/** Newest mtime (ms) under dir, 0 when missing. */
function newestMtime(dir) {
  if (!existsSync(dir)) return 0
  let newest = 0
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name)
    newest = Math.max(newest, entry.isDirectory() ? newestMtime(p) : statSync(p).mtimeMs)
  }
  return newest
}

function fileMtime(path) {
  return existsSync(path) ? statSync(path).mtimeMs : 0
}

const PACKAGE_CONFIG_FILES = ['package.json', 'tsconfig.json']

function workspaceDependencyMtime(dir) {
  return Math.max(
    newestMtime(join(dir, 'src')),
    ...PACKAGE_CONFIG_FILES.map((file) => fileMtime(join(dir, file))),
  )
}

const workspacePackages = new Map()
for (const entry of readdirSync('packages', { withFileTypes: true })) {
  if (!entry.isDirectory()) continue
  const dir = join('packages', entry.name)
  try {
    const manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'))
    if (manifest.name) workspacePackages.set(manifest.name, { dir, manifest })
  } catch {}
}

function workspaceDependencyDirs(app) {
  const manifest = JSON.parse(readFileSync(join('apps', app, 'package.json'), 'utf8'))
  const pending = Object.keys({ ...manifest.dependencies, ...manifest.devDependencies })
  const seen = new Set()
  const dirs = []
  while (pending.length > 0) {
    const name = pending.pop()
    const workspace = workspacePackages.get(name)
    if (!workspace || seen.has(name)) continue
    seen.add(name)
    dirs.push(workspace.dir)
    pending.push(
      ...Object.keys(workspace.manifest.dependencies ?? {}),
      ...Object.keys(workspace.manifest.devDependencies ?? {}),
    )
  }
  return dirs
}

const stale = APPS.filter((app) => {
  const artifact = join('apps', app, 'out', 'preload', 'index.js')
  const built = existsSync(artifact) ? statSync(artifact).mtimeMs : 0
  // shared/ is included: preloads import IPC channel/type modules from there
  const src = Math.max(
    newestMtime(join('apps', app, 'src', 'preload')),
    newestMtime(join('apps', app, 'src', 'shared')),
    fileMtime(join('apps', app, 'package.json')),
    fileMtime(join('apps', app, 'electron.vite.config.ts')),
    fileMtime(join('apps', app, 'tsconfig.json')),
    fileMtime('package.json'),
    fileMtime('tsconfig.base.json'),
    ...workspaceDependencyDirs(app).map(workspaceDependencyMtime),
  )
  return src > built
})

if (stale.length) {
  console.log(`Rebuilding stale preloads: ${stale.join(', ')}`)
  const [{ resolveConfig }, { build }] = await Promise.all([
    import('electron-vite'),
    import('vite'),
  ])
  process.env.NODE_ENV_ELECTRON_VITE = 'production'
  for (const app of stale) {
    const root = resolve('apps', app)
    const { config } = await resolveConfig({ root }, 'build', 'production')
    if (!config?.preload) throw new Error(`Missing electron-vite preload configuration for ${app}`)
    await build({ ...config.preload, root })
  }
}
