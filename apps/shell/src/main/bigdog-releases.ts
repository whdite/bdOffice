import { app } from 'electron'
import { parseXml } from 'builder-util-runtime'
import { verifyReleaseManifest } from './release-manifest'
import type { UpdateChannel } from '../shared/update-api'
import {
  BIGDOG_RELEASE_CATALOG_URL,
  BIGDOG_RELEASE_ATOM_URL,
  BIGDOG_RELEASES_URL,
  bigDogReleasePage,
  type BigDogReleaseStatus,
} from '../shared/bigdog-release'

type Version = { core: string[]; pre: string[] }
function parseVersion(value: unknown): Version | null {
  if (typeof value !== 'string' || value.length > 80) return null
  const match =
    /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/.exec(
      value,
    )
  if (!match) return null
  const pre = match[4]?.split('.') ?? []
  if (pre.some((id) => /^0\d+$/.test(id))) return null
  return { core: match.slice(1, 4), pre }
}

function numericCompare(a: string, b: string): number {
  return a.length - b.length || (a < b ? -1 : a > b ? 1 : 0)
}

function compare(a: Version, b: Version): number {
  for (let i = 0; i < 3; i++) {
    const diff = numericCompare(a.core[i], b.core[i])
    if (diff) return diff
  }
  if (!a.pre.length || !b.pre.length) return Number(!a.pre.length) - Number(!b.pre.length)
  for (let i = 0; i < Math.max(a.pre.length, b.pre.length); i++) {
    const x = a.pre[i],
      y = b.pre[i]
    if (x === undefined || y === undefined) return a.pre.length - b.pre.length
    if (x === y) continue
    const xn = /^\d+$/.test(x),
      yn = /^\d+$/.test(y)
    if (xn && yn) return numericCompare(x, y)
    if (xn !== yn) return xn ? -1 : 1
    return x < y ? -1 : 1
  }
  return 0
}

function object(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {}
}

function installers(version: string, platform: string, arch: string) {
  const target = ({ win32: 'windows', darwin: 'macos', linux: 'linux' } as Record<string, string>)[
    platform
  ]
  if (!target || !['x64', 'arm64'].includes(arch)) return []
  const formats =
    platform === 'win32' ? ['exe'] : platform === 'darwin' ? ['dmg'] : ['AppImage', 'deb', 'rpm']
  const archs = platform === 'darwin' ? [arch, 'universal'] : [arch]
  return archs.flatMap((cpu) =>
    formats.map((format) => ({
      name: `BigDogOffice-${version}-${target}-${cpu}.${format}`,
      platform: target,
      arch: cpu,
      format,
    })),
  )
}

/** GitHub metadata is a notification catalog, not an installer command. The
 * client never downloads or executes a binary from this response. */
export function selectBigDogRelease(
  catalog: unknown,
  current: string,
  channel: UpdateChannel,
  platform: string,
  arch: string,
): BigDogReleaseStatus {
  const installed = parseVersion(current)
  const releases = catalog
  if (!installed || !Array.isArray(releases)) throw new Error('Invalid release catalog')
  let latest: { version: string; parsed: Version } | null = null
  for (const item of releases) {
    const release = object(item)
    const version = typeof release.tag_name === 'string' ? release.tag_name.replace(/^v/, '') : ''
    const parsed = parseVersion(version)
    if (
      !parsed ||
      release.tag_name !== `v${version}` ||
      version.includes('+') ||
      release.draft !== false ||
      typeof release.prerelease !== 'boolean' ||
      (channel === 'stable' && release.prerelease) ||
      release.html_url !== bigDogReleasePage(version) ||
      typeof release.published_at !== 'string' ||
      !Number.isFinite(Date.parse(release.published_at))
    )
      continue
    if (channel === 'stable' && parsed.pre.length) continue
    if (!Array.isArray(release.assets)) continue
    const assets = release.assets.map(object)
    const uploaded = (name: string): boolean =>
      assets.some(
        (asset) =>
          asset.name === name &&
          asset.state === 'uploaded' &&
          typeof asset.size === 'number' &&
          Number.isSafeInteger(asset.size) &&
          asset.size > 0 &&
          typeof asset.digest === 'string' &&
          /^sha256:[a-f0-9]{64}$/.test(asset.digest) &&
          asset.browser_download_url === `${BIGDOG_RELEASES_URL}/download/v${version}/${name}`,
      )
    // Drafts and partially uploaded releases must not prompt users to update.
    if (!uploaded('release-manifest.json') || !uploaded('release-manifest.json.minisig')) continue
    if (
      !installers(version, platform, arch).some(
        ({ name }) => uploaded(name) && uploaded(name + '.minisig'),
      )
    )
      continue
    if (compare(parsed, installed) <= 0 || (latest && compare(parsed, latest.parsed) <= 0)) continue
    latest = { version, parsed }
  }
  return latest ? { phase: 'available', version: latest.version } : { phase: 'current' }
}

async function readBounded(response: Response, limit: number): Promise<Buffer> {
  if (!response.ok || !response.body) throw new Error('Release request failed')
  const reader = response.body.getReader(),
    chunks: Uint8Array[] = []
  let bytes = 0
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      bytes += value.byteLength
      if (bytes > limit) throw new Error('Release metadata too large')
      chunks.push(value)
    }
  } finally {
    await reader.cancel().catch(() => {})
    reader.releaseLock()
  }
  return Buffer.concat(chunks)
}

const assetHosts = new Set([
  'release-assets.githubusercontent.com',
  'objects.githubusercontent.com',
  'github-releases.githubusercontent.com',
])

/** The tag and filename come from our own rules. GitHub delegates asset delivery to its HTTPS CDN. */
async function assetResponse(
  version: string,
  name: string,
  method: 'GET' | 'HEAD',
  signal: AbortSignal,
) {
  const url = `${BIGDOG_RELEASES_URL}/download/v${version}/${name}`
  const response = await fetch(url, { method, signal, redirect: 'manual' })
  if (response.status === 404 || response.status === 410) {
    await response.body?.cancel()
    return null
  }
  if (response.status === 200) return response
  const location = response.headers.get('location')
  await response.body?.cancel()
  if (![301, 302, 303, 307, 308].includes(response.status) || !location)
    throw new Error('Release asset unavailable')
  const cdn = new URL(location)
  if (
    cdn.protocol !== 'https:' ||
    cdn.username ||
    cdn.password ||
    cdn.port ||
    !assetHosts.has(cdn.hostname)
  )
    throw new Error('Unexpected release asset host')
  const result = await fetch(cdn.href, { method, signal, redirect: 'error' })
  if (!result.ok) {
    await result.body?.cancel()
    if (result.status === 404 || result.status === 410) return null
    throw new Error('Release asset unavailable')
  }
  return result
}

async function signedReleaseAvailable(
  version: string,
  channel: UpdateChannel,
  platform: string,
  arch: string,
  signal: AbortSignal,
) {
  const [bytes, signature] = await Promise.all(
    ['release-manifest.json', 'release-manifest.json.minisig'].map(async (name, index) => {
      const response = await assetResponse(version, name, 'GET', signal)
      return response && readBounded(response, index === 0 ? 128 * 1024 : 4096)
    }),
  )
  if (!bytes || !signature) return false
  const key = verifyReleaseManifest(
    bytes,
    signature,
    compare(parseVersion(version)!, parseVersion('0.9.0')!) <= 0,
  )
  const manifest = object(JSON.parse(bytes.toString('utf8')))
  if (
    manifest.schemaVersion !== 1 ||
    manifest.product !== 'BigDogOffice' ||
    manifest.version !== version ||
    manifest.tag !== 'v' + version ||
    manifest.releaseUrl !== bigDogReleasePage(version) ||
    manifest.publicKey !== key ||
    !['stable', 'beta'].includes(String(manifest.channel)) ||
    !Array.isArray(manifest.artifacts)
  )
    throw new Error('Invalid release manifest')
  if (channel === 'stable' && manifest.channel !== 'stable') return false
  const artifacts = manifest.artifacts.map(object)
  for (const expected of installers(version, platform, arch)) {
    const url = `${BIGDOG_RELEASES_URL}/download/v${version}/${expected.name}`
    const item = artifacts.find(
      (a) =>
        a.name === expected.name &&
        a.platform === expected.platform &&
        a.arch === expected.arch &&
        a.format === expected.format &&
        a.url === url &&
        a.signatureUrl === url + '.minisig' &&
        typeof a.bytes === 'number' &&
        Number.isSafeInteger(a.bytes) &&
        a.bytes > 0 &&
        typeof a.sha256 === 'string' &&
        /^[a-f0-9]{64}$/.test(a.sha256),
    )
    if (!item) continue
    const [installer, sig] = await Promise.all(
      [expected.name, expected.name + '.minisig'].map((name) =>
        assetResponse(version, name, 'HEAD', signal),
      ),
    )
    if (
      installer &&
      sig &&
      (!installer.headers.has('content-length') ||
        Number(installer.headers.get('content-length')) === item.bytes) &&
      (!sig.headers.has('content-length') || Number(sig.headers.get('content-length')) > 0)
    )
      return true
  }
  return false
}

export async function checkPublicFeed(
  current: string,
  channel: UpdateChannel,
  platform: string,
  arch: string,
): Promise<BigDogReleaseStatus> {
  const installed = parseVersion(current)
  if (!installed) throw new Error('Invalid installed version')
  const signal = AbortSignal.timeout(20_000)
  const response = await fetch(BIGDOG_RELEASE_ATOM_URL, {
    signal,
    redirect: 'error',
    headers: { Accept: 'application/atom+xml' },
  })
  const feed = parseXml((await readBounded(response, 1024 * 1024)).toString('utf8'))
  if (
    feed.name !== 'feed' ||
    feed.attributes?.xmlns !== 'http://www.w3.org/2005/Atom' ||
    feed.elementValueOrEmpty('id') !== 'tag:github.com,2008:' + BIGDOG_RELEASES_URL
  )
    throw new Error('Invalid public release feed')
  const entries = feed.getElements('entry')
  const versions = new Map<string, Version>()
  for (const entry of entries) {
    const href = entry.getElements('link').find((link) => link.attributes?.rel === 'alternate')
      ?.attributes?.href
    const version = href?.slice((BIGDOG_RELEASES_URL + '/tag/v').length)
    const parsed = parseVersion(version)
    if (href && version && parsed && href === bigDogReleasePage(version))
      versions.set(version, parsed)
  }
  if (entries.length && !versions.size) throw new Error('No recognized release metadata')
  const candidates = [...versions].sort((a, b) => compare(b[1], a[1]))
  for (const [version, parsed] of candidates) {
    if (compare(parsed, installed) <= 0) break
    if (channel === 'stable' && parsed.pre.length) continue
    if (await signedReleaseAvailable(version, channel, platform, arch, signal))
      return { phase: 'available', version }
  }
  // GitHub's Atom feed is a recent window. An older installation may need a platform release outside it.
  if (
    (entries.length >= 10 ||
      feed.getElements('link').some((link) => link.attributes?.rel === 'next')) &&
    !candidates.some(([, parsed]) => compare(parsed, installed) <= 0)
  )
    throw new Error('Public release feed is incomplete')
  return { phase: 'current' }
}

const pending = new Map<UpdateChannel, Promise<BigDogReleaseStatus>>()
export function checkBigDogRelease(channel: UpdateChannel): Promise<BigDogReleaseStatus> {
  const existing = pending.get(channel)
  if (existing) return existing
  const request = (async (): Promise<BigDogReleaseStatus> => {
    try {
      const installed = app.getVersion()
      let latest: BigDogReleaseStatus = { phase: 'current' }
      const signal = AbortSignal.timeout(10_000)
      for (let page = 1; ; page++) {
        // GitHub may paginate through a numeric repository URL. Keep every request on our canonical endpoint.
        const url =
          page === 1 ? BIGDOG_RELEASE_CATALOG_URL : `${BIGDOG_RELEASE_CATALOG_URL}&page=${page}`
        const response = await fetch(url, {
          headers: { Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2026-03-10' },
          redirect: 'error',
          signal,
        })
        const baseline = latest.phase === 'available' ? latest.version : installed
        if (
          response.status === 429 ||
          (response.status === 403 &&
            (response.headers.get('x-ratelimit-remaining') === '0' ||
              response.headers.has('retry-after')))
        ) {
          await response.body?.cancel()
          const backup = await checkPublicFeed(baseline, channel, process.platform, process.arch)
          return backup.phase === 'available' ? backup : latest
        }
        if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) {
          await response.body?.cancel()
          throw new Error('Invalid release response')
        }
        const selected = selectBigDogRelease(
          JSON.parse((await readBounded(response, 1024 * 1024)).toString('utf8')),
          baseline,
          channel,
          process.platform,
          process.arch,
        )
        if (selected.phase === 'available') latest = selected
        if (!response.headers.get('link')?.includes('rel="next"')) return latest
      }
    } catch {
      return { phase: 'error' }
    }
  })()
  pending.set(channel, request)
  void request.finally(() => pending.delete(channel))
  return request
}
