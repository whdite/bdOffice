/** Public distribution only. Never point clients at the private build repository. */
export const BIGDOG_REPOSITORY = 'whdite/bdOffice'
export const BIGDOG_WEBSITE_URL = 'https://bdoffice.deepthink-x.com/'
export const BIGDOG_REPOSITORY_URL = `https://github.com/${BIGDOG_REPOSITORY}`
export const BIGDOG_RELEASES_URL = `${BIGDOG_REPOSITORY_URL}/releases`
export const BIGDOG_RELEASE_CATALOG_URL = `https://api.github.com/repos/${BIGDOG_REPOSITORY}/releases?per_page=20`
export const BIGDOG_RELEASE_ATOM_URL = `${BIGDOG_RELEASES_URL}.atom`
export const BIGDOG_RELEASE_MANIFEST_URL = `${BIGDOG_RELEASES_URL}/latest/download/release-manifest.json`

export function bigDogReleasePage(version?: unknown): string {
  return typeof version === 'string' &&
    /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/.test(
      version,
    )
    ? `${BIGDOG_RELEASES_URL}/tag/v${version}`
    : BIGDOG_RELEASES_URL
}

export type BigDogReleaseStatus =
  { phase: 'available'; version: string } | { phase: 'current' } | { phase: 'error' }
