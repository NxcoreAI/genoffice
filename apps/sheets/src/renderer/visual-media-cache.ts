import type { WorkbookMediaRequest } from '../shared/desktop-api'

/// Resolved data URLs for workbook visual media (floating pictures, in-cell
/// pictures, shape image fills), keyed per workbook session + visual.
///
/// Every viewport change re-installs the whole float layer (see
/// queueVisualInstall), which remounts these components. Without a cache each
/// remount dropped back to the "图片加载中…" placeholder and read the image out
/// of the archive through the sidecar again — scrolling a picture-heavy sheet
/// kept the visible pictures stuck on the loading placeholder. Media bytes are
/// immutable per (session, visual), so the resolved URL is kept and concurrent
/// mounts share one in-flight read.
const CACHE_LIMIT = 512

const resolved = new Map<string, string>()
const inflight = new Map<string, Promise<string | null>>()

function cacheKey(sessionId: string, visualId: string): string {
  return `${sessionId}\u0000${visualId}`
}

function remember(key: string, url: string): void {
  // Map preserves insertion order: re-insert on write so the oldest entry is
  // the first evicted once the safety cap is reached.
  resolved.delete(key)
  resolved.set(key, url)
  while (resolved.size > CACHE_LIMIT) {
    const oldest = resolved.keys().next().value
    if (oldest === undefined) break
    resolved.delete(oldest)
  }
}

/// Synchronous snapshot for a mount: null until this renderer resolved the
/// visual's media (callers then load it through loadVisualMediaUrl).
export function cachedVisualMediaUrl(
  sessionId: string | undefined,
  visualId: string,
): string | null {
  if (sessionId === undefined) return null
  const key = cacheKey(sessionId, visualId)
  const hit = resolved.get(key)
  if (hit === undefined) return null
  remember(key, hit)
  return hit
}

/// Reads a visual's media once per (session, visual). Concurrent callers share
/// the in-flight read; failures are not cached so a later mount can retry.
export function loadVisualMediaUrl(
  sessionId: string,
  visualId: string,
  read: (request: WorkbookMediaRequest) => Promise<string | null>,
): Promise<string | null> {
  const key = cacheKey(sessionId, visualId)
  const hit = resolved.get(key)
  if (hit !== undefined) {
    remember(key, hit)
    return Promise.resolve(hit)
  }
  const pending = inflight.get(key)
  if (pending) return pending
  const request = read({ sessionId, visualId })
    .then((url) => {
      if (url) remember(key, url)
      return url
    })
    .finally(() => {
      if (inflight.get(key) === request) inflight.delete(key)
    })
  inflight.set(key, request)
  return request
}

/// Drops every resolved URL of a closed workbook session (ids are never
/// reused, so keeping them would only leak image bytes). Reads still in flight
/// can re-add their own entry afterwards; the cap above bounds that.
export function forgetVisualMedia(sessionId: string): void {
  const prefix = `${sessionId}\u0000`
  for (const key of [...resolved.keys()]) {
    if (key.startsWith(prefix)) resolved.delete(key)
  }
  for (const key of [...inflight.keys()]) {
    if (key.startsWith(prefix)) inflight.delete(key)
  }
}

/// Test hook: drop every resolved/in-flight entry.
export function resetVisualMediaCache(): void {
  resolved.clear()
  inflight.clear()
}
