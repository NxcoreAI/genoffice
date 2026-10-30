import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  cachedVisualMediaUrl,
  forgetVisualMedia,
  loadVisualMediaUrl,
  resetVisualMediaCache,
} from '../src/renderer/visual-media-cache'

beforeEach(() => {
  resetVisualMediaCache()
})

describe('visual media cache', () => {
  it('resolves a visual once and serves re-mounts from the cache', async () => {
    const read = vi.fn(async () => 'data:image/png;base64,AAAA')
    await expect(loadVisualMediaUrl('session-1', 'visual-1', read)).resolves.toBe(
      'data:image/png;base64,AAAA',
    )
    expect(read).toHaveBeenCalledTimes(1)

    // Re-install (scroll) remounts the component: no second sidecar read.
    await expect(loadVisualMediaUrl('session-1', 'visual-1', read)).resolves.toBe(
      'data:image/png;base64,AAAA',
    )
    expect(read).toHaveBeenCalledTimes(1)
    expect(cachedVisualMediaUrl('session-1', 'visual-1')).toBe('data:image/png;base64,AAAA')
  })

  it('shares one in-flight read between concurrent mounts', async () => {
    let resolveRead: (url: string | null) => void = () => undefined
    const read = vi.fn(
      () =>
        new Promise<string | null>((resolve) => {
          resolveRead = resolve
        }),
    )
    const first = loadVisualMediaUrl('session-1', 'visual-1', read)
    const second = loadVisualMediaUrl('session-1', 'visual-1', read)
    expect(read).toHaveBeenCalledTimes(1)
    resolveRead('data:image/png;base64,BBBB')
    await expect(first).resolves.toBe('data:image/png;base64,BBBB')
    await expect(second).resolves.toBe('data:image/png;base64,BBBB')
  })

  it('keeps sessions and visuals apart', async () => {
    const read = (url: string) => async () => url
    await loadVisualMediaUrl('session-1', 'visual-1', read('one'))
    await loadVisualMediaUrl('session-1', 'visual-2', read('two'))
    await loadVisualMediaUrl('session-2', 'visual-1', read('three'))
    expect(cachedVisualMediaUrl('session-1', 'visual-1')).toBe('one')
    expect(cachedVisualMediaUrl('session-1', 'visual-2')).toBe('two')
    expect(cachedVisualMediaUrl('session-2', 'visual-1')).toBe('three')
    expect(cachedVisualMediaUrl(undefined, 'visual-1')).toBeNull()
  })

  it('does not cache failures so a later mount can retry', async () => {
    const read = vi.fn(async () => {
      throw new Error('sidecar unavailable')
    })
    await expect(loadVisualMediaUrl('session-1', 'visual-1', read)).rejects.toThrow(
      'sidecar unavailable',
    )
    expect(cachedVisualMediaUrl('session-1', 'visual-1')).toBeNull()
    await expect(loadVisualMediaUrl('session-1', 'visual-1', read)).rejects.toThrow(
      'sidecar unavailable',
    )
    expect(read).toHaveBeenCalledTimes(2)
  })

  it('drops a closed session without touching other sessions', async () => {
    await loadVisualMediaUrl('session-1', 'visual-1', async () => 'one')
    await loadVisualMediaUrl('session-2', 'visual-1', async () => 'two')
    forgetVisualMedia('session-1')
    expect(cachedVisualMediaUrl('session-1', 'visual-1')).toBeNull()
    expect(cachedVisualMediaUrl('session-2', 'visual-1')).toBe('two')
  })
})
