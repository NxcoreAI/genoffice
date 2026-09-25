/**
 * EverRoom Agent deck generation: assemble a multi-slide .pptx headlessly
 * from LLM-written page specs. Each page runs through the same tolerant
 * parser + local builder as the renderer's "local page generate" path, then
 * the one-slide results are merged with pptx-engine primitives. No renderer
 * view or slide session is involved — this is safe to call from an embed
 * host's main process.
 */
import { nativeImage } from 'electron'

import { fetchRemoteImage } from '@genoffice/electron-utils'
import {
  mergeSlideFromPptx,
  openPptx,
  promoteSlideBackground,
  savePptx,
} from '@genoffice/pptx-engine'

import { sniffImageMime } from './media-mime'
import { buildPagePptx, parsePageSpec } from './page-spec'
import { getFontMetrics } from './session-state'

export interface AgentDeckWarnings {
  page: number
  messages: string[]
}

export interface AgentDeckImageFailures {
  page: number
  url: string
}

export interface AgentDeckResult {
  bytes: Uint8Array
  warnings: AgentDeckWarnings[]
  imageFailures: AgentDeckImageFailures[]
}

const MAX_AGENT_PAGES = 24

export function agentPageDeps() {
  return {
    fontMetrics: getFontMetrics(),
    fetchImage: async (url: string): Promise<{ bytes: Uint8Array; ext: string } | null> => {
      const resp = await fetchRemoteImage(url)
      if (!resp || !resp.ok) return null
      const buf = new Uint8Array(await resp.arrayBuffer())
      const mime = sniffImageMime(buf) ?? resp.headers.get('content-type') ?? ''
      const ext = /png/.test(mime)
        ? 'png'
        : /gif/.test(mime)
          ? 'gif'
          : /webp/.test(mime)
            ? 'webp'
            : /bmp/.test(mime)
              ? 'bmp'
              : 'jpg'
      return { bytes: buf, ext }
    },
    imageDims: (bytes: Uint8Array): { width: number; height: number } | null => {
      try {
        const size = nativeImage.createFromBuffer(Buffer.from(bytes)).getSize()
        return size.width > 0 && size.height > 0 ? size : null
      } catch {
        return null
      }
    },
  }
}

/**
 * Build one deck from page spec JSON strings (page order = array order).
 * All pages come from the same single-slide builder, so the merge only moves
 * slide XML + media — layout/master/theme stay shared by construction.
 */
export async function buildAgentDeckPptx(
  pageSpecJsons: string[],
): Promise<{ ok: true; deck: AgentDeckResult } | { ok: false; error: string }> {
  if (!Array.isArray(pageSpecJsons) || pageSpecJsons.length === 0) {
    return { ok: false, error: 'at least one page spec is required' }
  }
  if (pageSpecJsons.length > MAX_AGENT_PAGES) {
    return { ok: false, error: `too many pages (${pageSpecJsons.length} > ${MAX_AGENT_PAGES})` }
  }

  const warnings: AgentDeckWarnings[] = []
  const imageFailures: AgentDeckImageFailures[] = []
  const pages: Uint8Array[] = []
  for (const [index, specJson] of pageSpecJsons.entries()) {
    const parsed = parsePageSpec(String(specJson ?? ''))
    if (!parsed.ok) {
      return { ok: false, error: `page ${index + 1}: ${parsed.error}` }
    }
    try {
      const built = await buildPagePptx(parsed.spec, agentPageDeps())
      pages.push(built.bytes)
      if (parsed.warnings.length > 0) {
        warnings.push({ page: index + 1, messages: parsed.warnings })
      }
      for (const url of built.imageFailures) {
        imageFailures.push({ page: index + 1, url })
      }
    } catch (err) {
      return {
        ok: false,
        error: `page ${index + 1}: ${err instanceof Error ? err.message : String(err)}`,
      }
    }
  }

  try {
    const base = await openPptx(pages[0]!)
    for (const one of pages.slice(1)) await mergeSlideFromPptx(base, one)
    for (const slide of base.deck.slides) promoteSlideBackground(slide, base.deck.size)
    return { ok: true, deck: { bytes: await savePptx(base), warnings, imageFailures } }
  } catch (err) {
    return { ok: false, error: `merge failed: ${err instanceof Error ? err.message : String(err)}` }
  }
}
