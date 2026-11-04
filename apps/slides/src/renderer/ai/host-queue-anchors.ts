/**
 * Host annotation queue → element geometry.
 *
 * The EverRoom host owns the annotation queue and pushes flat snapshots down
 * (`SlidesAskQueueItem`): a note, a page hint, and the durable ids of the
 * elements it was made on. A stored page index breaks as soon as a page is
 * deleted or reordered, so it is kept only as a hint and every anchor is
 * re-resolved by searching the deck for its element ids — the same rule the
 * element edit queue uses (see edit-queue.ts).
 *
 * Pure helpers: no React, no DOM, so the resolution and excerpt rules can be
 * unit-tested with fabricated slides.
 */
import type { GroupRenderNode, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { SlidesAskQueueItem } from '../../shared/ipc'
import { anchorId } from './edit-queue'

/** Hard cap on the excerpt snapshot handed to the host agent. */
export const ASK_EXCERPT_MAX = 200

/** One element and its top-left corner in slide space (group offsets folded in). */
export interface ElementPlacement {
  node: RenderNode
  /** Slide-space x of the element's top-left (not group-local) */
  x: number
  /** Slide-space y of the element's top-left (not group-local) */
  y: number
}

/** Depth-first walk of a page's render tree, folding group offsets into x/y. */
export function flattenElements(nodes: RenderNode[], dx = 0, dy = 0): ElementPlacement[] {
  const out: ElementPlacement[] = []
  for (const node of nodes) {
    out.push({ node, x: dx + node.box.x, y: dy + node.box.y })
    if (node.type === 'group') {
      out.push(
        ...flattenElements((node as GroupRenderNode).children, dx + node.box.x, dy + node.box.y),
      )
    }
  }
  return out
}

/** Per-slide flatten cache: a queue snapshot re-resolves on every deck edit */
const placementCache = new WeakMap<RenderSlide, ElementPlacement[]>()

function placementsFor(slide: RenderSlide): ElementPlacement[] {
  const cached = placementCache.get(slide)
  if (cached) return cached
  const placements = flattenElements(slide.nodes)
  placementCache.set(slide, placements)
  return placements
}

/**
 * Locate every anchored element on one page. Durable ids are matched first;
 * parse-time source ids are the fallback for elements that lost their durable
 * bytes (same id discipline as the edit queue). Null = an element is missing,
 * which means the note no longer has a complete anchor and must not render.
 */
export function findElements(
  slide: RenderSlide,
  elementIds: readonly string[],
): ElementPlacement[] | null {
  const flat = placementsFor(slide)
  const found: ElementPlacement[] = []
  for (const elementId of elementIds) {
    const hit =
      flat.find((placement) => anchorId(placement.node) === elementId) ??
      flat.find((placement) => placement.node.sourceId === elementId)
    if (!hit) return null
    found.push(hit)
  }
  return found.length > 0 ? found : null
}

export interface ResolvedAskAnchor {
  item: SlidesAskQueueItem
  slideIndex: number
  placements: ElementPlacement[]
}

/**
 * Resolve one queued note against the current deck. The cached page is tried
 * first (cheap and exact in the common case), then the whole deck; a match on
 * several pages is ambiguous and dropped rather than guessed.
 */
export function resolveAskQueueItem(
  slides: readonly RenderSlide[],
  item: SlidesAskQueueItem,
): ResolvedAskAnchor | null {
  const elementIds = item.elementIds ?? []
  if (elementIds.length === 0) return null
  const hintIndex = typeof item.slideIndex === 'number' ? item.slideIndex : -1
  const hinted = hintIndex >= 0 ? slides[hintIndex] : undefined
  if (hinted) {
    const placements = findElements(hinted, elementIds)
    if (placements) return { item, slideIndex: hintIndex, placements }
  }
  const hits: ResolvedAskAnchor[] = []
  for (let index = 0; index < slides.length; index += 1) {
    if (index === hintIndex) continue
    const slide = slides[index]
    if (!slide) continue
    const placements = findElements(slide, elementIds)
    if (placements) hits.push({ item, slideIndex: index, placements })
  }
  return hits.length === 1 ? (hits[0] ?? null) : null
}

/** Resolve a whole snapshot, dropping notes whose elements are gone/ambiguous. */
export function resolveAskQueue(
  slides: readonly RenderSlide[],
  items: readonly SlidesAskQueueItem[],
): ResolvedAskAnchor[] {
  const resolved: ResolvedAskAnchor[] = []
  for (const item of items) {
    const anchor = resolveAskQueueItem(slides, item)
    if (anchor) resolved.push(anchor)
  }
  return resolved
}

/** Union bounding box of placements in slide space (for the outline + card anchor). */
export interface SlideBox {
  x: number
  y: number
  w: number
  h: number
}

export function unionBox(placements: readonly ElementPlacement[]): SlideBox | null {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const placement of placements) {
    minX = Math.min(minX, placement.x)
    minY = Math.min(minY, placement.y)
    maxX = Math.max(maxX, placement.x + placement.node.box.w)
    maxY = Math.max(maxY, placement.y + placement.node.box.h)
  }
  if (!Number.isFinite(minX) || !Number.isFinite(minY)) return null
  return { x: minX, y: minY, w: Math.max(0, maxX - minX), h: Math.max(0, maxY - minY) }
}

/**
 * Concatenate the selected elements' text into one bounded excerpt. The host
 * puts it next to the note as a cross-check (the agent re-reads the deck and
 * the locator may have drifted), so it is whitespace-collapsed and capped.
 */
export function buildElementExcerpt(
  descriptors: ReadonlyArray<{ text?: string }>,
  max = ASK_EXCERPT_MAX,
): string {
  const text = descriptors
    .map((descriptor) => descriptor.text ?? '')
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim()
  return text.length > max ? text.slice(0, max) : text
}
