/**
 * EverRoom embed annotation channel: bound every untrusted payload before it
 * reaches the host agent or the renderer.
 *
 * Both directions carry user-typed text over IPC (renderer -> host ask, host ->
 * renderer queue snapshot), so each field is trimmed, whitespace-collapsed and
 * capped here, mirroring the docs implementation. Kept free of Electron imports
 * so the normalization is unit-testable.
 */
import type { SlidesAgentAskOp, SlidesAskQueueAction, SlidesAskQueueItem } from '../shared/ipc'

/** Hard cap on one instruction; longer requests belong in the main composer. */
export const ASK_INSTRUCTION_MAX = 2000
/** Excerpt cap (the UI shows a line or two; the agent uses it as a cross-check). */
export const ASK_EXCERPT_MAX = 200
/** Hard cap on a queue snapshot: one batch the editor can render at once. */
export const ASK_QUEUE_MAX = 50
/** Draft ids are host-generated; this only bounds pathological input. */
export const ASK_ID_MAX = 80
const ASK_TARGET_ID_MAX = 200
const ASK_TARGETS_MAX = 20

/** Collapse runs of whitespace and trim — the form both excerpt and notes take. */
function collapse(text: string): string {
  return text.replace(/\s+/g, ' ').trim()
}

function clip(text: string, max: number): string {
  return text.length > max ? text.slice(0, max) : text
}

function normalizeExcerpt(raw: unknown): string {
  return typeof raw === 'string' ? clip(collapse(raw), ASK_EXCERPT_MAX) : ''
}

/** Non-negative integer, or null when the input is not a usable number. */
function normalizeIndex(raw: unknown): number | null {
  if (typeof raw !== 'number' || !Number.isFinite(raw) || raw < 0) return null
  return Math.trunc(raw)
}

/**
 * Normalize an untrusted renderer ask payload. Null = unusable (empty
 * instruction, no valid targets, or a bad page index), which the caller
 * answers with an error instead of invoking the host hook.
 */
export function normalizeAgentAskOp(raw: unknown): SlidesAgentAskOp | null {
  const value = (raw ?? {}) as Partial<SlidesAgentAskOp>
  const instruction = typeof value.instruction === 'string' ? collapse(value.instruction) : ''
  const slideIndex = normalizeIndex(value.slideIndex)
  const rawTargets = Array.isArray(value.targets) ? value.targets : []
  const targets: SlidesAgentAskOp['targets'] = []
  for (const rawTarget of rawTargets) {
    if (!rawTarget || typeof rawTarget !== 'object') continue
    const id = typeof rawTarget.id === 'string' ? rawTarget.id.trim() : ''
    if (!id) continue
    const desc = (rawTarget.desc ?? {}) as {
      type?: unknown
      text?: unknown
      rows?: unknown
      cols?: unknown
    }
    const rows = normalizeIndex(desc.rows)
    const cols = normalizeIndex(desc.cols)
    const text = typeof desc.text === 'string' ? clip(collapse(desc.text), ASK_EXCERPT_MAX) : ''
    targets.push({
      id: clip(id, ASK_TARGET_ID_MAX),
      desc: {
        type: typeof desc.type === 'string' ? clip(desc.type, 40) : 'shape',
        ...(text ? { text } : {}),
        ...(rows !== null ? { rows } : {}),
        ...(cols !== null ? { cols } : {}),
      },
    })
    if (targets.length >= ASK_TARGETS_MAX) break
  }
  if (!instruction || slideIndex === null || targets.length === 0) return null
  const excerpt = normalizeExcerpt(value.excerpt)
  return {
    instruction: clip(instruction, ASK_INSTRUCTION_MAX),
    slideIndex,
    targets,
    ...(excerpt ? { excerpt } : {}),
    ...(value.mode === 'queue' ? { mode: 'queue' as const } : {}),
  }
}

/** Normalize one host queue row; null = unusable (missing id/note/anchor). */
export function normalizeAskQueueItem(raw: unknown): SlidesAskQueueItem | null {
  const value = (raw ?? {}) as Partial<SlidesAskQueueItem>
  const id = typeof value.id === 'string' ? clip(value.id.trim(), ASK_ID_MAX) : ''
  const instruction =
    typeof value.instruction === 'string'
      ? clip(collapse(value.instruction), ASK_INSTRUCTION_MAX)
      : ''
  const elementIds = Array.isArray(value.elementIds)
    ? [
        ...new Set(
          value.elementIds
            .filter((elementId): elementId is string => typeof elementId === 'string')
            .map((elementId) => clip(elementId.trim(), ASK_TARGET_ID_MAX))
            .filter(Boolean),
        ),
      ]
    : []
  if (!id || !instruction || elementIds.length === 0) return null
  const excerpt = normalizeExcerpt(value.excerpt)
  return {
    id,
    instruction,
    slideIndex: normalizeIndex(value.slideIndex) ?? 0,
    elementIds,
    ...(excerpt ? { excerpt } : {}),
  }
}

/** Normalize a whole snapshot: unusable rows are dropped, duplicate ids keep the first. */
export function normalizeAskQueueItems(raw: unknown): SlidesAskQueueItem[] {
  if (!Array.isArray(raw)) return []
  const items: SlidesAskQueueItem[] = []
  const seen = new Set<string>()
  for (const row of raw) {
    const item = normalizeAskQueueItem(row)
    if (!item || seen.has(item.id)) continue
    seen.add(item.id)
    items.push(item)
    if (items.length >= ASK_QUEUE_MAX) break
  }
  return items
}

/** Normalize an ask-queue action; null = unknown kind or missing fields. */
export function normalizeAskQueueAction(raw: unknown): SlidesAskQueueAction | null {
  const value = (raw ?? {}) as { kind?: unknown; draftId?: unknown; instruction?: unknown }
  if (value.kind === 'send-all') return { kind: 'send-all' }
  const draftId = typeof value.draftId === 'string' ? clip(value.draftId.trim(), ASK_ID_MAX) : ''
  if (value.kind === 'remove') return draftId ? { kind: 'remove', draftId } : null
  if (value.kind === 'update') {
    const instruction =
      typeof value.instruction === 'string'
        ? clip(collapse(value.instruction), ASK_INSTRUCTION_MAX)
        : ''
    return draftId && instruction ? { kind: 'update', draftId, instruction } : null
  }
  return null
}
