/**
 * EverRoom embed: payload guard rails for the host annotation channels
 * (`markdown:agent-ask` / `markdown:ask-queue` / `markdown:ask-queue-action`).
 *
 * Deliberately free of electron imports so the validation stays unit-testable:
 * the IPC layer only forwards payloads that already came through here, and an
 * invalid anchor is dropped instead of reaching the host agent or the editor
 * overlay (docs/pdf parity).
 */
import type { MdAgentAskOp, MdAskQueueAction, MdAskQueueItem } from '../shared/ipc'

/** Matches the host's own instruction bound (docs/pdf use the same value). */
export const MAX_INSTRUCTION_CHARS = 2000
/** Selection excerpt snapshot bound — a cross-check for the agent, not the payload. */
export const MAX_EXCERPT_CHARS = 200
/** Host queue snapshot size cap. */
export const MAX_QUEUE_ITEMS = 200
export const MAX_DRAFT_ID_CHARS = 80

function trimmedText(value: unknown, max: number): string {
  return typeof value === 'string' ? value.trim().slice(0, max) : ''
}

/** Collapse whitespace (excerpts are single paragraph snapshots in the UI). */
function excerptText(value: unknown): string {
  if (typeof value !== 'string') return ''
  return value.replace(/\s+/g, ' ').trim().slice(0, MAX_EXCERPT_CHARS)
}

/** Non-negative integer block index; null when the payload has no usable one. */
function blockIndex(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) return null
  return Math.trunc(value)
}

/** Non-negative character offset (optional field; null only when malformed). */
function charOffset(value: unknown): number | null {
  if (value === undefined || value === null) return 0
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) return null
  return Math.trunc(value)
}

/**
 * Normalize an untrusted renderer payload into a bounded agent-ask op.
 * Returns null when the payload carries no instruction or no usable anchor.
 */
export function normalizeAgentAskOp(raw: unknown): MdAgentAskOp | null {
  const value = (raw ?? {}) as Partial<MdAgentAskOp>
  const instruction = trimmedText(value.instruction, MAX_INSTRUCTION_CHARS)
  if (!instruction) return null
  const block = blockIndex(value.blockIndex)
  if (block === null) return null
  const offset = charOffset(value.offset)
  if (offset === null) return null
  const excerpt = excerptText(value.excerpt)
  return {
    instruction,
    blockIndex: block,
    ...(excerpt ? { excerpt } : {}),
    ...(offset > 0 ? { offset } : {}),
    ...(value.mode === 'queue' ? { mode: 'queue' as const } : {}),
    ...(value.scope === 'document' ? { scope: 'document' as const } : {}),
  }
}

/** Normalize a host queue snapshot; malformed rows are dropped, never trusted. */
export function normalizeAskQueueItems(raw: unknown): MdAskQueueItem[] {
  if (!Array.isArray(raw)) return []
  const out: MdAskQueueItem[] = []
  // draft ids key the anchors/cards; a duplicate would break React keys and
  // make click-edit target the wrong row
  const seen = new Set<string>()
  for (const entry of raw) {
    if (out.length >= MAX_QUEUE_ITEMS) break
    if (!entry || typeof entry !== 'object') continue
    const value = entry as Partial<MdAskQueueItem>
    const id = trimmedText(value.id, MAX_DRAFT_ID_CHARS)
    const instruction = trimmedText(value.instruction, MAX_INSTRUCTION_CHARS)
    const block = blockIndex(value.blockIndex)
    const offset = charOffset(value.offset)
    if (!id || seen.has(id) || !instruction || block === null || offset === null) continue
    seen.add(id)
    const excerpt = excerptText(value.excerpt)
    out.push({
      id,
      instruction,
      ...(excerpt ? { excerpt } : {}),
      blockIndex: block,
      ...(offset > 0 ? { offset } : {}),
    })
  }
  return out
}

/** Normalize one renderer queue action; null when the shape is not one of the three. */
export function normalizeAskQueueAction(raw: unknown): MdAskQueueAction | null {
  const value = (raw ?? {}) as { kind?: unknown; draftId?: unknown; instruction?: unknown }
  if (value.kind === 'send-all') return { kind: 'send-all' }
  const draftId = trimmedText(value.draftId, MAX_DRAFT_ID_CHARS)
  if (!draftId) return null
  if (value.kind === 'remove') return { kind: 'remove', draftId }
  if (value.kind === 'update') {
    const instruction = trimmedText(value.instruction, MAX_INSTRUCTION_CHARS)
    return instruction ? { kind: 'update', draftId, instruction } : null
  }
  return null
}
