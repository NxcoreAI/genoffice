/**
 * EverRoom embed: payload guard rails for the host annotation channels
 * (`pdf:agent-ask` / `pdf:ask-queue` / `pdf:ask-queue-action`).
 *
 * Deliberately free of electron imports so the validation stays unit-testable:
 * the IPC layer only forwards payloads that already came through here, and an
 * invalid anchor is dropped instead of reaching the host agent or the page
 * overlay.
 */
import type { PdfAgentAskOp, PdfAskQueueAction, PdfAskQueueItem } from '../shared/ipc'

/** Matches the host's own instruction bound (docs uses the same value). */
export const MAX_INSTRUCTION_CHARS = 2000
/** Selection excerpt snapshot bound — a cross-check for the agent, not the payload. */
export const MAX_EXCERPT_CHARS = 200
/** One selection can span many lines; cap what we accept so a hostile payload cannot flood IPC. */
export const MAX_QUADS = 200
/** Host queue snapshot size cap. */
export const MAX_QUEUE_ITEMS = 200
export const MAX_DRAFT_ID_CHARS = 80

/** PDF user-space quad: [x1,yMax,x2,yMax,x1,yMin,x2,yMin] (y up), as the native markups store it. */
const QUAD_NUMBERS = 8

function trimmedText(value: unknown, max: number): string {
  return typeof value === 'string' ? value.trim().slice(0, max) : ''
}

/** Collapse whitespace (excerpts are single paragraph snapshots in the UI). */
function excerptText(value: unknown): string {
  if (typeof value !== 'string') return ''
  return value.replace(/\s+/g, ' ').trim().slice(0, MAX_EXCERPT_CHARS)
}

/** 0-based page index; null when the payload does not carry a usable one. */
function pageIndex(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 0) return null
  return Math.trunc(value)
}

/** Keep only well-formed 8-number quads; null when nothing usable is left. */
function quads(value: unknown): number[][] | null {
  if (!Array.isArray(value)) return null
  const out: number[][] = []
  for (const raw of value) {
    if (out.length >= MAX_QUADS) break
    if (!Array.isArray(raw) || raw.length !== QUAD_NUMBERS) continue
    const quad: number[] = []
    let valid = true
    for (const n of raw) {
      if (typeof n !== 'number' || !Number.isFinite(n)) {
        valid = false
        break
      }
      quad.push(n)
    }
    if (valid) out.push(quad)
  }
  return out.length > 0 ? out : null
}

/**
 * Normalize an untrusted renderer payload into a bounded agent-ask op.
 * Returns null when the payload carries no instruction or no usable anchor.
 */
export function normalizeAgentAskOp(raw: unknown): PdfAgentAskOp | null {
  const value = (raw ?? {}) as Partial<PdfAgentAskOp>
  const instruction = trimmedText(value.instruction, MAX_INSTRUCTION_CHARS)
  if (!instruction) return null
  const page = pageIndex(value.pageIndex)
  const pageQuads = quads(value.quads)
  if (page === null || pageQuads === null) return null
  const excerpt = excerptText(value.excerpt)
  return {
    instruction,
    pageIndex: page,
    quads: pageQuads,
    ...(excerpt ? { excerpt } : {}),
    ...(value.mode === 'queue' ? { mode: 'queue' as const } : {}),
  }
}

/** Normalize a host queue snapshot; malformed rows are dropped, never trusted. */
export function normalizeAskQueueItems(raw: unknown): PdfAskQueueItem[] {
  if (!Array.isArray(raw)) return []
  const out: PdfAskQueueItem[] = []
  // draft ids key the anchors/cards; a duplicate would break React keys and
  // make hover/edit target the wrong row
  const seen = new Set<string>()
  for (const entry of raw) {
    if (out.length >= MAX_QUEUE_ITEMS) break
    if (!entry || typeof entry !== 'object') continue
    const value = entry as Partial<PdfAskQueueItem>
    const id = trimmedText(value.id, MAX_DRAFT_ID_CHARS)
    const instruction = trimmedText(value.instruction, MAX_INSTRUCTION_CHARS)
    const page = pageIndex(value.pageIndex)
    const pageQuads = quads(value.quads)
    if (!id || seen.has(id) || !instruction || page === null || pageQuads === null) continue
    seen.add(id)
    const excerpt = excerptText(value.excerpt)
    out.push({
      id,
      instruction,
      ...(excerpt ? { excerpt } : {}),
      pageIndex: page,
      quads: pageQuads,
    })
  }
  return out
}

/** Normalize one renderer queue action; null when the shape is not one of the three. */
export function normalizeAskQueueAction(raw: unknown): PdfAskQueueAction | null {
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
