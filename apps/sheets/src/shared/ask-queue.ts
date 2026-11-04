/**
 * Pure validators for the EverRoom host annotation queue (sheets side).
 *
 * Everything that crosses the IPC boundary is untrusted: the renderer payload
 * is bounded here before the main process forwards it, the host's queue
 * snapshot is bounded before it reaches React, and the host's queue actions are
 * bounded before they are handed to the host hook. Kept zod-free so the preload
 * bundle stays lean; the main process and the renderer share the same rules.
 */
import type {
  SheetsAgentAskMode,
  SheetsAgentAskOp,
  SheetsAskQueueAction,
  SheetsAskQueueItem,
} from './desktop-api'

/** Cap for one note / instruction, matching the docs annotation contract. */
export const MAX_ASK_INSTRUCTION_CHARS = 2000
/** Cap for the selected-content snapshot carried with an ask. */
export const MAX_ASK_EXCERPT_CHARS = 200
/** Cap for one queue snapshot pushed down by the host (defensive bound). */
export const MAX_ASK_QUEUE_ITEMS = 100
const MAX_ID_CHARS = 80
const MAX_SHEET_ID_CHARS = 64
const MAX_SHEET_NAME_CHARS = 120
const MAX_RANGE_CHARS = 64

/** Collapse whitespace and cap at the excerpt limit: the host shows it verbatim. */
export function collapseExcerpt(text: string, max = MAX_ASK_EXCERPT_CHARS): string {
  return text.replace(/\s+/g, ' ').trim().slice(0, max)
}

/**
 * Build the excerpt from the selected cells' displayed text. Only the top-left
 * corner of a large selection is read — the payload cap is 200 characters, so
 * reading more would waste a facade round-trip.
 */
export function excerptFromMatrix(
  matrix: readonly (readonly (string | number | boolean | null | undefined)[])[],
  max = MAX_ASK_EXCERPT_CHARS,
): string {
  const parts: string[] = []
  let length = 0
  for (const row of matrix) {
    for (const value of row) {
      const text = value === null || value === undefined ? '' : String(value)
      if (!text) continue
      parts.push(text)
      length += text.length + 1
      if (length > max) return collapseExcerpt(parts.join(' '), max)
    }
  }
  return collapseExcerpt(parts.join(' '), max)
}

function trimmedString(value: unknown, max: number): string {
  return typeof value === 'string' ? value.trim().slice(0, max) : ''
}

function askMode(value: unknown): SheetsAgentAskMode | undefined {
  return value === 'now' || value === 'queue' ? value : undefined
}

/**
 * Normalize an untrusted renderer payload into a bounded ask op. Returns null
 * when the instruction is missing (the main process turns that into an
 * EMPTY_INSTRUCTION reply instead of calling the host).
 */
export function normalizeSheetsAgentAskOp(raw: unknown): SheetsAgentAskOp | null {
  if (!raw || typeof raw !== 'object') return null
  const value = raw as Record<string, unknown>
  const instruction = trimmedString(value.instruction, MAX_ASK_INSTRUCTION_CHARS)
  if (!instruction) return null
  const sheetId = trimmedString(value.sheetId, MAX_SHEET_ID_CHARS)
  const sheetName = trimmedString(value.sheetName, MAX_SHEET_NAME_CHARS)
  const range = trimmedString(value.range, MAX_RANGE_CHARS)
  const excerpt = typeof value.excerpt === 'string' ? collapseExcerpt(value.excerpt) : ''
  const mode = askMode(value.mode)
  return {
    instruction,
    ...(sheetId ? { sheetId } : {}),
    ...(sheetName ? { sheetName } : {}),
    ...(range ? { range } : {}),
    ...(excerpt ? { excerpt } : {}),
    ...(mode ? { mode } : {}),
  }
}

/**
 * Normalize a host-pushed queue snapshot. Invalid rows are dropped rather than
 * failing the whole snapshot: one malformed note must not blank the editor's
 * anchors while the user is mid-queue.
 */
export function normalizeSheetsAskQueueItems(raw: unknown): SheetsAskQueueItem[] {
  if (!Array.isArray(raw)) return []
  const items: SheetsAskQueueItem[] = []
  for (const entry of raw.slice(0, MAX_ASK_QUEUE_ITEMS)) {
    if (!entry || typeof entry !== 'object') continue
    const value = entry as Record<string, unknown>
    const id = trimmedString(value.id, MAX_ID_CHARS)
    const instruction = trimmedString(value.instruction, MAX_ASK_INSTRUCTION_CHARS)
    const sheetId = trimmedString(value.sheetId, MAX_SHEET_ID_CHARS)
    const range = trimmedString(value.range, MAX_RANGE_CHARS)
    // A note without an instruction is meaningless; a missing locator only
    // means the grid cannot draw it (the host card still shows the text).
    if (!id || !instruction) continue
    const excerpt = typeof value.excerpt === 'string' ? collapseExcerpt(value.excerpt) : ''
    const sheetName = trimmedString(value.sheetName, MAX_SHEET_NAME_CHARS)
    items.push({
      id,
      instruction,
      ...(sheetId ? { sheetId } : {}),
      ...(range ? { range } : {}),
      ...(excerpt ? { excerpt } : {}),
      ...(sheetName ? { sheetName } : {}),
    })
  }
  return items
}

/** Normalize a renderer-sent queue action (update / remove / send-all). */
export function normalizeSheetsAskQueueAction(raw: unknown): SheetsAskQueueAction | null {
  if (!raw || typeof raw !== 'object') return null
  const value = raw as Record<string, unknown>
  if (value.kind === 'send-all') return { kind: 'send-all' }
  const draftId = trimmedString(value.draftId, MAX_ID_CHARS)
  if (!draftId) return null
  if (value.kind === 'remove') return { kind: 'remove', draftId }
  if (value.kind === 'update') {
    const instruction =
      typeof value.instruction === 'string'
        ? value.instruction.replace(/\s+/g, ' ').trim().slice(0, MAX_ASK_INSTRUCTION_CHARS)
        : ''
    if (!instruction) return null
    return { kind: 'update', draftId, instruction }
  }
  return null
}
