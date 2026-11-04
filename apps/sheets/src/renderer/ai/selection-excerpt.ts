/**
 * Selected-cell text snapshot for the host annotation queue.
 *
 * The host shows this excerpt on the conversation-side note card and forwards
 * it to the agent as "what the user had selected", so it must stay short and
 * readable: displayed values (not raw formulas), whitespace collapsed, capped
 * at 200 characters. A large selection is read from its top-left corner only —
 * the cap makes anything beyond that useless and the facade round-trip costly.
 */
import { parseRange } from '../../domain/cell-address'
import { excerptFromMatrix } from '../../shared/ask-queue'
import type { UniverRuntime } from '../univer-state'

/** Cell block read before the excerpt cap is applied (bounded read cost). */
const EXCERPT_MAX_ROWS = 40
const EXCERPT_MAX_COLUMNS = 20

export function readSelectionExcerpt(runtime: UniverRuntime | null, rangeA1: string): string {
  if (!runtime || !rangeA1) return ''
  try {
    const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!worksheet) return ''
    const bounds = parseRange(rangeA1.replace(/\$/g, ''))
    const rows = Math.min(bounds.endRow - bounds.startRow + 1, EXCERPT_MAX_ROWS)
    const columns = Math.min(bounds.endColumn - bounds.startColumn + 1, EXCERPT_MAX_COLUMNS)
    if (rows <= 0 || columns <= 0) return ''
    const values = worksheet
      .getRange(bounds.startRow, bounds.startColumn, rows, columns)
      .getDisplayValues()
    return excerptFromMatrix(values)
  } catch {
    // A selection that no longer exists (sheet switched mid-send) has no excerpt.
    return ''
  }
}
