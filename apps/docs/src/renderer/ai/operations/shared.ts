import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import type { OperationOutcome } from './types'

/**
 * Shared helpers for the structured operation modules (table/page/section/
 * notes/fields): outcome constructors, scalar validation, block addressing and
 * caret placement. Keeping them here means every op family validates and
 * reports failures the same way, and per-op isolation in `runOperations` can
 * rely on a consistent `{ ok, output, error }` shape.
 */

export function failure(error: string): OperationOutcome {
  return { ok: false, output: error, error }
}

export function success(
  output: string,
  summary: string,
  record?: Record<string, unknown>,
): OperationOutcome {
  return { ok: true, output, summary, ...(record ? { record } : {}) }
}

export function planned(output: string, record?: Record<string, unknown>): OperationOutcome {
  return { ok: true, output, summary: 'planned', planned: true, ...(record ? { record } : {}) }
}

export function intOf(value: unknown): number | null {
  if (typeof value === 'number') return Number.isInteger(value) ? value : null
  if (typeof value === 'string' && value.trim() !== '') {
    const n = Number(value)
    return Number.isInteger(n) ? n : null
  }
  return null
}

export function clampInt(value: unknown, fallback: number, min: number, max: number): number {
  const n = intOf(value)
  if (n === null) return fallback
  return Math.min(max, Math.max(min, n))
}

/** 6-digit hex without '#', uppercased; null clears; undefined = not provided. */
export function hexOf(value: unknown): string | null | undefined {
  if (value === undefined) return undefined
  if (value === null) return null
  const hex = String(value).replace(/^#/, '').trim()
  return /^[0-9A-Fa-f]{6}$/.test(hex) ? hex.toUpperCase() : undefined
}

/** Doc position before the child at `index`. */
export function blockStartPos(doc: PmNode, index: number): number {
  let pos = 0
  for (let i = 0; i < index; i++) pos += doc.child(i).nodeSize
  return pos
}

/** Doc position right after the child at `index` (insertion point after that block). */
export function blockEndPos(doc: PmNode, index: number): number {
  return blockStartPos(doc, index) + doc.child(index).nodeSize
}

/**
 * Put a synthetic caret at the end of `blockIndex` (block-level helpers that
 * insert at the selection, e.g. page break / blank page / field insertion).
 * Near(-1) lands on the last text position inside the block, which tiptap's
 * insertContent then splits after the block. false when out of range.
 */
export function placeCaretAtBlockEnd(editor: Editor, blockIndex: number): boolean {
  const doc = editor.state.doc
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= doc.childCount) return false
  const selection = TextSelection.near(doc.resolve(blockEndPos(doc, blockIndex)), -1)
  editor.view.dispatch(editor.state.tr.setSelection(selection))
  return true
}

/** Put a synthetic caret at the end of the document (before the trailing doc end). */
export function placeCaretAtDocEnd(editor: Editor): void {
  const doc = editor.state.doc
  const selection = TextSelection.near(doc.resolve(doc.content.size), -1)
  editor.view.dispatch(editor.state.tr.setSelection(selection))
}

/** Count text nodes carrying a given mark (e.g. instrField for F9 fields). */
export function countMarkedText(editor: Editor, markName: string): number {
  let count = 0
  editor.state.doc.descendants((node) => {
    if (node.isText && node.marks.some((mark) => mark.type.name === markName)) count++
  })
  return count
}

/** Ensure a document block index is present (used by block-addressed ops). */
export function checkBlockIndex(editor: Editor, value: unknown): number | string {
  const doc = editor.state.doc
  const index = intOf(value)
  if (index === null || index < 0 || index >= doc.childCount) {
    return `blockIndex must be an integer 0..${doc.childCount - 1}`
  }
  return index
}
/** Block types that accept inline formatting (paragraph-like text blocks). */
export const TEXT_BLOCK_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

/** Target-matched blocks that are text blocks (image/protected blocks carry no inline text). */
export function textBlocksOf<T extends { node: { type: { name: string } } }>(blocks: T[]): T[] {
  return blocks.filter((block) => TEXT_BLOCK_TYPES.has(block.node.type.name))
}

/** Doc range covering the given target blocks (positions inside their text). */
export function blocksRange(
  blocks: Array<{ pos: number; node: { nodeSize: number } }>,
): { from: number; to: number } | null {
  if (blocks.length === 0) return null
  const first = blocks[0]
  const last = blocks[blocks.length - 1]
  const from = first.pos + 1
  const to = Math.max(from, last.pos + last.node.nodeSize - 1)
  return { from, to }
}

/**
 * Put a synthetic text selection across the target blocks so selection-driven
 * editor helpers (applyCase, setParagraphDirection, applyParagraphStyle,
 * clearParagraphFormatting, list commands) run on exactly the matched blocks.
 */
export function selectBlocks(
  editor: Editor,
  blocks: Array<{ pos: number; node: { nodeSize: number } }>,
): { from: number; to: number } | null {
  const range = blocksRange(blocks)
  if (!range) return null
  const doc = editor.state.doc
  const from = Math.min(range.from, doc.content.size)
  const to = Math.min(range.to, doc.content.size)
  const selection =
    from < to
      ? TextSelection.create(doc, from, to)
      : TextSelection.near(doc.resolve(from), 1)
  editor.view.dispatch(editor.state.tr.setSelection(selection))
  return { from, to }
}
