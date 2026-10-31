import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Selection } from '@tiptap/pm/state'
import { tableOperations } from './table'
import { pageOperations } from './page'
import { sectionOperations } from './sections'
import { noteOperations } from './notes'
import { fieldOperations } from './fields'
import { formatOperations } from './format'
import { objectOperations } from './objects'
import { referenceOperations } from './references'
import { reviewOperations } from './review'
import type {
  AgentDocumentContext,
  AppliedOperation,
  OperationBatchResult,
  OperationContext,
  OperationDef,
  OperationOutcome,
} from './types'

export type {
  AgentDocumentContext,
  AgentProtectResult,
  AppliedOperation,
  OperationBatchResult,
  OperationContext,
  OperationDef,
  OperationHandler,
  OperationOutcome,
} from './types'
export { contentWidthPxFromSection, tableOperations } from './table'
export { pageOperations } from './page'
export { sectionOperations } from './sections'
export { noteOperations } from './notes'
export { fieldOperations } from './fields'
export { formatOperations } from './format'
export { objectOperations } from './objects'
export { referenceOperations } from './references'
export { reviewOperations } from './review'

/** Hard cap on ops per apply_operations call (matches the host's tool-call cap). */
export const MAX_OPERATIONS = 50

const OPERATIONS: OperationDef[] = [
  ...tableOperations,
  ...pageOperations,
  ...sectionOperations,
  ...noteOperations,
  ...fieldOperations,
  ...formatOperations,
  ...objectOperations,
  ...referenceOperations,
  ...reviewOperations,
]
const byName = new Map(OPERATIONS.map((def) => [def.name, def]))

/** Every registered op name, sorted for stable guidance/tests. */
export const OPERATION_NAMES: string[] = [...byName.keys()].sort()

/** Registry metadata (name/group/description/schema) — vocabulary for the host. */
export function describeOperations(): Array<{
  name: string
  group: string
  description: string
  inputSchema: Record<string, unknown>
}> {
  return OPERATIONS.map(({ name, group, description, inputSchema }) => ({
    name,
    group,
    description,
    inputSchema,
  }))
}

/**
 * Compact operations guide sent with the describe payload: shape of the call,
 * addressing/units, then the grouped op vocabulary. The per-op input schemas
 * live in the tool definition (describeOperations), so this stays token-light.
 */
function buildOperationsGuide(): string {
  const groups = new Map<string, OperationDef[]>()
  for (const def of OPERATIONS) {
    const list = groups.get(def.group) ?? []
    list.push(def)
    groups.set(def.group, list)
  }
  const lines = [
    'apply_operations runs a batch of structured document operations (max 50 per call).',
    'Input: { ops: [{ op: "table.insert", ...input }, ...], dryRun?: boolean }. Ops execute in order with per-op isolation: a failing op is reported with its array index and does not stop later ops; dryRun validates and describes the batch without changing the document.',
    '',
    'Addressing and units:',
    '- blockIndex = index of the target block in the document block list; table ops need a "table" block, page/notes/field ops any block.',
    '- Table cells are addressed by 0-based row/col grid coordinates as displayed; a merged cell covers its whole span and is addressed by its top-left row/col.',
    '- sections are 0-based document indexes (1 of N); note ids for notes.update/notes.delete are the footnoteIds/endnoteIds in the document state.',
    '- widths: CSS px (96dpi; 15 twips = 1px). heights/indents/margins: twips (1 cm = 567, 1 inch = 1440). colors: 6-digit hex without "#".',
    '- Before editing, read the block list / document state to confirm the table block, section count and note ids; prefer apply_commands for text/paragraph formatting. Page/section/notes/fields ops write App state and are persisted by the silent save after the batch.',
  ]
  for (const [group, defs] of groups) {
    lines.push('', `${group} operations:`)
    for (const def of defs) lines.push(`- ${def.description}`)
  }
  return lines.join('\n')
}

/** Static guide text for the describe payload (operationsGuide). */
export const OPERATIONS_GUIDE = buildOperationsGuide()

/** Tool description for the App host-agent dispatch / describe tools list. */
export const APPLY_OPERATIONS_DESCRIPTION =
  'Apply a batch of structured document operations (max 50) that apply_commands cannot express. '
  + 'Table structure/layout: insert/delete a table, insert/delete rows and columns, merge/split cells, '
  + 'set cell text/shading/borders/vAlign/width, set table alignment/indent/rowHeight/colWidths/borders/'
  + 'repeatHeaderRows/stylePreset, AutoFit. Page: page break, blank page, cover page (built-in presets), '
  + 'page-number format. Section: insert a section break, per-section page setup (orientation/paper/margins/'
  + 'columns/pageColor/themes). Notes: insert/update/delete footnotes and endnotes. Fields: insert an inline '
  + 'DATE/TIME/PAGE/NUMPAGES/FILENAME field, recompute all field caches (F9). '
  + 'Format: change case, clear direct formatting, LTR/RTL direction, hanging indent, paragraph + character '
  + 'styles (built-in emphasis/strong presets and document character style ids), ruler tab stops, hard line/'
  + 'column breaks, restart/continue numbering, multilevel lists and custom list styles. Objects: insert text '
  + 'box / WordArt / drop cap / equation '
  + '(LaTeX) / update an existing equation in place / symbol, format pictures (size/align/wrap/position/'
  + 'rotation/flip/z-order) and shapes '
  + '(fill/outline/text/size/z-order). References: bookmarks, cross-references, captions, bibliography '
  + 'sources/citations/bibliography, index entries + index field, find & replace. Review: tracked-changes '
  + 'toggle, accept/reject revisions (destructive; confirm:true), document protection and open-password '
  + 'encryption (destructive; confirm:true, passwords never echoed). Picture replace/crop/background-removal '
  + 'is NOT supported. '
  + 'Input: { ops: [{ op, ...input }], dryRun? }. Ops run in order; a failing op is reported by index and '
  + 'later ops still run. Blocks are addressed by block index, table cells by 0-based row/col, sections by '
  + 'section index. Read the "operations guide" in the document description for each op payload.'

/** JSON schema for the apply_operations tool input. */
export const APPLY_OPERATIONS_INPUT_SCHEMA: Record<string, unknown> = {
  type: 'object',
  properties: {
    ops: {
      type: 'array',
      description:
        'operations to apply in order (max 50). Each item is { op: "<name>", ...op-specific input }; see the operations guide',
      minItems: 1,
      maxItems: MAX_OPERATIONS,
      items: {
        type: 'object',
        properties: {
          op: { type: 'string', enum: OPERATION_NAMES },
        },
        required: ['op'],
      },
    },
    dryRun: {
      type: 'boolean',
      description: 'true = validate and describe the batch without changing the document',
    },
  },
  required: ['ops'],
}

export interface RunOperationsOptions {
  dryRun?: boolean
  /** Width of the current text column in px (needed by table/autoFit ops). */
  contentWidthPx?: number
  /**
   * App-state bundle for page/section/notes/fields ops (App's live
   * agentDocRefs.current). Ops that only touch ProseMirror ignore it; ops that
   * need it fail with a clear message when it is missing.
   */
  app?: AgentDocumentContext | null
}

function clip(text: string, max = 500): string {
  const flat = text.replace(/\s+/g, ' ').trim()
  return flat.length > max ? `${flat.slice(0, max - 1)}…` : flat
}

/**
 * Restore the document to `before` after a half-applied op (per-op isolation).
 * Content is replaced wholesale (the doc node is immutable), the restore step
 * is kept out of the undo history, and the cursor is moved to a valid position.
 */
function restoreDoc(editor: Editor, before: PmNode): void {
  const { state, view } = editor
  const tr = state.tr.replaceWith(0, state.doc.content.size, before.content)
  tr.setMeta('addToHistory', false)
  try {
    tr.setSelection(Selection.near(tr.doc.resolve(0), 1))
  } catch {
    // keep whatever selection the restore transaction mapped
  }
  view.dispatch(tr)
}

function formatLine(result: AppliedOperation): string {
  const label = result.op || '(missing op)'
  if (!result.ok) return `[${result.index}] ${label}: FAILED — ${clip(result.error ?? result.output, 300)}`
  if (result.planned) return `[${result.index}] ${label}: would apply — ${clip(result.output)}`
  return `[${result.index}] ${label}: ok — ${clip(result.output)}`
}

/**
 * Execute a batch of structured operations against the live editor.
 *
 * Per-op isolation: each op is planned/validated first, then applied; if a
 * handler fails or throws after touching the document, the document is
 * restored to its pre-op state and the failure is recorded per index. Later
 * ops still run, so one bad op never blocks the rest of the batch.
 */
export async function runOperations(
  editor: Editor,
  rawOps: unknown,
  options: RunOperationsOptions = {},
): Promise<OperationBatchResult> {
  const dryRun = options.dryRun === true
  const ctx: OperationContext = {
    editor,
    contentWidthPx: options.contentWidthPx ?? 600,
    dryRun,
    app: options.app ?? null,
  }
  const batchFailure = (message: string): OperationBatchResult => ({
    ok: false,
    dryRun,
    applied: 0,
    failed: 1,
    results: [{ index: 0, op: '', ok: false, output: message, error: message }],
    output: message,
  })
  if (!Array.isArray(rawOps)) return batchFailure('ops must be an array of operations')
  if (rawOps.length === 0) return batchFailure('ops is empty; pass at least one operation')
  const truncated = rawOps.length > MAX_OPERATIONS
  const ops = rawOps.slice(0, MAX_OPERATIONS)

  const results: AppliedOperation[] = []
  let applied = 0
  let failed = 0
  for (let index = 0; index < ops.length; index++) {
    const raw = ops[index]
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      failed++
      const failure = 'operation must be an object like { op: "table.insert", ... }'
      results.push({ index, op: '', ok: false, output: failure, error: failure })
      continue
    }
    const input = raw as Record<string, unknown>
    const name = typeof input.op === 'string' ? input.op : ''
    const def = byName.get(name)
    if (!def) {
      failed++
      const failure = name
        ? `unknown operation "${name}"; supported: ${OPERATION_NAMES.join(', ')}`
        : `missing op name; supported: ${OPERATION_NAMES.join(', ')}`
      results.push({ index, op: name, ok: false, output: failure, error: failure })
      continue
    }
    const beforeDoc = editor.state.doc
    let outcome: OperationOutcome
    try {
      // handlers may be async (B4 protect/encrypt await password hashing / IPC)
      outcome = await def.handler(input, ctx)
    } catch (error) {
      outcome = {
        ok: false,
        output: error instanceof Error ? error.message : String(error),
        error: error instanceof Error ? error.message : String(error),
      }
    }
    const mutated = editor.state.doc !== beforeDoc
    if (mutated && (!outcome.ok || dryRun)) {
      restoreDoc(editor, beforeDoc)
    } else if (mutated) {
      applied++
    }
    if (!outcome.ok) failed++
    results.push({
      index,
      op: def.name,
      ok: outcome.ok,
      output: clip(outcome.output),
      ...(outcome.summary ? { summary: outcome.summary } : {}),
      ...(outcome.error ? { error: clip(outcome.error, 300) } : {}),
      ...(outcome.record ? { record: outcome.record } : {}),
      ...(outcome.planned ? { planned: true } : {}),
    })
  }

  const okCount = results.filter((result) => result.ok).length
  const header = dryRun
    ? `Dry run: ${okCount}/${results.length} operation(s) validated, nothing changed.`
    : `Applied ${applied} of ${results.length} operation(s)${failed > 0 ? `, ${failed} failed` : ''}.`
  const lines = [header, ...results.map(formatLine)]
  if (truncated) {
    lines.push(`Only the first ${MAX_OPERATIONS} of ${rawOps.length} operations were processed.`)
  }
  return {
    ok: failed === 0,
    dryRun,
    applied,
    failed,
    results,
    output: lines.join('\n'),
  }
}