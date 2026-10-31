import { TextSelection } from '@tiptap/pm/state'
import {
  blockEndPos,
  countMarkedText,
  failure,
  intOf,
  placeCaretAtBlockEnd,
  placeCaretAtDocEnd,
  planned,
  success,
} from './shared'
import type { OperationContext, OperationDef, OperationOutcome } from './types'

/**
 * B2 field ops: insert an inline Word field (DATE/TIME/PAGE/NUMPAGES/FILENAME)
 * through the App's real `insertField` helper (same cached-value + instrField
 * mark pipeline as Insert → Field), and recompute every field cache (F9) via
 * the real `updateFields`.
 */

const FIELD_KINDS = ['PAGE', 'NUMPAGES', 'DATE', 'TIME', 'FILENAME'] as const

function fieldInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const kindRaw = String(input.kind ?? '').trim().toUpperCase()
  if (!(FIELD_KINDS as readonly string[]).includes(kindRaw)) {
    return failure(`kind must be one of ${FIELD_KINDS.join('/')}`)
  }
  const app = ctx.app
  if (!app) return failure('field.insert needs the app context (host call path)')
  const doc = ctx.editor.state.doc
  const hasBlock = input.blockIndex !== undefined
  const hasAfter = input.afterBlockIndex !== undefined
  if (hasBlock && hasAfter) return failure('pass either blockIndex or afterBlockIndex, not both')
  let blockIndex: number | null = null
  if (hasBlock) {
    const parsed = intOf(input.blockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`blockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    blockIndex = parsed
  }
  let afterBlockIndex: number | null = null
  if (hasAfter) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    afterBlockIndex = parsed
  }
  const where =
    blockIndex !== null
      ? `at the end of block ${blockIndex}`
      : afterBlockIndex !== null
        ? `in a new paragraph after block ${afterBlockIndex}`
        : 'at the caret'
  if (ctx.dryRun) return planned(`would insert a ${kindRaw} field ${where}`, { kind: kindRaw })
  if (blockIndex !== null) {
    if (!placeCaretAtBlockEnd(ctx.editor, blockIndex)) return failure(`${where} is out of range`)
  } else if (afterBlockIndex !== null) {
    // own paragraph after the block, then the real helper writes into it
    const pos = blockEndPos(doc, afterBlockIndex)
    ctx.editor.chain().insertContentAt(pos, { type: 'docParagraph', attrs: { docxIndex: null } }).run()
    if (ctx.editor.state.doc.childCount <= doc.childCount) {
      return failure('could not insert the paragraph that should hold the field')
    }
    const selection = TextSelection.near(ctx.editor.state.doc.resolve(pos + 1), 1)
    ctx.editor.view.dispatch(ctx.editor.state.tr.setSelection(selection))
  } else if (!ctx.editor.state.selection.$from.parent.inlineContent) {
    // no usable caret (node/table selection): anchor at the document end
    placeCaretAtDocEnd(ctx.editor)
  }
  const before = countMarkedText(ctx.editor, 'instrField')
  app.insertField(kindRaw)
  const after = countMarkedText(ctx.editor, 'instrField')
  if (after <= before) return failure('field insertion failed: no inline field mark was added')
  return success(
    `Inserted a ${kindRaw} field ${where}.`,
    'insert field',
    {
      kind: kindRaw,
      ...(blockIndex !== null ? { blockIndex } : {}),
      ...(afterBlockIndex !== null ? { afterBlockIndex } : {}),
    },
  )
}

function fieldUpdateAllOp(_input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const app = ctx.app
  if (!app) return failure('field.updateAll needs the app context (host call path)')
  const count = countMarkedText(ctx.editor, 'instrField')
  if (ctx.dryRun) {
    return planned(
      count > 0 ? `would recompute ${count} inline field(s)` : 'there are no inline fields to update',
      { fields: count },
    )
  }
  // real F9 path: rewrites every PAGE/NUMPAGES/DATE/TIME/FILENAME cache in place
  app.updateFields()
  return success(
    count > 0 ? `Recomputed ${count} inline field(s).` : 'No inline fields to update.',
    'update fields',
    { fields: count },
  )
}

export const fieldOperations: OperationDef[] = [
  {
    name: 'field.insert',
    group: 'field',
    description:
      'field.insert { kind, blockIndex?, afterBlockIndex? } — insert an inline field with its current cached value '
      + `(kind: ${FIELD_KINDS.join('/')}); blockIndex appends at the end of that block, afterBlockIndex puts it in a new paragraph`,
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: FIELD_KINDS },
        blockIndex: {
          type: 'integer',
          minimum: 0,
          description: 'append the field at the end of this block',
        },
        afterBlockIndex: {
          type: 'integer',
          minimum: 0,
          description: 'insert the field in a new paragraph after this block',
        },
      },
      required: ['kind'],
    },
    handler: fieldInsertOp,
  },
  {
    name: 'field.updateAll',
    group: 'field',
    description:
      'field.updateAll {} — recompute every inline field cache (Word F9): page numbers, total pages, dates/times, file name',
    inputSchema: { type: 'object', properties: {} },
    handler: fieldUpdateAllOp,
  },
]