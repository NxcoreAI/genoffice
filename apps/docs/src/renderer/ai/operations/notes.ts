import { flushSync } from 'react-dom'
import type { NoteInfo } from '@genoffice/docx-engine'
import { deleteNote, submitNote } from '../../review-actions'
import { checkBlockIndex, failure, placeCaretAtBlockEnd, placeCaretAtDocEnd, planned, success } from './shared'
import type { OperationContext, OperationDef, OperationOutcome } from './types'

/**
 * B2 note ops: insert/update/delete footnotes and endnotes. Each op drives the
 * real review-actions pipeline (`submitNote` / `deleteNote`, the same code the
 * References tab and the note dialog use): the App-state note list and the
 * in-text reference markers stay in sync and the silent save persists them.
 *
 * `flushSync` around the actions makes the App re-render before the next op in
 * the batch reads `app.review()`, so two note ops in one batch never allocate
 * the same id / overwrite each other's list update.
 */

function noteKind(input: Record<string, unknown>): 'footnote' | 'endnote' | { error: string } {
  const kind = String(input.kind ?? '')
  if (kind !== 'footnote' && kind !== 'endnote') {
    return { error: 'kind must be "footnote" or "endnote"' }
  }
  return kind
}

function noteText(input: Record<string, unknown>): string | { error: string } {
  const text = typeof input.text === 'string' ? input.text.trim() : ''
  if (!text) return { error: 'text must be a non-empty string' }
  return text
}

function noteList(
  ctx: OperationContext,
  kind: 'footnote' | 'endnote',
): NoteInfo[] | null {
  const app = ctx.app
  if (!app) return null
  const review = app.review()
  return kind === 'footnote' ? review.footnotes : review.endnotes
}

/** ids of the in-text reference markers of one kind (live ProseMirror doc). */
function markerIds(ctx: OperationContext, kind: 'footnote' | 'endnote'): string[] {
  const ids: string[] = []
  ctx.editor.state.doc.descendants((node) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return
    ids.push(String(node.attrs.id))
  })
  return ids
}

function notesInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const kind = noteKind(input)
  if (typeof kind === 'object') return failure(kind.error)
  const text = noteText(input)
  if (typeof text === 'object') return failure(text.error)
  const app = ctx.app
  if (!app) return failure('notes.insert needs the app context (host call path)')
  const review = app.review()
  const list = kind === 'footnote' ? review.footnotes : review.endnotes
  let blockIndex: number | undefined
  if (input.blockIndex !== undefined) {
    const parsed = checkBlockIndex(ctx.editor, input.blockIndex)
    if (typeof parsed === 'string') return failure(parsed)
    blockIndex = parsed
  }
  const where = blockIndex === undefined ? 'at the caret' : `at the end of block ${blockIndex}`
  const preview = text.length > 60 ? `${text.slice(0, 57)}…` : text
  if (ctx.dryRun) {
    return planned(`would insert a ${kind} "${preview}" ${where} (note id ${list.length + 1})`, {
      kind,
      ...(blockIndex !== undefined ? { blockIndex } : {}),
    })
  }
  if (blockIndex !== undefined && !placeCaretAtBlockEnd(ctx.editor, blockIndex)) {
    return failure(`blockIndex out of range (document has ${ctx.editor.state.doc.childCount} blocks)`)
  }
  if (blockIndex === undefined && !ctx.editor.state.selection.$from.parent.inlineContent) {
    // no usable caret (e.g. a node/table selection): anchor at the document end
    placeCaretAtDocEnd(ctx.editor)
  }
  const before = markerIds(ctx, kind)
  flushSync(() => {
    // same semantics as the note dialog's submit (new id = max numeric id + 1,
    // marker num = existing count + 1) via the real review action
    submitNote({ ...review, notePrompt: { kind } }, text)
  })
  const after = markerIds(ctx, kind)
  const inserted = after.find((id) => !before.includes(id))
  if (inserted === undefined) {
    return failure('note marker insertion failed: no valid paragraph caret for the reference')
  }
  return success(
    `Inserted ${kind} ${inserted} ${where}: "${preview}".`,
    `insert ${kind}`,
    { kind, id: inserted, ...(blockIndex !== undefined ? { blockIndex } : {}) },
  )
}

function notesUpdateOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const kind = noteKind(input)
  if (typeof kind === 'object') return failure(kind.error)
  const text = noteText(input)
  if (typeof text === 'object') return failure(text.error)
  const id = String(input.id ?? '').trim()
  if (!id) return failure('id must be the note id (see the document state / block list)')
  const app = ctx.app
  if (!app) return failure('notes.update needs the app context (host call path)')
  const list = noteList(ctx, kind)
  if (!list) return failure('notes.update needs the app context (host call path)')
  if (!list.some((note) => note.id === id)) {
    return failure(`${kind} ${id} does not exist (${kind}s in the document: ${list.map((note) => note.id).join(', ') || 'none'})`)
  }
  if (ctx.dryRun) return planned(`would set ${kind} ${id} text to "${text.slice(0, 60)}"`, { kind, id })
  flushSync(() => {
    submitNote({ ...app.review(), notePrompt: { kind, id } }, text)
  })
  return success(`Updated ${kind} ${id}.`, `update ${kind}`, { kind, id })
}

function notesDeleteOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const kind = noteKind(input)
  if (typeof kind === 'object') return failure(kind.error)
  const id = String(input.id ?? '').trim()
  if (!id) return failure('id must be the note id (see the document state / block list)')
  const app = ctx.app
  if (!app) return failure('notes.delete needs the app context (host call path)')
  const list = noteList(ctx, kind)
  if (!list) return failure('notes.delete needs the app context (host call path)')
  if (!list.some((note) => note.id === id)) {
    return failure(`${kind} ${id} does not exist (${kind}s in the document: ${list.map((note) => note.id).join(', ') || 'none'})`)
  }
  if (ctx.dryRun) return planned(`would delete ${kind} ${id}`, { kind, id })
  flushSync(() => {
    // removes the note, its in-text marker and renumbers the remaining markers
    deleteNote(app.review(), kind, id)
  })
  return success(`Deleted ${kind} ${id}.`, `delete ${kind}`, { kind, id })
}

export const noteOperations: OperationDef[] = [
  {
    name: 'notes.insert',
    group: 'notes',
    description:
      'notes.insert { kind, text, blockIndex? } — insert a footnote/endnote; the in-text marker goes at the end of '
      + 'blockIndex (or at the caret); note id/number are assigned automatically',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['footnote', 'endnote'] },
        text: { type: 'string', minLength: 1, maxLength: 5000, description: 'note body text' },
        blockIndex: {
          type: 'integer',
          minimum: 0,
          description: 'block whose end receives the reference marker; omitted = the caret',
        },
      },
      required: ['kind', 'text'],
    },
    handler: notesInsertOp,
  },
  {
    name: 'notes.update',
    group: 'notes',
    description:
      'notes.update { kind, id, text } — replace the body text of an existing footnote/endnote '
      + '(ids come from the document state notes count / read_blocks marker list)',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['footnote', 'endnote'] },
        id: { type: 'string', minLength: 1, description: 'note id (e.g. "1")' },
        text: { type: 'string', minLength: 1, maxLength: 5000 },
      },
      required: ['kind', 'id', 'text'],
    },
    handler: notesUpdateOp,
  },
  {
    name: 'notes.delete',
    group: 'notes',
    description:
      'notes.delete { kind, id } — delete a footnote/endnote and its in-text marker; remaining markers are renumbered',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: ['footnote', 'endnote'] },
        id: { type: 'string', minLength: 1, description: 'note id (e.g. "1")' },
      },
      required: ['kind', 'id'],
    },
    handler: notesDeleteOp,
  },
]