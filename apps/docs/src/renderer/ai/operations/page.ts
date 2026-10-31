import { flushSync } from 'react-dom'
import { insertBlankPageAt, insertPageBreakAt } from '../../components/ribbon-tabs'
import { COVER_PRESETS, buildCoverNodes, insertCoverPage } from '../../editor/cover-pages'
import {
  blockEndPos,
  failure,
  intOf,
  placeCaretAtBlockEnd,
  placeCaretAtDocEnd,
  planned,
  success,
} from './shared'
import type { OperationContext, OperationDef, OperationOutcome } from './types'

/**
 * B2 page ops: page breaks / blank pages / cover pages (pure ProseMirror, the
 * exact ribbon helpers) and page-number format (App pgNum state).
 */

const PG_NUM_FORMATS: Record<string, string | undefined> = {
  decimal: undefined,
  romanLower: 'lowerRoman',
  romanUpper: 'upperRoman',
  lettersLower: 'lowerLetter',
  lettersUpper: 'upperLetter',
  // canonical OOXML values accepted as aliases (what sections/pageNumberFmt uses)
  lowerRoman: 'lowerRoman',
  upperRoman: 'upperRoman',
  lowerLetter: 'lowerLetter',
  upperLetter: 'upperLetter',
}

/** afterBlockIndex semantics shared by the page ops: omitted / -1 = end of document. */
function resolveAfterBlock(
  ctx: OperationContext,
  input: Record<string, unknown>,
): number | { error: string } | null {
  const doc = ctx.editor.state.doc
  if (input.afterBlockIndex === undefined) return null
  const parsed = intOf(input.afterBlockIndex)
  if (parsed === null || parsed < -1 || parsed >= doc.childCount) {
    return { error: `afterBlockIndex must be -1 (end of document) or 0..${doc.childCount - 1}` }
  }
  return parsed >= 0 ? parsed : null
}

function pageBreakOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const at = resolveAfterBlock(ctx, input)
  if (at !== null && typeof at === 'object') return failure(at.error)
  const where = at === null ? 'at the end of the document' : `after block ${at}`
  if (ctx.dryRun) return planned(`would insert a page-break paragraph ${where}`)
  const before = ctx.editor.state.doc
  if (at === null) placeCaretAtDocEnd(ctx.editor)
  else if (!placeCaretAtBlockEnd(ctx.editor, at)) return failure(where + ' is out of range')
  insertPageBreakAt(ctx.editor)
  if (ctx.editor.state.doc === before) return failure('page break insert failed')
  return success(`Inserted a page-break paragraph ${where}.`, 'insert page break', {
    blockIndex: at === null ? before.childCount : at + 1,
  })
}

function pageBlankOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const at = resolveAfterBlock(ctx, input)
  if (at !== null && typeof at === 'object') return failure(at.error)
  const where = at === null ? 'at the end of the document' : `after block ${at}`
  if (ctx.dryRun) return planned(`would insert a blank page ${where}`)
  const before = ctx.editor.state.doc
  if (at === null) placeCaretAtDocEnd(ctx.editor)
  else if (!placeCaretAtBlockEnd(ctx.editor, at)) return failure(where + ' is out of range')
  insertBlankPageAt(ctx.editor)
  if (ctx.editor.state.doc === before) return failure('blank page insert failed')
  return success(
    `Inserted a blank page ${where} (following content now starts on the next page).`,
    'insert blank page',
    { blockIndex: at === null ? before.childCount : at + 1 },
  )
}

function pageCoverOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const presetId = input.preset === undefined ? COVER_PRESETS[0].id : String(input.preset)
  const preset =
    COVER_PRESETS.find((candidate) => candidate.id === presetId)
    ?? COVER_PRESETS.find((candidate) => candidate.name === presetId)
  if (!preset) {
    return failure(`preset must be one of ${COVER_PRESETS.map((candidate) => candidate.id).join('/')}`)
  }
  const doc = ctx.editor.state.doc
  let at: number | null = null
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    at = parsed
  }
  const where = at === null ? 'at the start of the document' : `after block ${at}`
  if (ctx.dryRun) return planned(`would insert the "${preset.id}" cover page ${where}`)
  const before = ctx.editor.state.doc
  if (at === null) {
    insertCoverPage(ctx.editor, preset)
  } else {
    ctx.editor.chain().insertContentAt(blockEndPos(doc, at), buildCoverNodes(preset)).run()
  }
  if (ctx.editor.state.doc === before) return failure('cover page insert failed')
  return success(
    `Inserted the "${preset.id}" cover page ${where}; the following content starts on a new page.`,
    'insert cover page',
    { preset: preset.id },
  )
}

function pageNumberFormatOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const raw = String(input.fmt ?? '').trim()
  if (raw === '' || !(raw in PG_NUM_FORMATS)) {
    return failure(`fmt must be one of ${Object.keys(PG_NUM_FORMATS).join('/')}`)
  }
  const fmt = PG_NUM_FORMATS[raw]
  let startAt: number | undefined
  if (input.startAt !== undefined && input.startAt !== null) {
    const parsed = intOf(input.startAt)
    if (parsed === null || parsed < 0 || parsed > 2_000_000) {
      return failure('startAt must be an integer between 0 and 2000000')
    }
    startAt = parsed
  }
  const app = ctx.app
  if (!app) return failure('page.numberFormat needs the app context (host call path)')
  const sections = app.sections
  if (sections.length === 0) return failure('no sections available (empty document?)')
  let index: number
  if (input.sectionIndex === undefined) {
    index = Math.min(Math.max(0, app.activeSection), sections.length - 1)
  } else {
    const parsed = intOf(input.sectionIndex)
    if (parsed === null || parsed < 0 || parsed >= sections.length) {
      return failure(`sectionIndex must be an integer 0..${sections.length - 1}`)
    }
    index = parsed
  }
  const where = `section ${index + 1} of ${sections.length}`
  const startText = startAt === undefined ? '' : ` starting at ${startAt}`
  if (ctx.dryRun) {
    return planned(`would set page numbering to "${raw}"${startText} on ${where}`, {
      sectionIndex: index,
    })
  }
  const next = { ...sections[index], pageNumberFmt: fmt, pageNumberStart: startAt }
  // same write path as the Page Number Format dialog (applyPgNumFormat): the
  // final section goes through SaveOptions.pgNumType, earlier sections rewrite
  // their sectPr on save; flushSync so the silent save right after sees it.
  flushSync(() => {
    app.setSections((prev) => prev.map((item, i) => (i === index ? next : item)))
    if (index >= sections.length - 1) {
      app.setPgNumEdit({
        ...(fmt !== undefined ? { fmt } : {}),
        ...(startAt !== undefined ? { start: startAt } : {}),
      })
    } else {
      app.setPgNumDirtySections((dirty) => (dirty.includes(index) ? dirty : [...dirty, index]))
    }
  })
  return success(
    `Set page numbering to "${raw}"${startText} on ${where}.`,
    'page number format',
    { sectionIndex: index, fmt: raw, ...(startAt !== undefined ? { startAt } : {}) },
  )
}

export const pageOperations: OperationDef[] = [
  {
    name: 'page.break',
    group: 'page',
    description:
      'page.break { afterBlockIndex? } — insert a page-break paragraph (the following content starts on a new page); '
      + 'afterBlockIndex -1/omitted = end of document',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          minimum: -1,
          description: 'insert after this block index; -1/omitted = end of document',
        },
      },
    },
    handler: pageBreakOp,
  },
  {
    name: 'page.blank',
    group: 'page',
    description:
      'page.blank { afterBlockIndex? } — insert a Word-style blank page (an empty page-breaking paragraph, '
      + 'and the following block is pushed to the next page); afterBlockIndex -1/omitted = end of document',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          minimum: -1,
          description: 'insert after this block index; -1/omitted = end of document',
        },
      },
    },
    handler: pageBlankOp,
  },
  {
    name: 'page.cover',
    group: 'page',
    description:
      'page.cover { preset?, afterBlockIndex? } — insert a built-in cover page (default at the start of the document); '
      + `preset ids: ${COVER_PRESETS.map((candidate) => candidate.id).join('/')}`,
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          minimum: 0,
          description: 'insert after this block index; omitted = start of the document',
        },
        preset: {
          type: 'string',
          enum: COVER_PRESETS.map((candidate) => candidate.id),
          description: 'cover design id; omitted = the first/classic preset',
        },
      },
    },
    handler: pageCoverOp,
  },
  {
    name: 'page.numberFormat',
    group: 'page',
    description:
      'page.numberFormat { fmt, sectionIndex?, startAt? } — set page-number format/start for one section '
      + '(fmt: decimal/romanLower/romanUpper/lettersLower/lettersUpper); sectionIndex omitted = the caret section',
    inputSchema: {
      type: 'object',
      properties: {
        sectionIndex: {
          type: 'integer',
          minimum: 0,
          description: 'target section (0-based); omitted = the section the caret is in',
        },
        fmt: {
          type: 'string',
          enum: ['decimal', 'romanLower', 'romanUpper', 'lettersLower', 'lettersUpper'],
          description: 'page-number format (decimal resets to the default numbering)',
        },
        startAt: {
          type: 'integer',
          minimum: 0,
          maximum: 2000000,
          description: 'first page number of the section; omitted = continue/clear',
        },
      },
      required: ['fmt'],
    },
    handler: pageNumberFormatOp,
  },
]