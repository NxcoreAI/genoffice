import {
  collectBookmarks,
  insertBibliography,
  insertBookmark,
  insertCaption,
  insertCitation,
  insertCrossRef,
  insertIndexField,
  makeSourceInfo,
  markIndexEntry,
  removeBookmark,
} from '../../editor/references'
import type { OperationContext, OperationDef, OperationOutcome } from './types'
import {
  blockEndPos,
  checkBlockIndex,
  failure,
  intOf,
  placeCaretAtBlockEnd,
  placeCaretAtDocEnd,
  planned,
  success,
} from './shared'

/**
 * B4 reference ops: bookmarks / cross-references / captions / bibliography
 * sources + citations + bibliography / index entries + index field. All reuse
 * the shared editor/references.ts helpers (the same code the References and
 * Insert ribbon tabs call), so the agent path and the UI cannot drift.
 *
 * Pure-ProseMirror ops work without the app context; ops that read/write the
 * bibliography source list (source.add, citation.insert, bibliography.insert)
 * require it and fail clearly when it is absent.
 */

const PARA_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

function requireApp(ctx: OperationContext): NonNullable<OperationContext['app']> | null {
  return ctx.app
}

/** Put the caret inside a text block (bookmark anchors / citation text targets). */
function placeCaretInBlock(ctx: OperationContext, blockIndex: number): boolean {
  return placeCaretAtBlockEnd(ctx.editor, blockIndex)
}

function ensureInlineCaret(ctx: OperationContext): void {
  if (!ctx.editor.state.selection.$from.parent.inlineContent) placeCaretAtDocEnd(ctx.editor)
}

function bookmarkInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const name = String(input.name ?? '').trim()
  if (!name) return failure('name must be a non-empty bookmark name')
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  if (!PARA_TYPES.has(node.type.name)) {
    return failure(`block ${blockIndex} is a ${node.type.name}; bookmarks can only anchor to paragraph/heading/list-item blocks`)
  }
  if (collectBookmarks(ctx.editor).some((entry) => entry.name === name)) {
    return failure(`bookmark "${name}" already exists`)
  }
  if (ctx.dryRun) return planned(`would add bookmark "${name}" to block ${blockIndex}`, { blockIndex, name })
  placeCaretInBlock(ctx, blockIndex)
  const result = insertBookmark(ctx.editor, name)
  if (!result.ok) return failure(result.error ?? 'bookmark insert failed')
  return success(`Added bookmark "${name}" to block ${blockIndex}.`, 'insert bookmark', { blockIndex, name })
}

function bookmarkRemoveOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const name = String(input.name ?? '').trim()
  if (!name) return failure('name must be a non-empty bookmark name')
  if (!collectBookmarks(ctx.editor).some((entry) => entry.name === name)) {
    return failure(`bookmark "${name}" does not exist`)
  }
  if (ctx.dryRun) return planned(`would remove bookmark "${name}"`, { name })
  if (!removeBookmark(ctx.editor, name)) return failure(`bookmark "${name}" could not be removed`)
  return success(`Removed bookmark "${name}".`, 'remove bookmark', { name })
}

function crossRefInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const bookmark = String(input.bookmark ?? '').trim()
  if (!bookmark) return failure('bookmark must be a non-empty bookmark name')
  if (!collectBookmarks(ctx.editor).some((entry) => entry.name === bookmark)) {
    const names = collectBookmarks(ctx.editor).map((entry) => entry.name)
    return failure(
      `bookmark "${bookmark}" does not exist`
        + (names.length > 0 ? `; available: ${names.join(', ')}` : '; add one with reference.bookmark.insert first'),
    )
  }
  const doc = ctx.editor.state.doc
  let afterBlockIndex: number | null = null
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    afterBlockIndex = parsed
  }
  const where = afterBlockIndex === null ? 'at the caret' : `at the end of block ${afterBlockIndex}`
  if (ctx.dryRun) return planned(`would insert a cross-reference to "${bookmark}" ${where}`, { bookmark })
  if (afterBlockIndex !== null) placeCaretInBlock(ctx, afterBlockIndex)
  else ensureInlineCaret(ctx)
  const result = insertCrossRef(ctx.editor, bookmark)
  if (!result.ok) return failure(result.error ?? 'cross-reference insert failed')
  return success(
    `Inserted a cross-reference to "${bookmark}" ${where}.`,
    'insert cross-reference',
    { bookmark, ...(afterBlockIndex !== null ? { afterBlockIndex } : {}) },
  )
}

function captionInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const app = requireApp(ctx)
  if (!app) return failure('reference.caption.insert needs the app context (document blocks)')
  const label = String(input.label ?? '').trim()
  if (!label) return failure('label must be a non-empty caption label (e.g. "Figure"/"表")')
  const text = input.text === undefined || input.text === null ? '' : String(input.text)
  if (text.length > 500) return failure('text must be at most 500 characters')
  const doc = ctx.editor.state.doc
  let blockIndex: number | null = null
  if (input.blockIndex !== undefined) {
    const parsed = intOf(input.blockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`blockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    blockIndex = parsed
  }
  const where = blockIndex === null ? 'at the caret' : `after block ${blockIndex}`
  if (ctx.dryRun) {
    return planned(`would insert a "${label}" caption ${where}` + (text ? ` with text "${text}"` : ''))
  }
  if (blockIndex !== null) placeCaretInBlock(ctx, blockIndex)
  else ensureInlineCaret(ctx)
  const result = insertCaption(ctx.editor, app.getBlocks(), label, text)
  if (!result.ok) return failure(result.error ?? 'caption insert failed')
  return success(
    `Inserted caption "${label} ${result.number}" ${where}.`,
    'insert caption',
    {
      label,
      number: result.number,
      ...(blockIndex !== null ? { afterBlockIndex: blockIndex } : {}),
    },
  )
}

function sourceAddOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const app = requireApp(ctx)
  if (!app) return failure('reference.source.add needs the app context (bibliography sources)')
  const existing = app.getSources()
  const result = makeSourceInfo(existing, {
    title: String(input.title ?? ''),
    author: input.author === undefined ? '' : String(input.author),
    year: input.year === undefined ? '' : String(input.year),
    publisher: input.publisher === undefined ? '' : String(input.publisher),
    url: input.url === undefined ? '' : String(input.url),
    type: input.type === undefined ? 'Book' : String(input.type),
  })
  if (!result.ok || !result.source) return failure(result.error ?? 'could not build the source')
  const source = result.source
  if (ctx.dryRun) {
    return planned(`would add bibliography source "${source.tag}" (${source.title})`, { tag: source.tag })
  }
  // App.addSource flushes the state so a citation/bibliography later in the
  // same batch sees the new source and the batch save persists sourcesDirty
  app.addSource(source)
  return success(
    `Added bibliography source "${source.tag}": ${source.author || source.title}${source.year ? ` (${source.year})` : ''}.`,
    'add source',
    { tag: source.tag },
  )
}

function citationInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const app = requireApp(ctx)
  if (!app) return failure('reference.citation.insert needs the app context (bibliography sources)')
  const sources = app.getSources()
  if (sources.length === 0) {
    return failure('the document has no bibliography sources; add one with reference.source.add')
  }
  const hasIndex = input.sourceIndex !== undefined
  const hasId = input.sourceId !== undefined
  if (hasIndex === hasId) return failure('pass exactly one of sourceIndex or sourceId')
  let source = null as (typeof sources)[number] | null
  if (hasIndex) {
    const index = intOf(input.sourceIndex)
    if (index === null || index < 0 || index >= sources.length) {
      return failure(`sourceIndex must be an integer 0..${sources.length - 1}`)
    }
    source = sources[index]
  } else {
    const tag = String(input.sourceId ?? '').trim()
    source = sources.find((candidate) => candidate.tag === tag) ?? null
    if (!source) {
      return failure(`source "${tag}" does not exist; available tags: ${sources.map((s) => s.tag).join(', ')}`)
    }
  }
  const doc = ctx.editor.state.doc
  let afterBlockIndex: number | null = null
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    afterBlockIndex = parsed
  }
  const where = afterBlockIndex === null ? 'at the caret' : `at the end of block ${afterBlockIndex}`
  if (ctx.dryRun) {
    return planned(`would insert the citation for "${source.tag}" ${where}`, { tag: source.tag })
  }
  if (afterBlockIndex !== null) placeCaretInBlock(ctx, afterBlockIndex)
  else ensureInlineCaret(ctx)
  if (!insertCitation(ctx.editor, source)) return failure('citation insert failed (no valid caret position)')
  return success(`Inserted the citation for "${source.tag}" ${where}.`, 'insert citation', {
    tag: source.tag,
  })
}

function bibliographyInsertOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const app = requireApp(ctx)
  if (!app) return failure('reference.bibliography.insert needs the app context (bibliography sources)')
  const sources = app.getSources()
  if (sources.length === 0) {
    return failure('the document has no bibliography sources; add one with reference.source.add')
  }
  const doc = ctx.editor.state.doc
  let afterBlockIndex: number | null = null
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    afterBlockIndex = parsed
  }
  const where = afterBlockIndex === null ? 'at the caret' : `after block ${afterBlockIndex}`
  if (ctx.dryRun) {
    return planned(`would insert a bibliography with ${sources.length} source line(s) ${where}`)
  }
  if (afterBlockIndex !== null) placeCaretInBlock(ctx, afterBlockIndex)
  else ensureInlineCaret(ctx)
  const result = insertBibliography(ctx.editor, sources)
  if (!result.ok) return failure(result.error ?? 'bibliography insert failed')
  return success(
    `Inserted a bibliography with ${sources.length} source line(s) ${where}.`,
    'insert bibliography',
    { lines: sources.length, ...(afterBlockIndex !== null ? { afterBlockIndex } : {}) },
  )
}

function indexMarkOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const term = String(input.term ?? '').trim()
  if (!term) return failure('term must be a non-empty index entry term')
  if (term.length > 200) return failure('term must be at most 200 characters')
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  if (!PARA_TYPES.has(node.type.name)) {
    return failure(`block ${blockIndex} is a ${node.type.name}; index entries (XE marks) attach to paragraph text`)
  }
  const at = blockEndPos(doc, blockIndex) - 1
  if (ctx.dryRun) return planned(`would mark "${term}" as an index entry in block ${blockIndex}`, { blockIndex, term })
  if (!markIndexEntry(ctx.editor, at, term)) return failure('index entry insert failed')
  return success(`Marked "${term}" as an index entry in block ${blockIndex}.`, 'mark index entry', {
    blockIndex,
    term,
  })
}

function indexInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const doc = ctx.editor.state.doc
  let afterBlockIndex: number | null = null
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    afterBlockIndex = parsed
  }
  const entries: string[] = []
  doc.descendants((node) => {
    if (node.type.name === 'docXeMark') entries.push(String(node.attrs.term).trim())
  })
  if (entries.length === 0) {
    return failure('the document has no index entries; mark terms first with reference.index.mark')
  }
  const unique = [...new Set(entries.filter(Boolean))]
  const where = afterBlockIndex === null ? 'at the caret' : `after block ${afterBlockIndex}`
  if (ctx.dryRun) {
    return planned(`would insert an index field with ${unique.length} entr(ies) ${where}`, {
      terms: unique.length,
    })
  }
  if (afterBlockIndex !== null) placeCaretInBlock(ctx, afterBlockIndex)
  else ensureInlineCaret(ctx)
  const result = insertIndexField(ctx.editor)
  if (!result.ok) return failure(result.error ?? 'index field insert failed')
  return success(
    `Inserted an index field with ${result.terms?.length ?? unique.length} entr(ies) ${where}.`,
    'insert index',
    { terms: result.terms?.length ?? unique.length },
  )
}

export const referenceOperations: OperationDef[] = [
  {
    name: 'reference.bookmark.insert',
    group: 'reference',
    description:
      'reference.bookmark.insert { blockIndex, name } — add a bookmark to a paragraph/heading/list-item block '
      + '(names: letters/CJK/underscore, no spaces; used later by reference.crossRef.insert)',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0 },
        name: { type: 'string', minLength: 1, maxLength: 40 },
      },
      required: ['blockIndex', 'name'],
    },
    handler: bookmarkInsertOp,
  },
  {
    name: 'reference.bookmark.remove',
    group: 'reference',
    description: 'reference.bookmark.remove { name } — delete a bookmark by name',
    inputSchema: {
      type: 'object',
      properties: { name: { type: 'string', minLength: 1, maxLength: 40 } },
      required: ['name'],
    },
    handler: bookmarkRemoveOp,
  },
  {
    name: 'reference.crossRef.insert',
    group: 'reference',
    description:
      'reference.crossRef.insert { bookmark, afterBlockIndex? } — insert a live REF field to the bookmark '
      + '(displays the bookmark paragraph text) at the end of a block or at the caret',
    inputSchema: {
      type: 'object',
      properties: {
        bookmark: { type: 'string', minLength: 1 },
        afterBlockIndex: { type: 'integer', minimum: 0 },
      },
      required: ['bookmark'],
    },
    handler: crossRefInsertOp,
  },
  {
    name: 'reference.caption.insert',
    group: 'reference',
    description:
      'reference.caption.insert { label, text?, blockIndex? } — insert a numbered caption with a live SEQ field '
      + '(label e.g. Figure/Table/Equation; number continues the existing sequence for that label)',
    inputSchema: {
      type: 'object',
      properties: {
        label: { type: 'string', minLength: 1, maxLength: 60 },
        text: { type: 'string', maxLength: 500 },
        blockIndex: { type: 'integer', minimum: 0, description: 'insert the caption after this block' },
      },
      required: ['label'],
    },
    handler: captionInsertOp,
  },
  {
    name: 'reference.source.add',
    group: 'reference',
    description:
      'reference.source.add { title, author?, year?, publisher?, url?, type? } — add a bibliography source '
      + '(type default Book; Book/JournalArticle/InternetSite/Report/Misc); a unique tag is assigned',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', minLength: 1, maxLength: 300 },
        author: { type: 'string', maxLength: 300 },
        year: { type: 'string', maxLength: 20 },
        publisher: { type: 'string', maxLength: 200 },
        url: { type: 'string', maxLength: 500 },
        type: { type: 'string', enum: ['Book', 'JournalArticle', 'InternetSite', 'Report', 'Misc'] },
      },
      required: ['title'],
    },
    handler: sourceAddOp,
  },
  {
    name: 'reference.citation.insert',
    group: 'reference',
    description:
      'reference.citation.insert { sourceIndex? | sourceId?, afterBlockIndex? } — insert an inline citation '
      + '"(Author, Year)" for a bibliography source (index from the source list, or its tag as sourceId)',
    inputSchema: {
      type: 'object',
      properties: {
        sourceIndex: { type: 'integer', minimum: 0 },
        sourceId: { type: 'string', description: 'source tag, e.g. "Wang2024"' },
        afterBlockIndex: { type: 'integer', minimum: 0 },
      },
    },
    handler: citationInsertOp,
  },
  {
    name: 'reference.bibliography.insert',
    group: 'reference',
    description:
      'reference.bibliography.insert { afterBlockIndex? } — insert a generated bibliography (heading + one '
      + 'line per bibliography source) at the end of a block or at the caret',
    inputSchema: {
      type: 'object',
      properties: { afterBlockIndex: { type: 'integer', minimum: 0 } },
    },
    handler: bibliographyInsertOp,
  },
  {
    name: 'reference.index.mark',
    group: 'reference',
    description:
      'reference.index.mark { blockIndex, term } — mark a term as an index entry (XE field) at the end of a block',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0 },
        term: { type: 'string', minLength: 1, maxLength: 200 },
      },
      required: ['blockIndex', 'term'],
    },
    handler: indexMarkOp,
  },
  {
    name: 'reference.index.insert',
    group: 'reference',
    description:
      'reference.index.insert { afterBlockIndex? } — insert a generated INDEX field for all marked index entries '
      + '(mark terms first with reference.index.mark)',
    inputSchema: {
      type: 'object',
      properties: { afterBlockIndex: { type: 'integer', minimum: 0 } },
    },
    handler: indexInsertOp,
  },
]