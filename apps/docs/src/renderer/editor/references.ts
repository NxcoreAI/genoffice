import type { Editor } from '@tiptap/core'
import {
  bibliographyLine,
  citationText,
  generateCaptionXml,
  generateIndexFieldXml,
  type Block,
  type SourceInfo,
} from '@genoffice/docx-engine'
import { t } from '../i18n/locale'

/**
 * Reference helpers shared by the References/Insert ribbon tabs and the host
 * agent's `reference.*` operations (single source of truth): bookmarks and
 * cross-references, captions (SEQ fields), the bibliography source list,
 * citations, bibliography generation and index entries/fields.
 *
 * Every helper returns a plain result instead of alerting, so the same code
 * can drive both dialog buttons and agent ops (the UI renders alerts).
 */

// ---- bookmarks / cross-references ----

export interface BookmarkEntry {
  name: string
  /** doc position of the owning paragraph-like node */
  pos: number
  preview: string
}

const BOOKMARK_PARA_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])
const BOOKMARK_NAME_RE = /^[A-Za-z一-鿿_][\w一-鿿]*$/

/** all bookmarks in document order: name + owning node position + preview text */
export function collectBookmarks(editor: Editor): BookmarkEntry[] {
  const out: BookmarkEntry[] = []
  editor.state.doc.descendants((node, pos) => {
    const names = node.attrs?.bookmarks as string[] | null
    if (Array.isArray(names)) {
      for (const name of names) {
        out.push({ name, pos, preview: node.textContent.slice(0, 40) })
      }
    }
    return !node.isLeaf
  })
  return out
}

/** the paragraph-like node at the cursor that can carry bookmarks */
export function bookmarkTargetAtCursor(editor: Editor): { pos: number; typeName: string } | null {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth >= 1; depth--) {
    const node = $from.node(depth)
    if (BOOKMARK_PARA_TYPES.has(node.type.name)) {
      return { pos: $from.before(depth), typeName: node.type.name }
    }
  }
  return null
}

export interface BookmarkResult {
  ok: boolean
  error?: string
  name?: string
}

/** Add a bookmark to the paragraph-like block at the cursor (Word Bookmark dialog add). */
export function insertBookmark(editor: Editor, rawName: string): BookmarkResult {
  const name = String(rawName ?? '').trim()
  if (!BOOKMARK_NAME_RE.test(name)) {
    return { ok: false, error: 'bookmark name must start with a letter/CJK character or "_" and contain no spaces' }
  }
  if (collectBookmarks(editor).some((entry) => entry.name === name)) {
    return { ok: false, error: `bookmark "${name}" already exists` }
  }
  const target = bookmarkTargetAtCursor(editor)
  if (!target) {
    return { ok: false, error: 'the caret is not inside a bookmarkable paragraph (place it in a paragraph/heading/list item)' }
  }
  const node = editor.state.doc.nodeAt(target.pos)
  if (!node) return { ok: false, error: 'target paragraph not found' }
  const existing = (node.attrs?.bookmarks as string[] | null) ?? []
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(target.pos, undefined, {
      ...node.attrs,
      bookmarks: [...existing, name],
    }),
  )
  return { ok: true, name }
}

/** Delete a bookmark by name from wherever it is anchored. */
export function removeBookmark(editor: Editor, rawName: string): boolean {
  const name = String(rawName ?? '').trim()
  const entry = collectBookmarks(editor).find((candidate) => candidate.name === name)
  if (!entry) return false
  const node = editor.state.doc.nodeAt(entry.pos)
  if (!node) return false
  const rest = ((node.attrs?.bookmarks as string[] | null) ?? []).filter((item) => item !== name)
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(entry.pos, undefined, {
      ...node.attrs,
      bookmarks: rest.length > 0 ? rest : null,
    }),
  )
  return true
}

export interface CrossRefResult {
  ok: boolean
  error?: string
  bookmark?: string
  /** text the REF field initially displays (the bookmark paragraph's text) */
  text?: string
}

/** Insert a REF field to `bookmarkName` at the cursor (Word Cross-reference). */
export function insertCrossRef(editor: Editor, bookmarkName: string): CrossRefResult {
  const bookmark = collectBookmarks(editor).find((entry) => entry.name === bookmarkName)
  if (!bookmark) {
    const names = collectBookmarks(editor).map((entry) => entry.name)
    return {
      ok: false,
      error: `bookmark "${bookmarkName}" does not exist`
        + (names.length > 0 ? `; available bookmarks: ${names.join(', ')}` : '; the document has no bookmarks'),
    }
  }
  editor
    .chain()
    .focus()
    .insertContent({
      type: 'text',
      text: bookmark.preview || bookmark.name,
      marks: [{ type: 'refField', attrs: { name: bookmark.name } }],
    })
    .run()
  return { ok: true, bookmark: bookmark.name, text: bookmark.preview || bookmark.name }
}

// ---- OOXML helper for captions/index (same resolution as the TOC path) ----

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
export function xmlOfNode(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  return blocks.find((block) => block.docxIndex === idx)?.originalXml ?? ''
}

// ---- captions ----

/** Next SEQ number for a caption label (counts existing SEQ fields of that label). */
export function nextCaptionNumber(editor: Editor, blocks: Block[], label: string): number {
  let count = 0
  const re = new RegExp(`SEQ\\s+${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}[\\s\\\\]`)
  editor.state.doc.forEach((node) => {
    if (node.type.name !== 'docProtected') return
    if (re.test(xmlOfNode(node as never, blocks))) count += 1
  })
  return count + 1
}

export interface CaptionResult {
  ok: boolean
  error?: string
  number?: number
  label?: string
}

/** Insert a centered caption paragraph with a live SEQ field at the cursor. */
export function insertCaption(
  editor: Editor,
  blocks: Block[],
  label: string,
  text = '',
): CaptionResult {
  const cleanLabel = String(label ?? '').trim()
  if (!cleanLabel) return { ok: false, error: 'caption label must be a non-empty string (e.g. "Figure")' }
  const cleanText = String(text ?? '').trim()
  const number = nextCaptionNumber(editor, blocks, cleanLabel)
  const xml = generateCaptionXml(cleanLabel, number, cleanText)
  const display = `${cleanLabel} ${number}${cleanText ? ` ${cleanText}` : ''}`
  editor
    .chain()
    .focus()
    .insertContent({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonCaption'),
        genXml: xml,
        fieldDisplay: { kind: 'text', left: display },
      },
    } as never)
    .run()
  return { ok: true, number, label: cleanLabel }
}

// ---- bibliography sources / citations ----

export interface SourceInput {
  author?: string
  title: string
  year?: string
  publisher?: string
  url?: string
  type?: string
}

export interface SourceResult {
  ok: boolean
  error?: string
  source?: SourceInfo
}

/** Build a SourceInfo with a unique tag from the Create Source fields. */
export function makeSourceInfo(existing: SourceInfo[], input: SourceInput): SourceResult {
  const title = String(input.title ?? '').trim()
  if (!title) return { ok: false, error: 'title is required' }
  const author = String(input.author ?? '').trim()
  const year = String(input.year ?? '').trim()
  const publisher = String(input.publisher ?? '').trim()
  const url = String(input.url ?? '').trim()
  const type = String(input.type ?? 'Book').trim() || 'Book'
  // unique tag derived from author/title initials + year (same rule as the dialog)
  const base = (author || title).replace(/[\s,，、]+/g, '').slice(0, 8) || 'Src'
  let tag = `${base}${year}`
  let n = 1
  while (existing.some((source) => source.tag === tag)) tag = `${base}${year}_${++n}`
  return {
    ok: true,
    source: {
      tag,
      type,
      author,
      title,
      year,
      ...(publisher ? { publisher } : {}),
      ...(url ? { url } : {}),
    },
  }
}

/** Insert an inline citation "(Author, Year)" at the cursor. */
export function insertCitation(editor: Editor, source: SourceInfo): boolean {
  return editor
    .chain()
    .focus()
    .insertContent({ type: 'text', text: citationText(source) } as never)
    .run()
}

/** PM nodes for a generated bibliography (heading + one line per source). */
export function buildBibliographyNodes(sources: SourceInfo[]): Array<Record<string, unknown>> {
  return [
    {
      type: 'docHeading',
      attrs: { docxIndex: null, styleId: null, aiChanged: false, level: 1 },
      content: [{ type: 'text', text: t('ribbonBibliographyHeading') }],
    },
    ...sources.map((source) => ({
      type: 'docParagraph',
      attrs: { docxIndex: null, styleId: null, aiChanged: false },
      content: [{ type: 'text', text: bibliographyLine(source) }],
    })),
  ]
}

export interface BibliographyResult {
  ok: boolean
  error?: string
  lines?: number
}

/** Insert the bibliography (heading + APA-like lines per source) at the cursor. */
export function insertBibliography(editor: Editor, sources: SourceInfo[]): BibliographyResult {
  if (sources.length === 0) {
    return { ok: false, error: 'the document has no bibliography sources (add one with reference.source.add)' }
  }
  editor
    .chain()
    .focus()
    .insertContent(buildBibliographyNodes(sources) as never)
    .run()
  return { ok: true, lines: sources.length }
}

// ---- index entries / index field ----

/** Insert an XE (index entry) mark at `pos` for `term`. */
export function markIndexEntry(editor: Editor, pos: number, term: string): boolean {
  const trimmed = String(term ?? '').trim()
  if (!trimmed) return false
  const at = Math.min(Math.max(0, pos), editor.state.doc.content.size)
  return editor
    .chain()
    .focus()
    .insertContentAt(at, { type: 'docXeMark', attrs: { term: trimmed } } as never)
    .run()
}

export interface IndexResult {
  ok: boolean
  error?: string
  terms?: string[]
}

/** Insert a generated INDEX field for every XE term in the document. */
export function insertIndexField(editor: Editor): IndexResult {
  const terms: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'docXeMark') terms.push(String(node.attrs.term))
  })
  if (terms.length === 0) {
    return {
      ok: false,
      error: 'the document has no index entries (mark terms first with reference.index.mark)',
    }
  }
  const unique = [...new Set(terms.map((term) => term.trim()).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, 'zh-CN'),
  )
  const nodes = generateIndexFieldXml(unique).map((xml, i) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonIndexFieldLabel'),
      genXml: xml,
      fieldDisplay: { kind: 'tocLine', left: unique[i], right: '', level: 1 },
    },
  }))
  editor
    .chain()
    .focus()
    .insertContent(nodes as never)
    .run()
  return { ok: true, terms: unique }
}