/**
 * Markdown comment anchors.
 *
 * Unlike docs (where comment ids live in a TipTap mark that round-trips through
 * the docx), the markdown document must stay pure — anchors are stored in a
 * sidecar file (`<file>.comments.json`) and resolved against the live document
 * here: a comment pins its top-level block index plus an exact excerpt, and the
 * highlight is a ProseMirror decoration (never part of the doc, never
 * serialized). Resolution is fuzzy by design: the excerpt is searched in the
 * recorded block first, then across all blocks (a paragraph inserted above
 * shifts indices); a 40-char head fallback tolerates edits at the excerpt tail;
 * the recorded character offset is the last resort (position-anchored, Word
 * parity). When nothing matches the anchor dangles harmlessly (the panel still
 * shows the comment, Word shows those as unanchored too).
 */
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import { Extension as ExtensionBuilder } from '@tiptap/core'
import type { Extension, Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { MdComment } from '../../shared/ipc'

export const commentAnchorsKey = new PluginKey<DecorationSet>('mdCommentAnchors')

/** smallest unused numeric comment id (sidecar ids stay stable once written) */
export function nextCommentId(comments: MdComment[]): string {
  const max = comments.reduce((acc, c) => Math.max(acc, parseInt(c.id, 10) || 0), 0)
  return String(max + 1)
}

interface BlockRange {
  from: number
  to: number
  text: string
}

/** top-level block boundaries: [{from, to, text}] for blockIndex → range mapping */
export function topLevelBlocks(editor: Editor): BlockRange[] {
  return blocksFromDoc(editor.state.doc)
}

function blocksFromDoc(doc: PmNode): BlockRange[] {
  const out: BlockRange[] = []
  let pos = 0
  for (let i = 0; i < doc.childCount; i++) {
    const child = doc.child(i)
    out.push({ from: pos + 1, to: pos + child.nodeSize - 1, text: child.textContent })
    pos += child.nodeSize
  }
  return out
}

/**
 * Anchor captured from the current selection: block index + exact excerpt +
 * character offset. This is what gets persisted in the sidecar.
 */
export function captureSelectionAnchor(editor: Editor): {
  blockIndex: number
  excerpt: string
  offset: number
  from: number
  to: number
} | null {
  const { from, to, empty } = editor.state.selection
  if (empty || from >= to) return null
  const blocks = topLevelBlocks(editor)
  // the anchor block is the one containing the selection start
  let blockIndex = -1
  for (let i = 0; i < blocks.length; i++) {
    if (from >= blocks[i].from - 1 && from <= blocks[i].to + 1) {
      blockIndex = i
      break
    }
  }
  if (blockIndex < 0) return null
  const block = blocks[blockIndex]
  const text = editor.state.doc.textBetween(from, to, ' ', ' ')
  const excerpt = text.replace(/\s+/g, ' ').trim().slice(0, 120)
  if (!excerpt) return null
  const blockText = block.text.replace(/\s+/g, ' ')
  const offset = Math.max(0, blockText.indexOf(excerpt))
  return { blockIndex, excerpt, offset, from, to }
}

const normalize = (s: string): string => s.replace(/\s+/g, ' ')

/**
 * Map an index into the whitespace-normalized form of `raw` back to an index
 * into `raw` itself (runs of whitespace collapse to one space). Clamped to
 * raw.length.
 */
function rawIndexForNormalized(raw: string, normalizedIndex: number): number {
  let rawPos = 0
  let normPos = 0
  while (normPos < normalizedIndex && rawPos < raw.length) {
    if (/\s/.test(raw[rawPos])) {
      while (rawPos < raw.length && /\s/.test(raw[rawPos])) rawPos++
      normPos += 1
    } else {
      rawPos++
      normPos++
    }
  }
  return Math.min(rawPos, raw.length)
}

/** locate `needle` (already normalized) inside a block → absolute range, or null */
function locate(blocks: BlockRange[], blockIndex: number, needle: string) {
  const block = blocks[blockIndex]
  if (!block) return null
  const norm = normalize(block.text)
  const idx = norm.indexOf(needle)
  if (idx < 0) return null
  const rawFrom = rawIndexForNormalized(block.text, idx)
  const rawTo = rawIndexForNormalized(block.text, idx + needle.length)
  const from = block.from + rawFrom
  const to = Math.min(block.to, block.from + Math.max(rawFrom + 1, rawTo))
  return to > from ? { from, to } : null
}

/**
 * Resolve a sidecar anchor against the live document → decoration range, or
 * null when nothing matches (dangling anchor).
 */
function resolveRange(blocks: BlockRange[], comment: MdComment): { from: number; to: number } | null {
  const excerpt = normalize(comment.excerpt).trim()
  if (!excerpt) return null
  // search order: recorded block first, then every block (indices shift when
  // paragraphs are inserted or removed above the anchor)
  const order: number[] = []
  if (comment.blockIndex >= 0 && comment.blockIndex < blocks.length) order.push(comment.blockIndex)
  for (let i = 0; i < blocks.length; i++) if (!order.includes(i)) order.push(i)
  for (const i of order) {
    const hit = locate(blocks, i, excerpt)
    if (hit) return hit
  }
  // fuzzy: a 40-char head still locates the region when the excerpt tail was edited
  const head = excerpt.slice(0, 40)
  if (head && head !== excerpt) {
    for (const i of order) {
      const hit = locate(blocks, i, head)
      if (hit) return hit
    }
  }
  // last resort: stay position-anchored in the recorded block (Word keeps
  // comments at their position when the text under them changed)
  const block = blocks[comment.blockIndex]
  if (block && block.text.trim()) {
    const from = Math.min(block.to, block.from + Math.min(comment.offset, block.text.length - 1))
    const to = Math.min(block.to, from + Math.min(excerpt.length, 8))
    if (to > from) return { from, to }
  }
  return null
}

/** Editor-facing resolver (jump, panel dangling state) */
export function resolveAnchor(
  editor: Editor,
  comment: MdComment,
): { from: number; to: number } | null {
  return resolveRange(topLevelBlocks(editor), comment)
}

/**
 * Decoration extension: highlights every anchored comment range in the body.
 * `getComments` is read at plugin apply time; the plugin rebuilds on the
 * `mdCommentsChanged` meta so comment edits (not doc edits) also refresh.
 */
export function CommentAnchorsExtension(getComments: () => MdComment[]): Extension {
  return ExtensionBuilder.create({
    name: 'mdCommentAnchors',
    addProseMirrorPlugins() {
      return [
        new Plugin<{ deco: DecorationSet }>({
          key: commentAnchorsKey,
          state: {
            init: () => ({ deco: DecorationSet.empty }),
            apply(tr, value) {
              const changed = tr.getMeta('mdCommentsChanged')
              if (!tr.docChanged && changed === undefined) return value
              // comment list edits can't be mapped through the transaction — rebuild
              if (changed !== undefined) return { deco: DecorationSet.empty }
              return { deco: value.deco.map(tr.mapping, tr.doc) }
            },
          },
          props: {
            decorations(state) {
              const comments = getComments()
              if (comments.length === 0) return DecorationSet.empty
              const blocks = blocksFromDoc(state.doc)
              const decos: Decoration[] = []
              for (const comment of comments) {
                if (comment.done) continue
                const range = resolveRange(blocks, comment)
                if (!range) continue
                decos.push(
                  Decoration.inline(range.from, range.to, {
                    class: 'md-comment-anchor',
                    'data-md-comment-id': comment.id,
                  }),
                )
              }
              return DecorationSet.create(state.doc, decos)
            },
          },
        }),
      ]
    },
  })
}

/** jump to a comment's anchor (scroll + flash), false when dangling */
export function jumpToComment(editor: Editor, id: string, comments: MdComment[]): boolean {
  const threadIds = new Set([id, ...comments.filter((c) => c.parentId === id).map((c) => c.id)])
  const root = comments.find((c) => c.id === id) ?? comments.find((c) => c.parentId === id)
  if (!root) return false
  const range = resolveAnchor(editor, root)
  if (!range) return false
  editor.view.dispatch(
    editor.state.tr
      .setSelection(TextSelection.near(editor.state.doc.resolve(range.from)))
      .scrollIntoView(),
  )
  const el = editor.view.domAtPos(range.from)
  const target = el.node instanceof HTMLElement ? el.node : editor.view.dom
  target.scrollIntoView({ behavior: 'smooth', block: 'center' })
  // flash every anchor span carrying this thread's ids
  for (const span of document.querySelectorAll<HTMLElement>('.md-comment-anchor')) {
    if (!threadIds.has(String(span.dataset.mdCommentId ?? ''))) continue
    span.classList.remove('md-comment-flash')
    void span.offsetWidth
    span.classList.add('md-comment-flash')
  }
  return true
}
