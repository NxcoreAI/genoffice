import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { WORDART_PRESETS, type WordArtPreset } from '@genoffice/ui'
import { latexToOmml, type TextboxDisplay } from '@genoffice/docx-engine'
import { equationBlockJson, inlineEquationNodeJson } from '../../editor/equation'
import {
  RELATIVE_HEIGHT_BASE,
  readShapeRelativeHeight,
  relativeHeightForRank,
} from '../../editor/shape-z-order'
import {
  applyPictureBytesToImageBlock,
  cropImageDataUrl,
  imageBlockTargetAt,
  parseImagePayload,
  pictureNaturalSize,
  removeImageBackgroundDataUrl,
} from '../../editor/picture-ops'
import {
  insertTextboxAt,
  insertWordArtAt,
  setParaAttrs,
} from '../../components/ribbon-tabs'
import {
  insertEquationFromLatex,
  insertInlineEquationFromLatex,
} from '../../components/EquationModal'
import {
  executeCommands,
  resolveTargetBlocks,
  validateCommandTarget,
  type Target,
} from '../commands'
import type { OperationContext, OperationDef, OperationOutcome } from './types'
import {
  blockEndPos,
  blockStartPos,
  blocksRange,
  checkBlockIndex,
  failure,
  hexOf,
  intOf,
  placeCaretAtBlockEnd,
  placeCaretAtDocEnd,
  planned,
  success,
  textBlocksOf,
} from './shared'

/**
 * B3 object ops: insert/format text boxes, WordArt, drop caps, equations,
 * symbols; and adjust image/shape display attributes. Insertion reuses the
 * real ribbon helpers (insertTextboxAt / insertWordArtAt / equation inserters);
 * formatting patches the same display attrs the ribbon's context menus write,
 * which the save path turns back into OOXML (textbox paras/sizes/styles patches
 * and imagePatchOf). Replacing image bytes, cropping or background removal is
 * NOT available here (the host owns image bytes) and fails with a clear error.
 */

const WRAP_VALUES = ['square-left', 'square-right', 'topBottom', 'behind', 'front'] as const
const POSITION_H = ['left', 'center', 'right'] as const
const POSITION_V = ['top', 'center', 'bottom'] as const
const IMAGE_ALIGN = ['left', 'center', 'right'] as const
const MAX_TEXT = 5000

/** Top-level block index an insertion helper will produce for the current caret. */
function insertionIndexForCaret(editor: Editor): number {
  const { $from } = editor.state.selection
  if ($from.depth > 0) return $from.index(0) + 1
  const doc = editor.state.doc
  let count = 0
  let pos = 0
  doc.forEach((node) => {
    if (pos + node.nodeSize <= $from.pos) count++
    pos += node.nodeSize
  })
  return count
}

/** Place the caret where the insert helper appends, and predict the new block index. */
function prepareInsertion(
  editor: Editor,
  afterBlockIndex: unknown,
): { expectedIndex: number } | { error: string } {
  const doc = editor.state.doc
  if (afterBlockIndex !== undefined) {
    const parsed = intOf(afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return { error: `afterBlockIndex must be an integer 0..${doc.childCount - 1}` }
    }
    if (!placeCaretAtBlockEnd(editor, parsed)) {
      return { error: `afterBlockIndex must be an integer 0..${doc.childCount - 1}` }
    }
    return { expectedIndex: parsed + 1 }
  }
  if (!editor.state.selection.$from.parent.inlineContent) placeCaretAtDocEnd(editor)
  return { expectedIndex: insertionIndexForCaret(editor) }
}

interface InsertedBlock {
  pos: number
  index: number
  attrs: Record<string, unknown>
}

/** Verify the helper added a fresh editor-generated object block at `expectedIndex`. */
function verifyInsertedBlock(
  editor: Editor,
  expectedIndex: number,
  kind: 'textbox' | 'wordart',
): InsertedBlock | { error: string } {
  const doc = editor.state.doc
  if (expectedIndex >= doc.childCount) {
    return { error: 'insert failed: no block was added at the expected position' }
  }
  const node = doc.child(expectedIndex)
  const boxes = node.attrs?.textboxes as TextboxDisplay[] | undefined
  if (
    node.type.name !== 'docProtected'
    || node.attrs?.docxIndex !== null
    || typeof node.attrs?.genXml !== 'string'
    || !Array.isArray(boxes)
    || boxes.length === 0
  ) {
    return { error: 'insert failed: the expected block is not a newly created object block' }
  }
  if (kind === 'wordart' && !boxes[0].wordArtId) {
    return { error: 'insert failed: the expected block is not a WordArt block' }
  }
  return { pos: blockStartPos(doc, expectedIndex), index: expectedIndex, attrs: node.attrs }
}

function withBoxText(box: TextboxDisplay, text: string): TextboxDisplay {
  const firstPara = box.paras[0]
  if (box.wordArtId) {
    // WordArt strings never wrap and keep their preset run formatting
    const firstRun = firstPara?.runs[0] as unknown as Record<string, unknown> | undefined
    return {
      ...box,
      paras: [{ ...(firstPara ?? { runs: [] }), runs: [{ ...(firstRun ?? {}), text }] }],
    }
  }
  return {
    ...box,
    paras: [{ runs: [{ text }], align: firstPara?.align ?? 'left' }],
  }
}

function textboxInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  let text: string | undefined
  if (input.text !== undefined && input.text !== null) {
    text = String(input.text)
    if (text.length > MAX_TEXT) return failure(`text must be at most ${MAX_TEXT} characters`)
  }
  const doc = ctx.editor.state.doc
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
  }
  if (ctx.dryRun) {
    return planned(
      `would insert a text box ${input.afterBlockIndex !== undefined ? `after block ${String(input.afterBlockIndex)}` : 'at the caret'}`
      + (text ? ` with text "${text.slice(0, 40)}"` : ''),
    )
  }
  const prepared = prepareInsertion(ctx.editor, input.afterBlockIndex)
  if ('error' in prepared) return failure(prepared.error)
  insertTextboxAt(ctx.editor)
  const inserted = verifyInsertedBlock(ctx.editor, prepared.expectedIndex, 'textbox')
  if ('error' in inserted) return failure(inserted.error)
  if (text !== undefined) {
    const boxes = inserted.attrs.textboxes as TextboxDisplay[]
    ctx.editor.view.dispatch(
      ctx.editor.state.tr.setNodeMarkup(inserted.pos, undefined, {
        ...inserted.attrs,
        textboxes: [withBoxText(boxes[0], text), ...boxes.slice(1)],
      }),
    )
  }
  return success(
    `Inserted a text box at block ${inserted.index}${text !== undefined ? ' with text' : ''}.`,
    'insert text box',
    { blockIndex: inserted.index, ...(text !== undefined ? { chars: text.length } : {}) },
  )
}

function wordArtInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const presetId = input.preset === undefined ? WORDART_PRESETS[0].id : String(input.preset)
  const preset: WordArtPreset | undefined =
    WORDART_PRESETS.find((candidate) => candidate.id === presetId)
  if (!preset) {
    return failure(`preset must be one of ${WORDART_PRESETS.map((candidate) => candidate.id).join('/')}`)
  }
  let text: string | undefined
  if (input.text !== undefined && input.text !== null) {
    text = String(input.text)
    if (!text.trim()) return failure('text must be a non-empty string when provided')
    if (text.length > MAX_TEXT) return failure(`text must be at most ${MAX_TEXT} characters`)
  }
  const doc = ctx.editor.state.doc
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
  }
  if (ctx.dryRun) {
    return planned(
      `would insert a WordArt "${preset.id}" ${input.afterBlockIndex !== undefined ? `after block ${String(input.afterBlockIndex)}` : 'at the caret'}`,
    )
  }
  const prepared = prepareInsertion(ctx.editor, input.afterBlockIndex)
  if ('error' in prepared) return failure(prepared.error)
  insertWordArtAt(ctx.editor, preset)
  const inserted = verifyInsertedBlock(ctx.editor, prepared.expectedIndex, 'wordart')
  if ('error' in inserted) return failure(inserted.error)
  if (text !== undefined) {
    const boxes = inserted.attrs.textboxes as TextboxDisplay[]
    ctx.editor.view.dispatch(
      ctx.editor.state.tr.setNodeMarkup(inserted.pos, undefined, {
        ...inserted.attrs,
        textboxes: [withBoxText(boxes[0], text), ...boxes.slice(1)],
      }),
    )
  }
  return success(
    `Inserted WordArt "${preset.id}" at block ${inserted.index}${text !== undefined ? ` with text "${text.slice(0, 40)}"` : ''}.`,
    'insert WordArt',
    { blockIndex: inserted.index, preset: preset.id },
  )
}

const DROP_CAP_ATTRS: Record<string, string | null> = {
  none: null,
  dropped: JSON.stringify({ type: 'drop', lines: 3 }),
  margin: JSON.stringify({ type: 'margin', lines: 3 }),
}

function dropCapOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const mode = String(input.mode ?? '')
  if (!(mode in DROP_CAP_ATTRS)) {
    return failure('mode must be one of none/dropped/margin')
  }
  const targetError = validateCommandTarget(input.target)
  if (targetError) return failure(targetError)
  const doc = ctx.editor.state.doc
  const blocks = textBlocksOf(
    resolveTargetBlocks(doc, input.target as Target, ctx.editor.state.selection),
  )
  if (blocks.length === 0) {
    return failure('target matched no text blocks (a drop cap applies to paragraph text)')
  }
  const range = blocksRange(blocks)
  if (!range) return failure('target blocks have no text range')
  if (ctx.dryRun) {
    return planned(`would set drop cap "${mode}" on ${blocks.length} block(s)`)
  }
  setParaAttrs(ctx.editor, { dropCap: DROP_CAP_ATTRS[mode] }, range)
  return success(
    mode === 'none'
      ? `Removed the drop cap from ${blocks.length} block(s).`
      : `Set a "${mode}" drop cap on ${blocks.length} block(s).`,
    'drop cap',
    { blocks: blocks.length, mode },
  )
}

function equationInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const latex = typeof input.latex === 'string' ? input.latex.trim() : ''
  if (!latex) return failure('latex must be a non-empty string')
  if (latex.length > 10_000) return failure('latex must be at most 10000 characters')
  try {
    // validate against the same converter the insert helpers use (throws on unsupported syntax)
    latexToOmml(latex)
  } catch (error) {
    return failure(`invalid LaTeX: ${error instanceof Error ? error.message : String(error)}`)
  }
  const inline = input.inline === true
  const doc = ctx.editor.state.doc
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
  }
  const where = input.afterBlockIndex !== undefined ? `after block ${String(input.afterBlockIndex)}` : 'at the caret'
  if (ctx.dryRun) {
    return planned(`would insert a ${inline ? 'inline ' : ''}equation ${where}`, { latex, inline })
  }
  if (input.afterBlockIndex !== undefined) {
    placeCaretAtBlockEnd(ctx.editor, intOf(input.afterBlockIndex) as number)
  } else if (!ctx.editor.state.selection.$from.parent.inlineContent) {
    placeCaretAtDocEnd(ctx.editor)
  }
  const before = ctx.editor.state.doc
  if (inline) insertInlineEquationFromLatex(ctx.editor, latex)
  else insertEquationFromLatex(ctx.editor, latex)
  if (ctx.editor.state.doc === before) {
    return failure('equation insert failed (no valid caret position for the equation)')
  }
  return success(
    `Inserted ${inline ? 'an inline ' : 'a display '}equation ${where}.`,
    'insert equation',
    { inline, ...(input.afterBlockIndex !== undefined ? { afterBlockIndex: intOf(input.afterBlockIndex) } : {}) },
  )
}

function symbolInsertOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const char = typeof input.char === 'string' ? input.char : ''
  if (!char) return failure('char must be a non-empty string')
  if ([...char].length > 16) return failure('char must be at most 16 characters')
  // eslint-disable-next-line no-control-regex
  if (/[\u0000-\u001f\u007f]/.test(char)) return failure('char must not contain control characters or line breaks')
  const ownParagraph = input.ownParagraph === true
  const doc = ctx.editor.state.doc
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
  }
  const where = ownParagraph
    ? 'in its own paragraph'
    : input.afterBlockIndex !== undefined
      ? `at the end of block ${String(input.afterBlockIndex)}`
      : 'at the caret'
  if (ctx.dryRun) return planned(`would insert symbol "${char}" ${where}`)
  if (ownParagraph) {
    const pos =
      input.afterBlockIndex !== undefined
        ? blockEndPos(doc, intOf(input.afterBlockIndex) as number)
        : doc.content.size
    const before = doc.childCount
    ctx.editor
      .chain()
      .insertContentAt(pos, {
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [{ type: 'text', text: char }],
      })
      .run()
    if (ctx.editor.state.doc.childCount <= before) {
      return failure('symbol paragraph insertion failed')
    }
  } else {
    const before = doc
    if (input.afterBlockIndex !== undefined) {
      placeCaretAtBlockEnd(ctx.editor, intOf(input.afterBlockIndex) as number)
    } else if (!ctx.editor.state.selection.$from.parent.inlineContent) {
      placeCaretAtDocEnd(ctx.editor)
    }
    ctx.editor.chain().focus().insertContent({ type: 'text', text: char }).run()
    if (ctx.editor.state.doc === before) return failure('symbol insertion failed')
  }
  return success(`Inserted symbol "${char}" ${where}.`, 'insert symbol', { char })
}

/** Merge display-attr patches (the same attrs the ribbon image/shape handlers write) and dispatch. */
function patchProtectedAttrs(
  ctx: OperationContext,
  pos: number,
  attrs: Record<string, unknown>,
  patch: Record<string, unknown>,
): boolean {
  const before = ctx.editor.state.doc
  ctx.editor.view.dispatch(
    ctx.editor.state.tr.setNodeMarkup(pos, undefined, { ...attrs, ...patch }),
  )
  return ctx.editor.state.doc !== before
}

function imageFormatOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  if (node.type.name !== 'docProtected' || node.attrs.blockType !== 'image') {
    return failure(`block ${blockIndex} is a ${node.type.name}, not an image block`)
  }
  let blockPos = blockStartPos(doc, blockIndex)
  const patch: Record<string, unknown> = {}
  const changed: string[] = []
  // width/height/align go through the real apply_commands image path (proportional scaling included)
  const imageCommand: Record<string, unknown> = {}
  const fields: string[] = []
  if (input.widthPx !== undefined) {
    const width = Number(input.widthPx)
    if (!Number.isFinite(width) || width < 16 || width > 5000) return failure('widthPx must be 16..5000')
    imageCommand.widthPx = Math.round(width)
    fields.push('widthPx')
  }
  if (input.heightPx !== undefined) {
    const height = Number(input.heightPx)
    if (!Number.isFinite(height) || height < 16 || height > 5000) return failure('heightPx must be 16..5000')
    imageCommand.heightPx = Math.round(height)
    fields.push('heightPx')
  }
  if (input.align !== undefined) {
    const align = input.align === null ? null : String(input.align)
    if (align !== null && !(IMAGE_ALIGN as readonly string[]).includes(align)) {
      return failure('align must be left/center/right or null')
    }
    imageCommand.align = align
    fields.push('align')
  }
  if (input.wrap !== undefined) {
    const wrap = input.wrap === null ? null : String(input.wrap)
    if (wrap !== null && !(WRAP_VALUES as readonly string[]).includes(wrap)) {
      return failure(`wrap must be one of ${WRAP_VALUES.join('/')} or null (inline)`)
    }
    patch.imageWrap = wrap
    if (wrap === null) {
      // inline: position presets/offsets are meaningless (ribbon clears them too)
      Object.assign(patch, {
        imagePosH: null,
        imagePosV: null,
        imageOffsetXEmu: null,
        imageOffsetYEmu: null,
      })
    }
    changed.push('wrap')
  }
  if (input.position !== undefined) {
    const raw = input.position
    if (!raw || typeof raw !== 'object') {
      return failure('position must be { horizontal: left|center|right, vertical: top|center|bottom }')
    }
    const p = raw as Record<string, unknown>
    const horizontal = String(p.horizontal ?? '')
    const vertical = String(p.vertical ?? '')
    if (!(POSITION_H as readonly string[]).includes(horizontal)) {
      return failure(`position.horizontal must be one of ${POSITION_H.join('/')}`)
    }
    if (!(POSITION_V as readonly string[]).includes(vertical)) {
      return failure(`position.vertical must be one of ${POSITION_V.join('/')}`)
    }
    Object.assign(patch, {
      imageWrap: horizontal === 'right' ? 'square-right' : 'square-left',
      imagePosH: horizontal,
      imagePosV: vertical,
      imageOffsetXEmu: null,
      imageOffsetYEmu: null,
    })
    changed.push('position')
  }
  if (input.rotationDeg !== undefined) {
    const raw = Number(input.rotationDeg)
    if (!Number.isFinite(raw) || raw < -3600 || raw > 3600) {
      return failure('rotationDeg must be a number between -3600 and 3600')
    }
    const normalized = ((Math.round(raw) % 360) + 360) % 360
    patch.imageRotDeg = normalized === 0 ? null : normalized
    changed.push('rotationDeg')
  }
  if (input.flipH !== undefined) {
    if (typeof input.flipH !== 'boolean') return failure('flipH must be a boolean')
    patch.imageFlipH = input.flipH
    changed.push('flipH')
  }
  if (input.flipV !== undefined) {
    if (typeof input.flipV !== 'boolean') return failure('flipV must be a boolean')
    patch.imageFlipV = input.flipV
    changed.push('flipV')
  }
  if (input.zOrder !== undefined) {
    const z = intOf(input.zOrder)
    if (z === null || z < -1000 || z > 1000) return failure('zOrder must be an integer between -1000 and 1000')
    patch.imageZOrder = z
    const currentWrap = String(node.attrs.imageWrap ?? '')
    if (currentWrap !== 'front' && currentWrap !== 'behind') patch.imageWrap = 'front'
    changed.push('zOrder')
  }
  if (fields.length === 0 && changed.length === 0) {
    return failure(
      'nothing to change: pass widthPx/heightPx/align/wrap/position/rotationDeg/flipH/flipV/zOrder',
    )
  }
  if (ctx.dryRun) {
    return planned(
      `would set ${[...fields, ...changed].join(', ')} on image block ${blockIndex}`,
      { blockIndex },
    )
  }
  if (fields.length > 0) {
    const outcome = executeCommands(ctx.editor, {
      commands: [
        {
          updateImageProperties: {
            target: { blockIndexes: [blockIndex] },
            properties: imageCommand as { widthPx?: number; heightPx?: number; align?: 'left' | 'center' | 'right' | null },
            fields: fields as Array<'widthPx' | 'heightPx' | 'align'>,
          },
        },
      ],
    })
    if (!outcome.ok) return failure(outcome.error ?? 'image size/alignment update failed')
  }
  if (Object.keys(patch).length > 0) {
    const current = ctx.editor.state.doc.child(blockIndex)
    blockPos = blockStartPos(ctx.editor.state.doc, blockIndex)
    if (!patchProtectedAttrs(ctx, blockPos, current.attrs, patch)) {
      return failure('image formatting produced no change (the attributes were already set?)')
    }
  }
  return success(
    `Updated image block ${blockIndex}: ${[...fields, ...changed].join(', ')}.`,
    'format image',
    { blockIndex, fields: [...fields, ...changed] },
  )
}

function shapeFormatOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  const boxes = node.attrs?.textboxes as TextboxDisplay[] | undefined
  if (node.type.name !== 'docProtected' || !Array.isArray(boxes) || boxes.length === 0) {
    return failure(
      `block ${blockIndex} is not a textbox/shape block (no textboxes display model); `
      + 'use object.image.format for pictures',
    )
  }
  const box = boxes[0]
  const nextBox: TextboxDisplay = { ...box }
  const patch: Record<string, unknown> = {}
  const changed: string[] = []
  if (input.fillHex !== undefined) {
    if (input.fillHex === null) nextBox.fill = undefined
    else {
      const hex = hexOf(input.fillHex)
      if (hex === undefined) return failure('fillHex must be 6-digit hex without #, or null to clear the fill')
      nextBox.fill = hex ?? undefined
    }
    changed.push('fillHex')
  }
  if (input.lineHex !== undefined) {
    if (input.lineHex === null) nextBox.borderColor = undefined
    else {
      const hex = hexOf(input.lineHex)
      if (hex === undefined) return failure('lineHex must be 6-digit hex without #, or null to remove the outline')
      nextBox.borderColor = hex ?? undefined
    }
    changed.push('lineHex')
  }
  if (input.text !== undefined) {
    const text = input.text === null ? '' : String(input.text)
    if (text.length > MAX_TEXT) return failure(`text must be at most ${MAX_TEXT} characters`)
    const withText = withBoxText(nextBox, text)
    nextBox.paras = withText.paras
    changed.push('text')
  }
  if (input.widthEmu !== undefined) {
    const width = intOf(input.widthEmu)
    if (width === null || width < 91440 || width > 7772400) {
      return failure('widthEmu must be an integer between 91440 and 7772400 (91440 EMU = 0.1 cm)')
    }
    nextBox.widthPx = Math.round(width / 9525)
    changed.push('widthEmu')
  }
  if (input.heightEmu !== undefined) {
    const height = intOf(input.heightEmu)
    if (height === null || height < 91440 || height > 7772400) {
      return failure('heightEmu must be an integer between 91440 and 7772400')
    }
    nextBox.heightPx = Math.round(height / 9525)
    changed.push('heightEmu')
  }
  if (input.zOrder !== undefined) {
    const z = intOf(input.zOrder)
    if (z === null || z < -1000 || z > 1000) return failure('zOrder must be an integer between -1000 and 1000')
    // z-order save-side patching exists for editor-generated shapes only: imported
    // shapes keep their original anchor bytes (relativeHeight included), so a
    // reorder there would silently not persist — fail instead of faking it.
    if (node.attrs.docxIndex !== null || typeof node.attrs.genXml !== 'string') {
      return failure(
        'zOrder on imported/anchored shapes is unsupported: only shapes/text boxes created in the editor '
        + 'can be re-ordered (use object.image.format for pictures)',
      )
    }
    patch.imageZOrder = z
    const currentWrap = String(node.attrs.imageWrap ?? '')
    if (currentWrap !== 'front' && currentWrap !== 'behind') patch.imageWrap = 'front'
    changed.push('zOrder')
  }
  if (changed.length === 0) {
    return failure('nothing to change: pass fillHex/lineHex/text/widthEmu/heightEmu/zOrder')
  }
  if (ctx.dryRun) {
    return planned(`would set ${changed.join(', ')} on shape/textbox block ${blockIndex}`, {
      blockIndex,
    })
  }
  patch.textboxes = [nextBox, ...boxes.slice(1)]
  if (!patchProtectedAttrs(ctx, blockStartPos(doc, blockIndex), node.attrs, patch)) {
    return failure('shape formatting produced no change (the attributes were already set?)')
  }
  return success(
    `Updated shape/textbox block ${blockIndex}: ${changed.join(', ')}.`,
    'format shape',
    { blockIndex, fields: changed },
  )
}

/** Update an existing equation in place (never inserts a second one). */
function equationUpdateOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const latex = typeof input.latex === 'string' ? input.latex.trim() : ''
  if (!latex) return failure('latex must be a non-empty string')
  if (latex.length > 10_000) return failure('latex must be at most 10000 characters')
  try {
    latexToOmml(latex)
  } catch (error) {
    return failure(`invalid LaTeX: ${error instanceof Error ? error.message : String(error)}`)
  }
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  const blockPos = blockStartPos(doc, blockIndex)
  let inlineIndex: number | null = null
  if (input.inlineIndex !== undefined) {
    const parsed = intOf(input.inlineIndex)
    if (parsed === null || parsed < 0) return failure('inlineIndex must be a non-negative integer')
    inlineIndex = parsed
  }
  const schema = ctx.editor.schema

  if (inlineIndex !== null) {
    const inlineMath: Array<{ pos: number; size: number }> = []
    // descendants positions are relative to the block CONTENT start (blockPos + 1)
    node.descendants((child, offset) => {
      if (child.type.name === 'docInlineMath') {
        inlineMath.push({ pos: blockPos + 1 + offset, size: child.nodeSize })
      }
    })
    if (inlineMath.length === 0) {
      return failure(
        `block ${blockIndex} contains no inline equation; omit inlineIndex to update a block equation`,
      )
    }
    if (inlineIndex >= inlineMath.length) {
      return failure(
        `inlineIndex ${inlineIndex} is out of range; block ${blockIndex} has ${inlineMath.length} inline equation(s)`,
      )
    }
    if (ctx.dryRun) {
      return planned(
        `would update inline equation ${inlineIndex} in block ${blockIndex} to the given LaTeX`,
      )
    }
    const target = inlineMath[inlineIndex]
    const next = schema.nodeFromJSON(inlineEquationNodeJson(latex))
    ctx.editor.view.dispatch(ctx.editor.state.tr.replaceWith(target.pos, target.pos + target.size, next))
    return success(
      `Updated inline equation ${inlineIndex} in block ${blockIndex}.`,
      'update equation',
      { blockIndex, inlineIndex },
    )
  }

  const isEquationBlock =
    node.type.name === 'docProtected' && !!node.attrs?.formulaDisplay
  if (!isEquationBlock) {
    return failure(
      `block ${blockIndex} is a ${node.type.name}, not an equation block`
        + (node.type.name === 'docProtected'
          ? ' (it has no equation display model; for an inline equation pass inlineIndex)'
          : ''),
    )
  }
  if (ctx.dryRun) return planned(`would update the block equation at block ${blockIndex} to the given LaTeX`)
  const next = schema.nodeFromJSON(equationBlockJson(latex))
  ctx.editor.view.dispatch(
    ctx.editor.state.tr.replaceWith(blockPos, blockPos + node.nodeSize, next),
  )
  return success(`Updated the block equation at block ${blockIndex}.`, 'update equation', {
    blockIndex,
  })
}

// ── A1: picture replace / crop / local background removal ──

/** Current in-editor bytes of an image block (parse populates imageDataUrl). */
function currentImageDataUrl(ctx: OperationContext, blockIndex: number): string | { error: string } {
  const target = imageBlockTargetAt(ctx.editor.state.doc, blockIndex)
  if ('error' in target) return target
  const dataUrl = target.node.attrs?.imageDataUrl
  if (typeof dataUrl !== 'string' || !dataUrl) {
    return { error: 'the image bytes are not available in the editor (reopen the document first)' }
  }
  return dataUrl
}

function imageReplaceOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): Promise<OperationOutcome> | OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const target = imageBlockTargetAt(ctx.editor.state.doc, blockIndex)
  if ('error' in target) return failure(target.error)
  const parsed = parseImagePayload({
    dataUrl: typeof input.dataUrl === 'string' ? input.dataUrl : undefined,
    base64: typeof input.base64 === 'string' ? input.base64 : undefined,
    mime: typeof input.mime === 'string' ? input.mime : undefined,
  })
  if (!parsed.ok) return failure(parsed.error)
  if (ctx.dryRun) {
    return planned(
      `would replace the picture bytes of block ${blockIndex} (${parsed.mime})`
        + (target.imported ? ' in place (imported image, drawing XML preserved)' : ' (editor-created image)'),
      { blockIndex, mime: parsed.mime, imported: target.imported },
    )
  }
  return applyPictureBytesToImageBlock(ctx.editor, blockIndex, parsed.dataUrl).then((result) =>
    result.ok
      ? success(
          `Replaced the picture in block ${blockIndex} (${parsed.mime}`
            + `${result.imported ? ', bytes swapped in place' : ', pending new image'}); `
            + `display size ${result.widthPx}×${result.heightPx}px.`,
          'replace picture',
          {
            blockIndex,
            mime: parsed.mime,
            imported: result.imported === true,
            widthPx: result.widthPx,
            heightPx: result.heightPx,
          },
        )
      : failure(result.error ?? 'picture replace failed'),
  )
}

function imageCropOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): Promise<OperationOutcome> | OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const target = imageBlockTargetAt(ctx.editor.state.doc, blockIndex)
  if ('error' in target) return failure(target.error)
  if (!input.rect || typeof input.rect !== 'object') {
    return failure('rect must be { x, y, w, h } — natural image pixels, or 0..1 fractions with unit:"fraction"')
  }
  const unit = input.unit === undefined ? 'px' : String(input.unit)
  if (unit !== 'px' && unit !== 'fraction') {
    return failure('unit must be "px" (default) or "fraction" (0..1 of the image natural size)')
  }
  const dataUrl = currentImageDataUrl(ctx, blockIndex)
  if (typeof dataUrl === 'object') return failure(dataUrl.error)
  return (async () => {
    const natural = await pictureNaturalSize(dataUrl)
    if (!natural) return failure('the current picture could not be decoded')
    let pxRect: Record<string, unknown> = input.rect as Record<string, unknown>
    if (unit === 'fraction') {
      const raw = input.rect as Record<string, unknown>
      const fractions = ['x', 'y', 'w', 'h'].map((key) => raw[key])
      if (fractions.some((value) => typeof value !== 'number' || !Number.isFinite(value))) {
        return failure('fraction rect.x/y/w/h must be finite numbers 0..1')
      }
      const [fx, fy, fw, fh] = fractions as number[]
      if (fx < 0 || fy < 0 || fw <= 0 || fh <= 0 || fx + fw > 1 + 1e-9 || fy + fh > 1 + 1e-9) {
        return failure('fraction rect must satisfy 0 ≤ x,y and x+w ≤ 1, y+h ≤ 1 with positive w/h')
      }
      pxRect = {
        x: Math.round(fx * natural.width),
        y: Math.round(fy * natural.height),
        w: Math.max(1, Math.round(fw * natural.width)),
        h: Math.max(1, Math.round(fh * natural.height)),
      }
    }
    const cropped = await cropImageDataUrl(dataUrl, pxRect)
    if (!cropped.ok || !cropped.dataUrl) return failure(cropped.error ?? 'crop failed')
    if (ctx.dryRun) {
      return planned(
        `would crop block ${blockIndex} to ${cropped.rect?.w}×${cropped.rect?.h}px at (${cropped.rect?.x}, ${cropped.rect?.y}) `
          + `of the ${natural.width}×${natural.height}px image (unit: ${unit})`,
        { blockIndex, unit, widthPx: cropped.widthPx, heightPx: cropped.heightPx },
      )
    }
    const applied = await applyPictureBytesToImageBlock(ctx.editor, blockIndex, cropped.dataUrl)
    if (!applied.ok) return failure(applied.error ?? 'crop apply failed')
    return success(
      `Cropped the picture in block ${blockIndex} to ${cropped.rect?.w}×${cropped.rect?.h}px; `
        + `display size ${applied.widthPx}×${applied.heightPx}px.`,
      'crop picture',
      {
        blockIndex,
        unit,
        widthPx: cropped.widthPx,
        heightPx: cropped.heightPx,
        rect: cropped.rect,
      },
    )
  })()
}

function imageRemoveBackgroundOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): Promise<OperationOutcome> | OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const target = imageBlockTargetAt(ctx.editor.state.doc, blockIndex)
  if ('error' in target) return failure(target.error)
  let tolerance = 25
  if (input.tolerance !== undefined) {
    const parsed = intOf(input.tolerance)
    if (parsed === null || parsed < 0 || parsed > 100) {
      return failure('tolerance must be an integer 0..100')
    }
    tolerance = parsed
  }
  if (input.sample !== undefined && String(input.sample) !== 'corners') {
    return failure('sample must be "corners" (the background colors are sampled from the image border ring)')
  }
  const dataUrl = currentImageDataUrl(ctx, blockIndex)
  if (typeof dataUrl === 'object') return failure(dataUrl.error)
  if (ctx.dryRun) {
    return planned(
      `would remove the background of block ${blockIndex} with tolerance ${tolerance} (local cutout)`,
      { blockIndex, tolerance },
    )
  }
  return (async () => {
    const result = await removeImageBackgroundDataUrl(dataUrl, tolerance)
    if (!result.ok || !result.dataUrl) return failure(result.error ?? 'background removal failed')
    const applied = await applyPictureBytesToImageBlock(ctx.editor, blockIndex, result.dataUrl)
    if (!applied.ok) return failure(applied.error ?? 'background-removed picture apply failed')
    return success(
      `Removed the background of block ${blockIndex} (${result.removedPct ?? 0}% of pixels cleared, `
        + `tolerance ${tolerance}); display size ${applied.widthPx}×${applied.heightPx}px.`,
      'remove background',
      {
        blockIndex,
        tolerance,
        removedPct: result.removedPct ?? 0,
        widthPx: applied.widthPx,
        heightPx: applied.heightPx,
      },
    )
  })()
}

// ── A1: shape stacking order (imported shapes patch their original anchor) ──

const SHAPE_Z_ORDERS = ['front', 'back', 'forward', 'backward', 'absolute'] as const
const MAX_RELATIVE_HEIGHT = 4294967295

function floatingRanks(ctx: OperationContext): { min: number; max: number } {
  const app = ctx.app
  const blocks = app?.getBlocks() ?? []
  let min = Number.POSITIVE_INFINITY
  let max = Number.NEGATIVE_INFINITY
  ctx.editor.state.doc.forEach((node) => {
    if (node.type.name !== 'docProtected') return
    const boxes = node.attrs?.textboxes as TextboxDisplay[] | null
    if (!boxes || boxes.length === 0) return
    const docxIndex = node.attrs?.docxIndex as number | null | undefined
    let rank: number | null = null
    if (docxIndex !== null && docxIndex !== undefined) {
      const original = blocks.find((block) => block.docxIndex === docxIndex)
      if (original?.originalXml) {
        const raw = readShapeRelativeHeight(original.originalXml)
        if (raw !== null) rank = raw - RELATIVE_HEIGHT_BASE
      }
    } else if (node.attrs?.shapeZOrder != null) {
      rank = Number(node.attrs.shapeZOrder) - RELATIVE_HEIGHT_BASE
    } else {
      rank = Number(node.attrs?.imageZOrder ?? 0)
    }
    if (rank === null || !Number.isFinite(rank)) return
    min = Math.min(min, rank)
    max = Math.max(max, rank)
  })
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { min: 0, max: 0 }
  return { min, max }
}

function shapeZOrderOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  const boxes = node.attrs?.textboxes as TextboxDisplay[] | null
  if (node.type.name !== 'docProtected' || !Array.isArray(boxes) || boxes.length === 0) {
    return failure(`block ${blockIndex} is not a shape/text box block`)
  }
  const order = String(input.order ?? '')
  if (!(SHAPE_Z_ORDERS as readonly string[]).includes(order)) {
    return failure(`order must be one of ${SHAPE_Z_ORDERS.join('/')}`)
  }
  const imported = node.attrs?.docxIndex !== null && node.attrs?.docxIndex !== undefined
  const app = ctx.app
  // validate the absolute target first — it needs no current anchor lookup
  let absoluteRaw: number | null = null
  if (order === 'absolute') {
    if (input.relativeHeight === undefined) {
      return failure('order "absolute" needs relativeHeight (the raw wp:anchor relativeHeight value)')
    }
    const parsed = intOf(input.relativeHeight)
    if (parsed === null || parsed < 0 || parsed > MAX_RELATIVE_HEIGHT) {
      return failure(`relativeHeight must be an integer 0..${MAX_RELATIVE_HEIGHT}`)
    }
    absoluteRaw = parsed
  }
  let currentRaw = 0
  if (imported) {
    if (!app) return failure('object.shape.zOrder on imported shapes needs the app context (original XML)')
    const docxIndex = Number(node.attrs.docxIndex)
    const original = app.getBlocks().find((block) => block.docxIndex === docxIndex)
    if (!original?.originalXml) {
      return failure('the imported shape has no original anchor XML; z-order cannot be patched')
    }
    const raw = readShapeRelativeHeight(original.originalXml)
    if (raw === null) {
      return failure('the imported shape has no wp:anchor relativeHeight (inline or VML shape); z-order is unsupported')
    }
    currentRaw = raw
  } else {
    currentRaw = relativeHeightForRank(Number(node.attrs?.imageZOrder ?? 0))
  }
  const ranks = floatingRanks(ctx)
  let targetRaw: number
  if (absoluteRaw !== null) {
    targetRaw = absoluteRaw
  } else if (order === 'front') {
    targetRaw = relativeHeightForRank(ranks.max + 1)
  } else if (order === 'back') {
    targetRaw = Math.max(0, relativeHeightForRank(ranks.min - 1))
  } else if (order === 'forward') {
    targetRaw = Math.min(MAX_RELATIVE_HEIGHT, currentRaw + 1)
  } else {
    targetRaw = Math.max(0, currentRaw - 1)
  }
  const where = `${order}${order === 'absolute' ? ` (relativeHeight ${targetRaw})` : ''}`
  if (ctx.dryRun) {
    return planned(`would move shape block ${blockIndex} to ${where}`, {
      blockIndex,
      imported,
      relativeHeight: targetRaw,
    })
  }
  const patch: Record<string, unknown> = imported
    ? // imported anchors keep their bytes; the save path patches relativeHeight only
      { shapeZOrder: targetRaw }
    : {
        imageZOrder: targetRaw - RELATIVE_HEIGHT_BASE,
        imageWrap:
          node.attrs?.imageWrap === 'front' || node.attrs?.imageWrap === 'behind'
            ? node.attrs.imageWrap
            : 'front',
      }
  const before = ctx.editor.state.doc
  ctx.editor.view.dispatch(
    ctx.editor.state.tr.setNodeMarkup(blockStartPos(doc, blockIndex), undefined, {
      ...node.attrs,
      ...patch,
    }),
  )
  if (ctx.editor.state.doc === before) {
    return failure(`the shape is already at ${where}; no stacking change was needed`)
  }
  return success(
    `Moved the shape at block ${blockIndex} to ${where}${imported ? ' (patched onto the original anchor on save)' : ''}.`,
    'shape z-order',
    { blockIndex, order, imported, relativeHeight: targetRaw },
  )
}

// ── A1: equation block ↔ inline conversion ──

const EQUATION_CONVERT_TARGETS = ['inline', 'block'] as const

function inlineMathNodesAt(doc: PmNode, blockIndex: number): Array<{ pos: number; size: number; latex: string }> {
  const out: Array<{ pos: number; size: number; latex: string }> = []
  let base = 0
  for (let i = 0; i < blockIndex; i++) base += doc.child(i).nodeSize
  doc.child(blockIndex).descendants((node, offset) => {
    if (node.type.name !== 'docInlineMath') return
    // descendants positions are relative to the block CONTENT start
    out.push({ pos: base + 1 + offset, size: node.nodeSize, latex: String(node.attrs?.latex ?? '') })
  })
  return out
}

function equationConvertOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const to = String(input.to ?? '')
  if (!(EQUATION_CONVERT_TARGETS as readonly string[]).includes(to)) {
    return failure('to must be "inline" or "block"')
  }
  const editor = ctx.editor
  const doc = editor.state.doc
  const node = doc.child(blockIndex)
  const blockPos = blockStartPos(doc, blockIndex)

  if (to === 'inline') {
    if (node.type.name !== 'docProtected' || !node.attrs?.formulaDisplay) {
      return failure(`block ${blockIndex} is not a block equation; nothing to convert`)
    }
    const display = node.attrs.formulaDisplay as { latex?: string }
    const latex = String(display.latex ?? '').trim()
    if (!latex) return failure('the block equation has no LaTeX source; reopen/re-edit it first')
    if (ctx.dryRun) {
      return planned(`would convert the block equation at block ${blockIndex} into an inline equation in an adjacent paragraph`)
    }
    const inlineJson = inlineEquationNodeJson(latex)
    const inlineNode = editor.schema.nodeFromJSON(inlineJson)
    const prev = blockIndex > 0 ? doc.child(blockIndex - 1) : null
    const next = blockIndex + 1 < doc.childCount ? doc.child(blockIndex + 1) : null
    let tr = editor.state.tr
    const textTypes = new Set(['docParagraph', 'docHeading', 'docListItem'])
    if (prev && textTypes.has(prev.type.name as string)) {
      // append inside the paragraph above, then drop the equation block
      tr = tr.delete(blockPos, blockPos + node.nodeSize)
      tr = tr.insert(blockPos - 1, inlineNode)
    } else if (next && textTypes.has(next.type.name as string)) {
      tr = tr.delete(blockPos, blockPos + node.nodeSize)
      tr = tr.insert(blockPos + 1, inlineNode)
    } else {
      const paragraph = editor.schema.nodeFromJSON({
        type: 'docParagraph',
        attrs: { docxIndex: null },
        content: [inlineJson],
      })
      tr = tr.replaceWith(blockPos, blockPos + node.nodeSize, paragraph)
    }
    editor.view.dispatch(tr)
    return success(
      `Converted the block equation at block ${blockIndex} into an inline equation in an adjacent paragraph.`,
      'convert equation',
      { blockIndex, to: 'inline' },
    )
  }

  // to 'block': an inline equation in this block becomes its own equation block
  const inlineIndex = input.inlineIndex === undefined ? null : intOf(input.inlineIndex)
  if (inlineIndex === null || inlineIndex < 0) {
    return failure('converting to a block equation needs inlineIndex (0-based inline equation in the block)')
  }
  const inlineMath = inlineMathNodesAt(doc, blockIndex)
  if (inlineMath.length === 0) {
    return failure(`block ${blockIndex} contains no inline equation; nothing to convert`)
  }
  if (inlineIndex >= inlineMath.length) {
    return failure(`inlineIndex ${inlineIndex} is out of range; block ${blockIndex} has ${inlineMath.length} inline equation(s)`)
  }
  const target = inlineMath[inlineIndex]
  const latex = target.latex.trim()
  if (!latex) return failure('the inline equation has no LaTeX source; reopen/re-edit it first')
  if (ctx.dryRun) {
    return planned(`would convert inline equation ${inlineIndex} in block ${blockIndex} into its own block equation`)
  }
  const blockJson = equationBlockJson(latex)
  const blockNode = editor.schema.nodeFromJSON(blockJson)
  let tr = editor.state.tr
  // place the new block right after the containing block, then remove the inline node
  const afterBlock = blockPos + node.nodeSize
  tr = tr.insert(afterBlock, blockNode)
  tr = tr.delete(target.pos, target.pos + target.size)
  editor.view.dispatch(tr)
  return success(
    `Converted inline equation ${inlineIndex} in block ${blockIndex} into its own block equation.`,
    'convert equation',
    { blockIndex, to: 'block', inlineIndex },
  )
}

// After A1, replace/crop/local background removal are implemented. This op now
// only enumerates the genuinely unsupported image capability so the model can
// explain it instead of faking a change.
const UNSUPPORTED_IMAGE_FEATURES = ['autoBackgroundRemoval'] as const

function imageUnsupportedOp(
  input: Record<string, unknown>,
  _ctx: OperationContext,
): OperationOutcome {
  const feature = String(input.feature ?? '')
  if (!(UNSUPPORTED_IMAGE_FEATURES as readonly string[]).includes(feature)) {
    return failure(`feature must be one of ${UNSUPPORTED_IMAGE_FEATURES.join('/')}`)
  }
  return failure(
    'automatic (AI/model-based) subject segmentation is unsupported; the editor only offers the local '
    + 'tolerance cutout — use object.image.removeBackground { blockIndex, tolerance? } instead. '
    + 'Replacing bytes is object.image.replace; cropping is object.image.crop.',
  )
}

export const objectOperations: OperationDef[] = [
  {
    name: 'object.textbox.insert',
    group: 'object',
    description:
      'object.textbox.insert { text?, afterBlockIndex? } — insert a floating text box (real ribbon insertTextboxAt) '
      + 'with optional plain text; afterBlockIndex omitted = at the caret',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: { type: 'integer', minimum: 0, description: 'insert after this block index' },
        text: { type: 'string', maxLength: MAX_TEXT, description: 'plain text inside the box' },
      },
    },
    handler: textboxInsertOp,
  },
  {
    name: 'object.wordArt.insert',
    group: 'object',
    description:
      'object.wordArt.insert { preset?, text?, afterBlockIndex? } — insert a WordArt object from the gallery '
      + `(presets: ${WORDART_PRESETS.map((preset) => preset.id).join('/')}) with optional display text`,
    inputSchema: {
      type: 'object',
      properties: {
        preset: {
          type: 'string',
          enum: WORDART_PRESETS.map((preset) => preset.id),
          description: 'gallery preset id; omitted = the first preset',
        },
        text: { type: 'string', maxLength: MAX_TEXT },
        afterBlockIndex: { type: 'integer', minimum: 0 },
      },
    },
    handler: wordArtInsertOp,
  },
  {
    name: 'object.dropCap',
    group: 'object',
    description:
      'object.dropCap { target, mode } — set/remove the paragraph drop cap (mode: none/dropped/margin)',
    inputSchema: {
      type: 'object',
      properties: {
        target: {
          type: 'object',
          description: 'apply_commands target vocabulary',
          properties: {
            nodeType: { type: 'string', enum: ['docHeading', 'docParagraph', 'docListItem', 'image'] },
            headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
            containsText: { type: 'string' },
            matchCase: { type: 'boolean' },
            blockIndexes: { type: 'array', items: { type: 'integer', minimum: 0 } },
            scope: { type: 'string', enum: ['selection', 'document'] },
          },
        },
        mode: { type: 'string', enum: ['none', 'dropped', 'margin'] },
      },
      required: ['target', 'mode'],
    },
    handler: dropCapOp,
  },
  {
    name: 'object.equation.insert',
    group: 'object',
    description:
      'object.equation.insert { latex, inline?, afterBlockIndex? } — insert a native Word equation from LaTeX '
      + '(block equation by default, inline:true flows with the text); invalid LaTeX is rejected',
    inputSchema: {
      type: 'object',
      properties: {
        latex: { type: 'string', minLength: 1, maxLength: 10000 },
        inline: { type: 'boolean', description: 'true = inline equation at the caret' },
        afterBlockIndex: { type: 'integer', minimum: 0 },
      },
      required: ['latex'],
    },
    handler: equationInsertOp,
  },
  {
    name: 'object.equation.update',
    group: 'object',
    description:
      'object.equation.update { blockIndex, latex, inlineIndex? } — re-edit an EXISTING equation in place (block '
      + 'equation by default; pass inlineIndex for the Nth inline equation inside the block); never inserts a new one',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0, description: 'block containing the equation' },
        latex: { type: 'string', minLength: 1, maxLength: 10000 },
        inlineIndex: {
          type: 'integer',
          minimum: 0,
          description: '0-based inline equation index inside the block; omit for a block equation',
        },
      },
      required: ['blockIndex', 'latex'],
    },
    handler: equationUpdateOp,
  },
  {
    name: 'object.symbol.insert',
    group: 'object',
    description:
      'object.symbol.insert { char, afterBlockIndex?, ownParagraph? } — insert a symbol (e.g. "©", "→") at the end of '
      + 'a block / at the caret, or in its own paragraph when ownParagraph is true',
    inputSchema: {
      type: 'object',
      properties: {
        char: { type: 'string', minLength: 1, maxLength: 16 },
        afterBlockIndex: { type: 'integer', minimum: 0 },
        ownParagraph: { type: 'boolean', description: 'true = place the symbol in a new paragraph' },
      },
      required: ['char'],
    },
    handler: symbolInsertOp,
  },
  {
    name: 'object.image.format',
    group: 'object',
    description:
      'object.image.format { blockIndex, widthPx?, heightPx?, align?, wrap?, position?, rotationDeg?, flipH?, flipV?, zOrder? } '
      + '— format an existing picture block: size/alignment, text wrap (square-left/square-right/topBottom/behind/front/null=inline), '
      + 'position preset, rotation, mirror, stacking order',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0, description: 'index of the image block' },
        widthPx: { type: 'number', minimum: 16, maximum: 5000 },
        heightPx: { type: 'number', minimum: 16, maximum: 5000 },
        align: { type: ['string', 'null'], enum: ['left', 'center', 'right', null] },
        wrap: { type: ['string', 'null'], enum: [...WRAP_VALUES, null] },
        position: {
          type: 'object',
          properties: {
            horizontal: { type: 'string', enum: POSITION_H },
            vertical: { type: 'string', enum: POSITION_V },
          },
          required: ['horizontal', 'vertical'],
        },
        rotationDeg: { type: 'number', minimum: -3600, maximum: 3600 },
        flipH: { type: 'boolean' },
        flipV: { type: 'boolean' },
        zOrder: { type: 'integer', minimum: -1000, maximum: 1000 },
      },
      required: ['blockIndex'],
    },
    handler: imageFormatOp,
  },
  {
    name: 'object.shape.format',
    group: 'object',
    description:
      'object.shape.format { blockIndex, fillHex?, lineHex?, text?, widthEmu?, heightEmu?, zOrder? } — format an existing '
      + 'shape/text box/WordArt block: fill and outline colors, text (replaces the box text), size (EMU), stacking order '
      + '(zOrder works on editor-created shapes only)',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0, description: 'index of the shape/textbox block' },
        fillHex: { type: ['string', 'null'], description: '6-digit hex without #, null = no fill' },
        lineHex: { type: ['string', 'null'], description: '6-digit hex without #, null = no outline' },
        text: { type: ['string', 'null'], maxLength: MAX_TEXT },
        widthEmu: { type: 'integer', minimum: 91440, maximum: 7772400 },
        heightEmu: { type: 'integer', minimum: 91440, maximum: 7772400 },
        zOrder: { type: 'integer', minimum: -1000, maximum: 1000 },
      },
      required: ['blockIndex'],
    },
    handler: shapeFormatOp,
  },
  {
    name: 'object.image.unsupported',
    group: 'object',
    description:
      'object.image.unsupported { feature } — ALWAYS FAILS by design: enumerates the remaining unsupported image '
      + 'capability (autoBackgroundRemoval = AI subject segmentation; replace/crop/local cutout have dedicated ops).',
    inputSchema: {
      type: 'object',
      properties: {
        feature: { type: 'string', enum: UNSUPPORTED_IMAGE_FEATURES },
      },
      required: ['feature'],
    },
    handler: imageUnsupportedOp,
  },
  {
    name: 'object.image.replace',
    group: 'object',
    description:
      'object.image.replace { blockIndex, base64?, mime?, dataUrl? } — replace a picture block\'s bytes. The gateway '
      + 'resolves path/url to base64+mime and forwards {base64, mime}; a ready dataUrl is also accepted. Imported '
      + 'pictures swap bytes in place (drawing XML preserved), editor-created ones update their pending image; '
      + 'display width is kept, height follows the new aspect ratio',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0, description: 'index of the image block' },
        base64: { type: 'string', description: 'standard base64 bytes (no data: prefix); gateway-resolved' },
        mime: { type: 'string', enum: ['image/png', 'image/jpeg', 'image/gif'] },
        dataUrl: { type: 'string', description: 'alternative to base64+mime: a data:image/...;base64,… URL' },
      },
      required: ['blockIndex'],
    },
    handler: imageReplaceOp,
  },
  {
    name: 'object.image.crop',
    group: 'object',
    description:
      'object.image.crop { blockIndex, rect: { x, y, w, h }, unit?: px|fraction } — destructively crop a picture '
      + 'block (bake the crop into new bytes, keeping the display width). rect is against the image NATURAL size: '
      + 'natural PIXELS by default, or 0..1 FRACTIONS with unit:"fraction". documentState.images exposes DISPLAY '
      + 'widthPx/heightPx only — for relative requests ("middle half", "top third") prefer unit:"fraction"',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0 },
        rect: {
          type: 'object',
          properties: {
            x: { type: 'number' },
            y: { type: 'number' },
            w: { type: 'number' },
            h: { type: 'number' },
          },
          required: ['x', 'y', 'w', 'h'],
        },
        unit: {
          type: 'string',
          enum: ['px', 'fraction'],
          description: 'px = natural image pixels (default); fraction = 0..1 of the natural size',
        },
      },
      required: ['blockIndex', 'rect'],
    },
    handler: imageCropOp,
  },
  {
    name: 'object.image.removeBackground',
    group: 'object',
    description:
      'object.image.removeBackground { blockIndex, tolerance?, sample? } — LOCAL tolerance cutout (edge flood fill from '
      + 'the border color; no AI/external service): tolerance 0..100 (default 25), sample "corners" = border sampling; '
      + 'bakes transparent alpha into new PNG bytes',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0 },
        tolerance: { type: 'integer', minimum: 0, maximum: 100 },
        sample: { type: 'string', enum: ['corners'] },
      },
      required: ['blockIndex'],
    },
    handler: imageRemoveBackgroundOp,
  },
  {
    name: 'object.shape.zOrder',
    group: 'object',
    description:
      'object.shape.zOrder { blockIndex, order: front|back|forward|backward|absolute, relativeHeight? } — change the '
      + 'stacking order of a shape/text box: editor-created shapes update their anchor rank, imported shapes record a '
      + 'raw wp:anchor relativeHeight that the save path patches onto the original XML (order "absolute" needs '
      + 'relativeHeight, the raw Word value 251658240 + rank)',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0, description: 'index of the shape/text-box block' },
        order: { type: 'string', enum: SHAPE_Z_ORDERS },
        relativeHeight: {
          type: 'integer',
          minimum: 0,
          maximum: MAX_RELATIVE_HEIGHT,
          description: 'only for order "absolute": the raw wp:anchor relativeHeight',
        },
      },
      required: ['blockIndex', 'order'],
    },
    handler: shapeZOrderOp,
  },
  {
    name: 'object.equation.convert',
    group: 'object',
    description:
      'object.equation.convert { blockIndex, to: inline|block, inlineIndex? } — convert an existing BLOCK equation into '
      + 'an inline equation in an adjacent paragraph (to:"inline"), or the Nth inline equation inside a block into its '
      + 'own block equation (to:"block", needs inlineIndex); never duplicates the equation',
    inputSchema: {
      type: 'object',
      properties: {
        blockIndex: { type: 'integer', minimum: 0 },
        to: { type: 'string', enum: EQUATION_CONVERT_TARGETS },
        inlineIndex: { type: 'integer', minimum: 0, description: 'required when to:"block"' },
      },
      required: ['blockIndex', 'to'],
    },
    handler: equationConvertOp,
  },
]
