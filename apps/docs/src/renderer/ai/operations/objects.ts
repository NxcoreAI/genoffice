import type { Editor } from '@tiptap/core'
import { WORDART_PRESETS, type WordArtPreset } from '@genoffice/ui'
import { latexToOmml, type TextboxDisplay } from '@genoffice/docx-engine'
import { equationBlockJson, inlineEquationNodeJson } from '../../editor/equation'
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
    node.descendants((child, offset) => {
      if (child.type.name === 'docInlineMath') {
        inlineMath.push({ pos: blockPos + offset, size: child.nodeSize })
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

const UNSUPPORTED_IMAGE_FEATURES = ['replace', 'crop', 'removeBackground'] as const

function imageUnsupportedOp(
  input: Record<string, unknown>,
  _ctx: OperationContext,
): OperationOutcome {
  const feature = String(input.feature ?? '')
  if (!(UNSUPPORTED_IMAGE_FEATURES as readonly string[]).includes(feature)) {
    return failure(`feature must be one of ${UNSUPPORTED_IMAGE_FEATURES.join('/')}`)
  }
  return failure(
    `object.image.${feature} is unsupported: it needs the host to supply/process the picture bytes `
    + '(replacement image data, crop geometry applied to the file, or background removal). '
    + 'Insert a host-provided image with insert_image_base64, or re-insert the prepared image and '
    + 'remove the old block with apply_commands deleteBlocks.',
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
      'object.image.unsupported { feature } — ALWAYS FAILS by design: replace/crop/removeBackground need host-supplied '
      + 'picture bytes. Documented here so the model can explain the limitation instead of faking a change.',
    inputSchema: {
      type: 'object',
      properties: {
        feature: { type: 'string', enum: UNSUPPORTED_IMAGE_FEATURES },
      },
      required: ['feature'],
    },
    handler: imageUnsupportedOp,
  },
]
