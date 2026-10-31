import { flushSync } from 'react-dom'
import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import type { CustomNumberingLevel, TabStop } from '@genoffice/docx-engine'
import {
  resolveTargetBlocks,
  validateCommandTarget,
  type Target,
  type TargetBlock,
} from '../commands'
import { createCustomListDef, continueNumbering, restartNumbering } from '../../numbering-actions'
import { clearParagraphFormatting, applyParagraphStyle } from '../../components/ribbon-tabs'
import { applyCase, type CaseMode } from '../../editor/case-transform'
import { setParagraphDirection } from '../../editor/direction'
import { stepHangingIndent } from '../../editor/indent'
import {
  MULTILEVEL_LIBRARY,
  MULTILEVEL_PRESETS,
  multilevelLevelsFor,
} from '../../editor/multilevel-presets'
import {
  applyCharStyleToSelection,
  charStylePresetById,
  presetAccentOf,
} from '../../editor/character-styles'
import type { OperationContext, OperationDef, OperationOutcome } from './types'
import {
  blockEndPos,
  checkBlockIndex,
  failure,
  intOf,
  planned,
  selectBlocks,
  success,
  textBlocksOf,
} from './shared'

/**
 * B3 inline/paragraph formatting ops. They reuse the exact apply_commands
 * target vocabulary (validateCommandTarget/resolveTargetBlocks) and the real
 * ribbon/editor helpers (applyCase, setParagraphDirection, stepHangingIndent,
 * applyParagraphStyle, clearParagraphFormatting, setParaAttrs, numbering
 * actions), so an op behaves like the equivalent ribbon action.
 */

const CASE_MODES: CaseMode[] = ['upper', 'lower', 'title', 'sentence']
const STYLE_KEYS = ['p', 'h1', 'h2', 'h3'] as const

function resolveTargets(
  ctx: OperationContext,
  target: unknown,
): { blocks: TargetBlock[] } | { error: string } {
  const targetError = validateCommandTarget(target)
  if (targetError) return { error: targetError }
  const blocks = resolveTargetBlocks(
    ctx.editor.state.doc,
    target as Target,
    ctx.editor.state.selection,
  )
  if (blocks.length === 0) {
    return {
      error:
        'target matched no blocks; use blockIndexes / nodeType / containsText from the block list '
        + '(or scope:"selection" with a selection in the document)',
    }
  }
  return { blocks }
}

function resolveTextTargets(
  ctx: OperationContext,
  target: unknown,
): { blocks: TargetBlock[]; skippedProtected: number } | { error: string } {
  const resolved = resolveTargets(ctx, target)
  if ('error' in resolved) return resolved
  const text = textBlocksOf(resolved.blocks)
  if (text.length === 0) {
    return {
      error:
        'target matched no text blocks (paragraph/heading/list item); '
        + 'protected blocks such as images/tables cannot receive inline formatting',
    }
  }
  return { blocks: text, skippedProtected: resolved.blocks.length - text.length }
}

function selectTargets(
  editor: Editor,
  blocks: TargetBlock[],
): { from: number; to: number } | { error: string } {
  const range = selectBlocks(editor, blocks)
  if (!range) return { error: 'target blocks have no selectable text range' }
  return range
}

/** Paragraph-like target blocks with the cursor/selection placed on them. */
function prepareParagraphTargets(
  ctx: OperationContext,
  target: unknown,
): { blocks: TargetBlock[]; skippedProtected: number } | { error: string } {
  const resolved = resolveTextTargets(ctx, target)
  if ('error' in resolved) return resolved
  const selected = selectTargets(ctx.editor, resolved.blocks)
  if ('error' in selected) return selected
  return resolved
}

function listTargetBlocks(ctx: OperationContext, target: unknown): TargetBlock[] | { error: string } {
  const resolved = resolveTargets(ctx, target)
  if ('error' in resolved) return resolved
  const items = resolved.blocks.filter((block) => block.node.type.name === 'docListItem')
  if (items.length === 0) {
    return {
      error:
        'target matched no list items; format.listRestart/format.listContinue need a docListItem '
        + '(create one first with apply_commands createParagraphBullets or format.multilevelList)',
    }
  }
  return items
}

function changeCaseOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const mode = String(input.mode ?? '') as CaseMode
  if (!CASE_MODES.includes(mode)) {
    return failure(`mode must be one of ${CASE_MODES.join('/')}`)
  }
  const resolved = resolveTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const text = textBlocksOf(resolved.blocks)
  if (text.length === 0) {
    return failure('target matched no text blocks (the new case would have nothing to change)')
  }
  if (ctx.dryRun) {
    return planned(`would change the case of ${text.length} block(s) to "${mode}"`)
  }
  const selected = selectTargets(ctx.editor, text)
  if ('error' in selected) return failure(selected.error)
  const before = ctx.editor.state.doc
  const ran = applyCase(ctx.editor, mode)
  if (!ran) return failure('case change failed: the target text is empty or the range is invalid')
  const changed = ctx.editor.state.doc !== before
  return success(
    changed
      ? `Changed the case of ${text.length} block(s) to "${mode}".`
      : `Text in ${text.length} block(s) was already in "${mode}" case.`,
    'change case',
    { blocks: text.length, mode, changed },
  )
}

function clearFormattingOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const resolved = resolveTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const text = textBlocksOf(resolved.blocks)
  if (text.length === 0) return failure('target matched no text blocks to clear')
  if (ctx.dryRun) {
    return planned(
      `would clear direct character and paragraph formatting on ${text.length} block(s)`,
    )
  }
  const selected = selectTargets(ctx.editor, text)
  if ('error' in selected) return failure(selected.error)
  const before = ctx.editor.state.doc
  // Word's Ctrl+Space (character formatting) + Ctrl+Q (paragraph formatting)
  ctx.editor.chain().focus().unsetAllMarks().run()
  clearParagraphFormatting(ctx.editor)
  const changed = ctx.editor.state.doc !== before
  return success(
    changed
      ? `Cleared direct formatting on ${text.length} block(s).`
      : `Direct formatting on ${text.length} block(s) was already clear.`,
    'clear formatting',
    { blocks: text.length, changed },
  )
}

function setBidiOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const direction = String(input.direction ?? '')
  if (direction !== 'ltr' && direction !== 'rtl') {
    return failure('direction must be "ltr" or "rtl"')
  }
  const resolved = prepareParagraphTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const before = ctx.editor.state.doc
  if (ctx.dryRun) {
    return planned(
      `would set ${direction} writing direction on ${resolved.blocks.length} paragraph-like block(s)`,
    )
  }
  const changed = setParagraphDirection(ctx.editor, direction)
  return success(
    ctx.editor.state.doc !== before
      ? `Set ${direction} direction on ${resolved.blocks.length} block(s).`
      : `Direction of the target block(s) was already ${direction}.`,
    'set text direction',
    { blocks: resolved.blocks.length, direction, changed },
  )
}

function hangingIndentOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  let steps = 1
  if (input.steps !== undefined) {
    const parsed = intOf(input.steps)
    if (parsed === null || parsed === 0 || parsed < -10 || parsed > 10) {
      return failure('steps must be a non-zero integer between -10 and 10 (positive = increase hanging indent)')
    }
    steps = parsed
  }
  const resolved = prepareParagraphTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  if (ctx.dryRun) {
    return planned(
      `would ${steps > 0 ? 'increase' : 'decrease'} the hanging indent by ${Math.abs(steps)} `
      + `step(s) on ${resolved.blocks.length} block(s)`,
    )
  }
  const before = ctx.editor.state.doc
  let changed = false
  for (let i = 0; i < Math.abs(steps); i++) {
    if (stepHangingIndent(ctx.editor, steps > 0 ? 1 : -1)) changed = true
  }
  return success(
    changed || ctx.editor.state.doc !== before
      ? `${steps > 0 ? 'Increased' : 'Decreased'} the hanging indent by ${Math.abs(steps)} step(s) `
        + `on ${resolved.blocks.length} block(s).`
      : 'The hanging indent is already at its limit; no change.',
    'hanging indent',
    { blocks: resolved.blocks.length, steps },
  )
}

function applyStyleOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const style = String(input.style ?? '')
  if (!(STYLE_KEYS as readonly string[]).includes(style)) {
    return failure(`style must be one of ${STYLE_KEYS.join('/')}`)
  }
  const resolved = prepareParagraphTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  if (ctx.dryRun) {
    return planned(
      `would apply the "${style}" paragraph style to ${resolved.blocks.length} block(s)`,
    )
  }
  const before = ctx.editor.state.doc
  applyParagraphStyle(ctx.editor, style as (typeof STYLE_KEYS)[number])
  return success(
    ctx.editor.state.doc !== before
      ? `Applied the "${style}" paragraph style to ${resolved.blocks.length} block(s).`
      : `The target block(s) already use the "${style}" style.`,
    'apply paragraph style',
    { blocks: resolved.blocks.length, style },
  )
}

/** Put the caret inside a target list item (numbering actions read the cursor's list). */
function placeCaretInBlock(editor: Editor, block: TargetBlock): void {
  const position = Math.min(block.pos + 1, editor.state.doc.content.size)
  const selection = TextSelection.near(editor.state.doc.resolve(position), 1)
  editor.view.dispatch(editor.state.tr.setSelection(selection))
}

function listRestartOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const items = listTargetBlocks(ctx, input.target)
  if ('error' in items) return failure(items.error)
  const first = items[0]
  const numId = first.node.attrs.numId
  if (numId === null || numId === undefined) {
    return failure('the target list item has no numbering definition to restart (numId is unset)')
  }
  const app = ctx.app
  if (!app) return failure('format.listRestart needs the app context (host call path)')
  if (ctx.dryRun) {
    return planned(`would restart numbering at the target list (numId ${String(numId)})`)
  }
  placeCaretInBlock(ctx.editor, first)
  const before = ctx.editor.state.doc
  // pending numbering definitions must be flushed before the batch save reads them
  flushSync(() => restartNumbering(app.numbering()))
  if (ctx.editor.state.doc === before) {
    return failure('restart produced no change: no usable numbering definition was found')
  }
  return success('Restarted numbering at the target list.', 'restart list numbering', {
    numId: String(numId),
  })
}

function listContinueOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const items = listTargetBlocks(ctx, input.target)
  if ('error' in items) return failure(items.error)
  const app = ctx.app
  if (!app) return failure('format.listContinue needs the app context (host call path)')
  if (ctx.dryRun) {
    return planned('would continue numbering from the previous list')
  }
  placeCaretInBlock(ctx.editor, items[0])
  const before = ctx.editor.state.doc
  flushSync(() => continueNumbering(app.numbering()))
  if (ctx.editor.state.doc === before) {
    return failure(
      'continue produced no change: the target list is not preceded by another numbered list',
    )
  }
  return success('Continued numbering from the previous list.', 'continue list numbering')
}

function multilevelListOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const preset = multilevelLevelsFor(input.preset)
  if ('error' in preset) return failure(preset.error)
  const resolved = resolveTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const eligible = resolved.blocks.filter(
    (block) => block.node.type.name === 'docParagraph' || block.node.type.name === 'docListItem',
  )
  if (eligible.length === 0) {
    return failure(
      'target matched no paragraphs or list items (headings and protected blocks are not converted)',
    )
  }
  const skipped = resolved.blocks.length - eligible.length
  const app = ctx.app
  if (!app) return failure('format.multilevelList needs the app context (host call path)')
  if (ctx.dryRun) {
    return planned(
      `would apply the "${preset.id}" multilevel numbering to ${eligible.length} block(s)`
      + (skipped > 0 ? ` (${skipped} skipped)` : ''),
    )
  }
  // create the numbering definition through the real numbering-actions path
  let numId: string | null = null
  flushSync(() => {
    numId = createCustomListDef(app.numbering(), preset.levels)
  })
  if (!numId) {
    return failure('could not create the multilevel numbering definition (no document?)')
  }
  const levels: CustomNumberingLevel[] = preset.levels
  const kind = levels[0]?.numFmt === 'bullet' ? ('bullet' as const) : ('ordered' as const)
  const editor = ctx.editor
  let tr = editor.state.tr
  for (const block of eligible) {
    if (block.node.type.name === 'docListItem') {
      tr = tr.setNodeMarkup(block.pos, undefined, { ...block.node.attrs, kind, numId })
    } else {
      tr = tr.setNodeMarkup(block.pos, editor.schema.nodes.docListItem, {
        ...block.node.attrs,
        styleId: null,
        kind,
        numId,
        ilvl: 0,
        aiChanged: true,
      })
    }
  }
  editor.view.dispatch(tr)
  return success(
    `Applied the "${preset.id}" multilevel numbering (numId ${numId}) to ${eligible.length} block(s)`
    + (skipped > 0 ? `; ${skipped} non-paragraph block(s) skipped.` : '.'),
    'apply multilevel list',
    { blocks: eligible.length, skipped, numId: String(numId), preset: preset.id },
  )
}

// ── B5: character styles, tab stops, hard breaks, list styles ──

const CHAR_STYLE_MODES = ['apply', 'toggle'] as const
const BREAK_KINDS = ['line', 'column'] as const
const BREAK_ATS = ['start', 'end'] as const
const TAB_VALS: TabStop['val'][] = ['left', 'center', 'right', 'decimal', 'bar']
const TAB_LEADERS = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'] as const

function applyCharacterStyleOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): OperationOutcome {
  const hasPreset = input.preset !== undefined
  const hasStyleId = input.styleId !== undefined
  if (hasPreset === hasStyleId) return failure('pass exactly one of preset or styleId')
  const modeRaw = input.mode === undefined ? 'toggle' : String(input.mode)
  if (!(CHAR_STYLE_MODES as readonly string[]).includes(modeRaw)) {
    return failure('mode must be "apply" or "toggle"')
  }
  const mode = modeRaw as (typeof CHAR_STYLE_MODES)[number]
  let preset = null as ReturnType<typeof charStylePresetById> | null
  let styleId: string | undefined
  if (hasPreset) {
    preset = charStylePresetById(String(input.preset))
    if (!preset) {
      return failure('preset must be "emphasis" or "strong" (or pass a document styleId)')
    }
  } else {
    styleId = String(input.styleId ?? '').trim()
    if (!styleId) return failure('styleId must be a non-empty character style id')
    if (styleId.length > 120) return failure('styleId must be at most 120 characters')
  }
  const resolved = resolveTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const text = textBlocksOf(resolved.blocks)
  if (text.length === 0) return failure('target matched no text blocks')
  const selected = selectBlocks(ctx.editor, text)
  if (!selected) return failure('target blocks have no selectable text range')
  const selectedText = ctx.editor.state.doc.textBetween(selected.from, selected.to, '\n', '\n').trim()
  if (!selectedText) return failure('the target blocks contain no text to style')
  const accent = presetAccentOf(ctx.app?.themeColors)
  const label = preset ? preset.id : `styleId "${styleId}"`
  if (ctx.dryRun) {
    return planned(
      `would ${mode === 'toggle' ? 'toggle' : 'apply'} character style ${label} on ${text.length} block(s)`,
    )
  }
  const before = ctx.editor.state.doc
  const ran = applyCharStyleToSelection(
    ctx.editor,
    preset ? { preset, accent } : { styleId, accent },
    mode,
  )
  if (!ran) return failure('character style could not be applied (no selection?)')
  return success(
    ctx.editor.state.doc !== before
      ? `${mode === 'toggle' ? 'Toggled' : 'Applied'} character style ${label} on ${text.length} block(s).`
      : `Character style ${label} already matched the target.`,
    'apply character style',
    { blocks: text.length, ...(preset ? { preset: preset.id } : { styleId }), mode },
  )
}

function setTabStopsOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const stopsRaw = input.stops
  if (!Array.isArray(stopsRaw)) return failure('stops must be an array of tab-stop objects (use [] to clear)')
  if (stopsRaw.length > 32) return failure('stops must contain at most 32 entries')
  const stops: TabStop[] = []
  const seen = new Set<number>()
  for (const raw of stopsRaw) {
    if (!raw || typeof raw !== 'object') return failure('each stop must be { posTwips, val, leader? }')
    const entry = raw as Record<string, unknown>
    const pos = intOf(entry.posTwips)
    if (pos === null || pos < 0 || pos > 31680) return failure('stop.posTwips must be an integer 0..31680')
    if (seen.has(pos)) return failure(`duplicate tab stop position ${pos}`)
    seen.add(pos)
    const val = String(entry.val ?? '')
    if (!(TAB_VALS as readonly string[]).includes(val)) {
      return failure(`stop.val must be one of ${TAB_VALS.join('/')}`)
    }
    let leader: TabStop['leader']
    if (entry.leader !== undefined) {
      const leaderRaw = String(entry.leader)
      if (!(TAB_LEADERS as readonly string[]).includes(leaderRaw)) {
        return failure(`stop.leader must be one of ${TAB_LEADERS.join('/')}`)
      }
      leader = leaderRaw as TabStop['leader']
    }
    stops.push({ pos, val: val as TabStop['val'], ...(leader ? { leader } : {}) })
  }
  stops.sort((a, b) => a.pos - b.pos)
  const resolved = resolveTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const text = textBlocksOf(resolved.blocks)
  if (text.length === 0) return failure('target matched no paragraph/heading/list-item blocks')
  const where =
    stops.length === 0
      ? `clear tab stops on ${text.length} block(s)`
      : `set ${stops.length} tab stop(s) on ${text.length} block(s)`
  if (ctx.dryRun) return planned(`would ${where}`)
  // preserve display-only rel (w:ptab) stops: dropping them loses paragraph layout
  let tr = ctx.editor.state.tr
  for (const block of text) {
    const rawStops = block.node.attrs.tabStops as string | null
    let relStops: TabStop[] = []
    if (rawStops) {
      try {
        const parsed = JSON.parse(rawStops)
        if (Array.isArray(parsed)) relStops = parsed.filter((stop) => stop && stop.rel)
      } catch {
        /* malformed attr: replace it entirely */
      }
    }
    const next = stops.length > 0 || relStops.length > 0 ? [...stops, ...relStops] : null
    tr = tr.setNodeMarkup(block.pos, undefined, {
      ...block.node.attrs,
      tabStops: next ? JSON.stringify(next) : null,
    })
  }
  ctx.editor.view.dispatch(tr)
  return success(`${where[0].toUpperCase()}${where.slice(1)}.`, 'set tab stops', {
    blocks: text.length,
    stops: stops.length,
  })
}

function insertBreakOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const kind = String(input.kind ?? '')
  if (!(BREAK_KINDS as readonly string[]).includes(kind)) {
    return failure('kind must be "line" (Shift+Enter) or "column" (column break)')
  }
  const atRaw = input.at === undefined ? 'end' : String(input.at)
  if (!(BREAK_ATS as readonly string[]).includes(atRaw)) {
    return failure('at must be "start" or "end"')
  }
  const blockIndex = checkBlockIndex(ctx.editor, input.blockIndex)
  if (typeof blockIndex === 'string') return failure(blockIndex)
  const doc = ctx.editor.state.doc
  const node = doc.child(blockIndex)
  const blockPos = blockEndPos(doc, blockIndex) - node.nodeSize
  const pos = atRaw === 'start' ? blockPos + 1 : blockPos + node.nodeSize - 1
  const where = `${kind} break at the ${atRaw} of block ${blockIndex}`
  if (ctx.dryRun) return planned(`would insert a ${where}`)
  ctx.editor
    .chain()
    .insertContentAt(pos, {
      type: 'hardBreak',
      attrs: kind === 'column' ? { colBreak: true } : {},
    })
    .run()
  return success(`Inserted a ${where}.`, 'insert break', { kind, blockIndex, at: atRaw })
}

function applyListStyleOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const kind = String(input.kind ?? '')
  if (kind !== 'bullet' && kind !== 'ordered') return failure('kind must be "bullet" or "ordered"')
  let levels: CustomNumberingLevel[]
  let presetLabel: string
  if (input.levels === undefined || input.levels === null) {
    // built-in fallback: the gallery's multilevel presets (ordered = 1./1.1./1.1.1., bullet = •/○/■)
    levels = kind === 'bullet' ? MULTILEVEL_LIBRARY[2] : MULTILEVEL_LIBRARY[0]
    presetLabel =
      kind === 'bullet' ? 'bullet fallback (•/○/■)' : 'decimal fallback (1. / 1.1. / 1.1.1.)'
  } else {
    if (!Array.isArray(input.levels) || input.levels.length === 0 || input.levels.length > 9) {
      return failure('levels must be an array of 1..9 level definitions')
    }
    const parsed = input.levels.map((raw, index) => {
      if (!raw || typeof raw !== 'object') return null
      const entry = raw as Record<string, unknown>
      const level = intOf(entry.level)
      if (level === null || level !== index) return null
      const numFmt =
        entry.numFmt === undefined ? (kind === 'bullet' ? 'bullet' : 'decimal') : String(entry.numFmt)
      const lvlText =
        entry.lvlText === undefined
          ? kind === 'bullet'
            ? '•'
            : `%${index + 1}.`
          : String(entry.lvlText)
      const start = entry.start === undefined ? 1 : intOf(entry.start)
      if (start === null || start < 1 || start > 1000) return null
      return {
        numFmt,
        lvlText,
        start,
        indentLeft: 720 * (index + 1),
        hanging: 360,
      } satisfies CustomNumberingLevel
    })
    if (parsed.some((entry) => entry === null)) {
      return failure('levels must be contiguous from level 0, e.g. [{level:0,...},{level:1,...}]')
    }
    levels = parsed as CustomNumberingLevel[]
    presetLabel = 'custom levels'
  }
  const resolved = resolveTargets(ctx, input.target)
  if ('error' in resolved) return failure(resolved.error)
  const eligible = resolved.blocks.filter(
    (block) => block.node.type.name === 'docParagraph' || block.node.type.name === 'docListItem',
  )
  if (eligible.length === 0) {
    return failure('target matched no paragraphs/list items (headings and protected blocks are not converted)')
  }
  const skipped = resolved.blocks.length - eligible.length
  const app = ctx.app
  if (!app) return failure('format.applyListStyle needs the app context (numbering definitions)')
  if (ctx.dryRun) {
    return planned(
      `would apply ${presetLabel} as a ${kind} list to ${eligible.length} block(s)`
        + (skipped > 0 ? ` (${skipped} skipped)` : ''),
    )
  }
  let numId: string | null = null
  flushSync(() => {
    numId = createCustomListDef(app.numbering(), levels)
  })
  if (!numId) return failure('could not create the list numbering definition (no document?)')
  const listKind = levels[0]?.numFmt === 'bullet' ? ('bullet' as const) : ('ordered' as const)
  const editor = ctx.editor
  let tr = editor.state.tr
  for (const block of eligible) {
    if (block.node.type.name === 'docListItem') {
      tr = tr.setNodeMarkup(block.pos, undefined, { ...block.node.attrs, kind: listKind, numId })
    } else {
      tr = tr.setNodeMarkup(block.pos, editor.schema.nodes.docListItem, {
        ...block.node.attrs,
        styleId: null,
        kind: listKind,
        numId,
        ilvl: 0,
        aiChanged: true,
      })
    }
  }
  editor.view.dispatch(tr)
  return success(
    `Applied a ${kind} list (numId ${numId}, ${presetLabel}) to ${eligible.length} block(s)`
      + (skipped > 0 ? `; ${skipped} non-paragraph block(s) skipped.` : '.'),
    'apply list style',
    { blocks: eligible.length, skipped, numId: String(numId), kind },
  )
}

const TARGET_SCHEMA = {
  type: 'object',
  description:
    'same target vocabulary as apply_commands: { blockIndexes?, nodeType?, headingLevel?, containsText?, matchCase?, scope? }',
  properties: {
    nodeType: { type: 'string', enum: ['docHeading', 'docParagraph', 'docListItem', 'image'] },
    headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
    containsText: { type: 'string' },
    matchCase: { type: 'boolean' },
    blockIndexes: { type: 'array', items: { type: 'integer', minimum: 0 } },
    scope: { type: 'string', enum: ['selection', 'document'] },
  },
} as const

/** A3: create a document character style (written back to styles.xml on save). */
function createCharacterStyleOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  if (!ctx.editor.isEditable) {
    return failure('the document is read-only or protected; style definitions cannot be modified')
  }
  const app = ctx.app
  if (!app) return failure('format.createCharacterStyle needs the app context (host call path)')
  const name = String(input.name ?? '').trim()
  if (!name) return failure('name is required (the style display name)')
  if (name.length > 64) return failure('name must be at most 64 characters')
  const hex = input.color === undefined ? undefined : String(input.color).replace(/^#/, '').toUpperCase()
  if (hex !== undefined && !/^[0-9A-F]{6}$/.test(hex)) {
    return failure('color must be a 6-digit hex value without #')
  }
  let sizeHalfPoints: number | undefined
  if (input.sizeHalfPoints !== undefined) {
    const parsed = intOf(input.sizeHalfPoints)
    if (parsed === null || parsed < 2 || parsed > 400) {
      return failure('sizeHalfPoints must be an integer 2..400 (point size × 2)')
    }
    sizeHalfPoints = parsed
  }
  const font = input.font === undefined ? undefined : String(input.font).trim()
  if (font !== undefined && (font.length === 0 || font.length > 128)) {
    return failure('font must be a non-empty font name (max 128 chars)')
  }
  const used = new Set(app.getCharacterStyleIds())
  let base = String(input.styleId ?? name).replace(/[^A-Za-z0-9]/g, '')
  if (!base) base = 'AgentCharStyle'
  if (/^[0-9]/.test(base)) base = `Style${base}`
  if (used.has(base) && input.styleId !== undefined) {
    return failure(`styleId "${base}" already exists; pass a unique styleId (or omit it to auto-number)`)
  }
  let styleId = base
  for (let n = 2; used.has(styleId); n++) styleId = `${base}${n}`
  const rPr: NonNullable<import('@genoffice/docx-engine').StyleUpsert['rPr']> = {}
  if (input.bold !== undefined) {
    if (typeof input.bold !== 'boolean') return failure('bold must be a boolean')
    if (input.bold) rPr.bold = true
  }
  if (input.italic !== undefined) {
    if (typeof input.italic !== 'boolean') return failure('italic must be a boolean')
    if (input.italic) rPr.italic = true
  }
  if (input.underline !== undefined) {
    if (typeof input.underline !== 'boolean') return failure('underline must be a boolean')
    if (input.underline) rPr.underline = true
  }
  if (input.strike !== undefined) {
    if (typeof input.strike !== 'boolean') return failure('strike must be a boolean')
    if (input.strike) rPr.strike = true
  }
  if (hex !== undefined) rPr.color = hex
  if (font !== undefined) rPr.font = font
  if (sizeHalfPoints !== undefined) rPr.sizeHalfPoints = sizeHalfPoints
  const upsert: import('@genoffice/docx-engine').StyleUpsert = {
    styleId,
    type: 'character',
    name,
    ...(Object.keys(rPr).length > 0 ? { rPr } : {}),
  }
  if (ctx.dryRun) {
    return planned(`would create character style "${name}" (styleId ${styleId})`, { styleId, name })
  }
  app.addStyleUpsert(upsert)
  return success(
    `Created character style "${name}" (styleId ${styleId}); it is applied by styleId and saved with the document.`,
    'create character style',
    { styleId, name, rPr },
  )
}

export const formatOperations: OperationDef[] = [
  {
    name: 'format.changeCase',
    group: 'format',
    description:
      'format.changeCase { target, mode } — change text case (upper/lower/title/sentence) of the matched blocks',
    inputSchema: {
      type: 'object',
      properties: { target: TARGET_SCHEMA, mode: { type: 'string', enum: CASE_MODES } },
      required: ['target', 'mode'],
    },
    handler: changeCaseOp,
  },
  {
    name: 'format.clear',
    group: 'format',
    description:
      'format.clear { target } — clear direct character and paragraph formatting on the matched blocks '
      + '(Word Ctrl+Space + Ctrl+Q; paragraph style definitions still apply)',
    inputSchema: {
      type: 'object',
      properties: { target: TARGET_SCHEMA },
      required: ['target'],
    },
    handler: clearFormattingOp,
  },
  {
    name: 'format.setBidi',
    group: 'format',
    description:
      'format.setBidi { target, direction } — set LTR/RTL writing direction on the matched paragraph-like blocks',
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        direction: { type: 'string', enum: ['ltr', 'rtl'] },
      },
      required: ['target', 'direction'],
    },
    handler: setBidiOp,
  },
  {
    name: 'format.hangingIndent',
    group: 'format',
    description:
      'format.hangingIndent { target, steps? } — step the hanging indent (Word Ctrl+T / Ctrl+Shift+T); '
      + 'steps defaults to 1, negative steps decrease, one step = 0.5 inch',
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        steps: { type: 'integer', minimum: -10, maximum: 10 },
      },
      required: ['target'],
    },
    handler: hangingIndentOp,
  },
  {
    name: 'format.applyStyle',
    group: 'format',
    description:
      'format.applyStyle { target, style } — apply a gallery paragraph style: p (body) / h1 / h2 / h3 '
      + '(runs shed direct font/size/color so the style shows, like the ribbon gallery)',
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        style: { type: 'string', enum: STYLE_KEYS },
      },
      required: ['target', 'style'],
    },
    handler: applyStyleOp,
  },
  {
    name: 'format.listRestart',
    group: 'format',
    description:
      'format.listRestart { target } — restart numbering at the matched list (the first target list item; '
      + 'the new list keeps the abstract numbering definition and restarts at 1)',
    inputSchema: {
      type: 'object',
      properties: { target: TARGET_SCHEMA },
      required: ['target'],
    },
    handler: listRestartOp,
  },
  {
    name: 'format.listContinue',
    group: 'format',
    description:
      'format.listContinue { target } — merge the matched list into the preceding numbered list so numbering continues',
    inputSchema: {
      type: 'object',
      properties: { target: TARGET_SCHEMA },
      required: ['target'],
    },
    handler: listContinueOp,
  },
  {
    name: 'format.multilevelList',
    group: 'format',
    description:
      'format.multilevelList { target, preset? } — apply a multilevel list preset to the matched paragraphs/list items '
      + `(presets: ${MULTILEVEL_PRESETS.map((item) => `${item.id} = ${item.name}`).join('; ')}); headings are not converted`,
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        preset: {
          type: 'string',
          enum: MULTILEVEL_PRESETS.map((item) => item.id),
          description: 'gallery preset id; omitted = decimal (1. / 1.1. / 1.1.1.)',
        },
      },
      required: ['target'],
    },
    handler: multilevelListOp,
  },
  {
    name: 'format.applyCharacterStyle',
    group: 'format',
    description:
      'format.applyCharacterStyle { target, preset? | styleId?, mode? } — apply/toggle a character style: the built-in '
      + '"emphasis" (italic+accent) / "strong" (bold+accent) presets, or a document character style id (w:rStyle). '
      + 'mode defaults to toggle (active → unset)',
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        preset: { type: 'string', enum: ['emphasis', 'strong'] },
        styleId: { type: 'string', description: 'document character style id (see documentState.characterStyles)' },
        mode: { type: 'string', enum: CHAR_STYLE_MODES },
      },
      required: ['target'],
    },
    handler: applyCharacterStyleOp,
  },
  {
    name: 'format.setTabStops',
    group: 'format',
    description:
      'format.setTabStops { target, stops } — replace the ruler tab stops on the matched paragraphs (existing rel/ptab '
      + 'stops are preserved); stops: [{ posTwips, val: left|center|right|decimal|bar, leader? }], [] clears all',
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        stops: {
          type: 'array',
          maxItems: 32,
          items: {
            type: 'object',
            properties: {
              posTwips: { type: 'integer', minimum: 0, maximum: 31680 },
              val: { type: 'string', enum: TAB_VALS },
              leader: { type: 'string', enum: TAB_LEADERS },
            },
            required: ['posTwips', 'val'],
          },
        },
      },
      required: ['target', 'stops'],
    },
    handler: setTabStopsOp,
  },
  {
    name: 'format.insertBreak',
    group: 'format',
    description:
      'format.insertBreak { kind: line|column, blockIndex, at? } — insert a hard line break (Shift+Enter) or a column '
      + 'break (w:br type=column) at the start/end of a block (at defaults to end)',
    inputSchema: {
      type: 'object',
      properties: {
        kind: { type: 'string', enum: BREAK_KINDS },
        blockIndex: { type: 'integer', minimum: 0 },
        at: { type: 'string', enum: BREAK_ATS },
      },
      required: ['kind', 'blockIndex'],
    },
    handler: insertBreakOp,
  },
  {
    name: 'format.applyListStyle',
    group: 'format',
    description:
      'format.applyListStyle { target, kind: bullet|ordered, levels? } — create a numbering definition (custom levels '
      + 'or the built-in multilevel preset when levels are omitted) and convert the matched paragraphs/list items; '
      + 'headings and protected blocks are skipped',
    inputSchema: {
      type: 'object',
      properties: {
        target: TARGET_SCHEMA,
        kind: { type: 'string', enum: ['bullet', 'ordered'] },
        levels: {
          type: 'array',
          maxItems: 9,
          description: 'contiguous from level 0: [{ level, numFmt?, lvlText?, start? }]',
          items: {
            type: 'object',
            properties: {
              level: { type: 'integer', minimum: 0, maximum: 8 },
              numFmt: { type: 'string', description: 'OOXML numFmt, e.g. decimal/lowerLetter/chineseCountingThousand/bullet' },
              lvlText: { type: 'string', description: 'e.g. "%1." or "•"' },
              start: { type: 'integer', minimum: 1, maximum: 1000 },
            },
            required: ['level'],
          },
        },
      },
      required: ['target', 'kind'],
    },
    handler: applyListStyleOp,
  },
  {
    name: 'format.createCharacterStyle',
    group: 'format',
    description:
      'format.createCharacterStyle { name, styleId?, bold?, italic?, underline?, strike?, color?, font?, sizeHalfPoints? } '
      + '— define a reusable character style in styles.xml (styleId auto-derived from the name unless given; apply it '
      + 'with format.applyCharacterStyle { styleId })',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'style display name' },
        styleId: { type: 'string', description: 'optional explicit id (must be unique)' },
        bold: { type: 'boolean' },
        italic: { type: 'boolean' },
        underline: { type: 'boolean' },
        strike: { type: 'boolean' },
        color: { type: 'string', description: '6-digit hex without #' },
        font: { type: 'string', description: 'Latin/EA font name' },
        sizeHalfPoints: { type: 'integer', minimum: 2, maximum: 400, description: 'point size × 2' },
      },
      required: ['name'],
    },
    handler: createCharacterStyleOp,
  },
]
