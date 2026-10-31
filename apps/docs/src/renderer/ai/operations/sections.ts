import type { SectionInfo } from '@genoffice/docx-engine'
import { failure, intOf, planned, success } from './shared'
import type { OperationContext, OperationDef, OperationOutcome } from './types'

/**
 * B2 section ops: insert a section break and apply per-section page setup
 * through the App's real `applyPageSetup` pipeline (the same one the
 * set_page_setup tool and the Layout ribbon use).
 */

const SECTION_START_TYPES: SectionInfo['startType'][] = [
  'nextPage',
  'continuous',
  'evenPage',
  'oddPage',
]

/** Fields section.set forwards to applyPageSetup (same vocabulary as set_page_setup). */
const SETUP_FIELDS = [
  'orientation',
  'paperSize',
  'margins',
  'columns',
  'pageColor',
  'themeFonts',
  'themeColors',
] as const

function resolveSectionIndex(
  ctx: OperationContext,
  input: Record<string, unknown>,
): { index: number; total: number } | { error: string } {
  const app = ctx.app
  if (!app) return { error: 'section ops need the app context (host call path)' }
  const total = app.sections.length
  if (total === 0) return { error: 'no sections available (empty document?)' }
  if (input.sectionIndex === undefined) {
    return { index: Math.min(Math.max(0, app.activeSection), total - 1), total }
  }
  const parsed = intOf(input.sectionIndex)
  if (parsed === null || parsed < 0 || parsed >= total) {
    return { error: `sectionIndex must be an integer 0..${total - 1}` }
  }
  return { index: parsed, total }
}

function sectionBreakOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const type = String(input.type ?? '')
  if (!SECTION_START_TYPES.includes(type as SectionInfo['startType'])) {
    return failure(`type must be one of ${SECTION_START_TYPES.join('/')}`)
  }
  const doc = ctx.editor.state.doc
  let afterBlockIndex: number | undefined
  if (input.afterBlockIndex !== undefined) {
    const parsed = intOf(input.afterBlockIndex)
    if (parsed === null || parsed < 0 || parsed >= doc.childCount) {
      return failure(`afterBlockIndex must be an integer 0..${doc.childCount - 1}`)
    }
    afterBlockIndex = parsed
  }
  const app = ctx.app
  if (!app) return failure('section.break needs the app context (host call path)')
  const where =
    afterBlockIndex === undefined ? 'at the caret' : `after block ${afterBlockIndex}`
  if (ctx.dryRun) {
    return planned(`would insert a ${type} section break ${where}`, {
      startType: type,
      ...(afterBlockIndex !== undefined ? { afterBlockIndex } : {}),
    })
  }
  const before = doc.childCount
  // real ribbon helper: copies the current section's sectPr into a section-break
  // paragraph and writes the new section's start type on save
  app.insertSectionBreak(type as SectionInfo['startType'], afterBlockIndex)
  const after = ctx.editor.state.doc.childCount
  if (after <= before) {
    return failure('section break insertion failed (no section-break paragraph was added)')
  }
  return success(
    `Inserted a ${type} section break ${where}.`,
    'insert section break',
    { startType: type, ...(afterBlockIndex !== undefined ? { afterBlockIndex } : {}) },
  )
}

function sectionSetOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const app = ctx.app
  if (!app) return failure('section.set needs the app context (host call path)')
  const resolved = resolveSectionIndex(ctx, input)
  if ('error' in resolved) return failure(resolved.error)
  const { index, total } = resolved
  for (const key of Object.keys(input)) {
    if (key === 'op' || key === 'sectionIndex') continue
    if (!(SETUP_FIELDS as readonly string[]).includes(key)) {
      return failure(`unknown section.set field "${key}"; allowed: ${SETUP_FIELDS.join('/')}`)
    }
  }
  const payload: Record<string, unknown> = {}
  for (const field of SETUP_FIELDS) {
    if (input[field] !== undefined) payload[field] = input[field]
  }
  if (Object.keys(payload).length === 0) {
    return failure(`nothing to change: pass at least one of ${SETUP_FIELDS.join('/')}`)
  }
  // validate first through the real pipeline's dry-run (no setters touched), so
  // a failed op cannot leave half-applied section state behind
  const probe = app.applyPageSetup(payload, index, true)
  if (!probe.ok) return failure(probe.error)
  const where = `section ${index + 1} of ${total}`
  const detail = probe.output.replace(/^Would apply: /, '')
  if (ctx.dryRun) return planned(`would apply to ${where}: ${detail}`, { sectionIndex: index })
  const applied = app.applyPageSetup(payload, index, false)
  if (!applied.ok) return failure(applied.error)
  return success(
    `Applied to ${where}: ${applied.output.replace(/^Applied: /, '')}`,
    'section setup',
    { sectionIndex: index, fields: Object.keys(payload) },
  )
}

export const sectionOperations: OperationDef[] = [
  {
    name: 'section.break',
    group: 'section',
    description:
      'section.break { type, afterBlockIndex? } — insert a section break (nextPage/continuous/evenPage/oddPage) '
      + 'after a block; the new section inherits the current section geometry',
    inputSchema: {
      type: 'object',
      properties: {
        afterBlockIndex: {
          type: 'integer',
          minimum: 0,
          description: 'insert the break after this block index; omitted = at the caret',
        },
        type: {
          type: 'string',
          enum: SECTION_START_TYPES,
          description: 'how the new section starts',
        },
      },
      required: ['type'],
    },
    handler: sectionBreakOp,
  },
  {
    name: 'section.set',
    group: 'section',
    description:
      'section.set { sectionIndex?, orientation?, paperSize?, margins?, columns?, pageColor?, themeFonts?, themeColors? } '
      + '— page setup for one section (sectionIndex omitted = the caret section; pageColor/theme fields are document-wide, '
      + 'same as set_page_setup). Twips for margins (1 cm = 567); colors 6-digit hex without #',
    inputSchema: {
      type: 'object',
      properties: {
        sectionIndex: {
          type: 'integer',
          minimum: 0,
          description: 'target section (0-based); omitted = the section the caret is in',
        },
        orientation: { type: 'string', enum: ['portrait', 'landscape'] },
        paperSize: {
          type: 'string',
          enum: ['A4', 'letter', 'legal', 'B5'],
          description: 'paper size; combined with orientation',
        },
        margins: {
          type: 'object',
          description: 'twips; 1 cm = 567',
          properties: {
            top: { type: 'number', minimum: 0, maximum: 31680 },
            right: { type: 'number', minimum: 0, maximum: 31680 },
            bottom: { type: 'number', minimum: 0, maximum: 31680 },
            left: { type: 'number', minimum: 0, maximum: 31680 },
          },
        },
        columns: { type: 'integer', enum: [1, 2, 3] },
        pageColor: {
          type: ['string', 'null'],
          description: 'page background hex (no #), null to clear; document-wide',
        },
        themeFonts: {
          type: 'string',
          enum: ['office', 'songti', 'yahei', 'kaiti', 'times', 'arial'],
        },
        themeColors: { type: 'string', enum: ['office', 'facet', 'ion', 'retrospect'] },
      },
    },
    handler: sectionSetOp,
  },
]