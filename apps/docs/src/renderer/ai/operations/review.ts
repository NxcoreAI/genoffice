import { flushSync } from 'react-dom'
import { hashProtectionPassword } from '@genoffice/docx-engine'
import {
  acceptAllRevisions,
  applyRevisionRanges,
  collectRevisions,
  rejectAllRevisions,
} from '../../editor/revisions'
import {
  executeCommands,
  resolveTargetBlocks,
  validateCommandTarget,
  type Target,
} from '../commands'
import type { AgentProtectResult, OperationContext, OperationDef, OperationOutcome } from './types'
import { failure, intOf, planned, success } from './shared'

/**
 * B4 review / protection ops: tracked-changes toggle, accept/reject revisions,
 * document protection + open-password encryption (both confirm-gated), and
 * find/replace through the apply_commands engine.
 *
 * Confirm gate: destructive/protective ops MUST carry `confirm: true`; without
 * it they fail with a CONFIRMATION_REQUIRED error and touch nothing (dryRun
 * included). Passwords are never logged, echoed in output, or stored in a
 * record — results only say whether a password is set.
 */

const CONFIRM_REQUIRED =
  'CONFIRMATION_REQUIRED: this operation changes document protection, revision state or file encryption '
  + 'and cannot be undone. Repeat the call with confirm:true after the user explicitly approved it.'

const PROTECT_MODES = [
  'none',
  'readOnly',
  'comments',
  'forms',
  'trackedChanges',
  'readOnlyRecommended',
] as const

type ProtectMode = (typeof PROTECT_MODES)[number]

/** Current tracker state (the extension storage is the live source of truth). */
function trackChangesEnabled(ctx: OperationContext): boolean {
  const storage = ctx.editor.storage as unknown as {
    trackChanges?: { enabled?: boolean }
  }
  return storage.trackChanges?.enabled === true
}

interface RevisionIdParse {
  indexes: number[]
  error?: string
}

function parseRevisionIds(raw: unknown): RevisionIdParse {
  if (!Array.isArray(raw) || raw.length === 0) {
    return { indexes: [], error: 'revisionIds must be a non-empty array of ids from read_revisions (e.g. "r0")' }
  }
  const indexes: number[] = []
  for (const value of raw) {
    const text = String(value ?? '').trim()
    const match = /^r?(\d+)$/.exec(text)
    const index = match ? Number(match[1]) : NaN
    if (!Number.isInteger(index) || index < 0) {
      return { indexes: [], error: `invalid revision id "${text}"; use the [rN] ids from read_revisions` }
    }
    if (indexes.includes(index)) {
      return { indexes: [], error: `duplicate revision id "r${index}"` }
    }
    indexes.push(index)
  }
  return { indexes }
}

function runRevisionOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
  mode: 'accept' | 'reject',
): OperationOutcome {
  if (input.confirm !== true) return failure(CONFIRM_REQUIRED)
  const wantAll = input.all === true
  const hasIds = input.revisionIds !== undefined
  if (!wantAll && !hasIds) {
    return failure('pass revisionIds (from read_revisions) or all:true')
  }
  if (wantAll && hasIds) {
    return failure('pass either revisionIds or all:true, not both')
  }
  const revisions = collectRevisions(ctx.editor.state.doc)
  if (revisions.length === 0) return failure('the document has no pending tracked revisions')
  const label = mode === 'accept' ? 'accept' : 'reject'

  if (wantAll) {
    if (ctx.dryRun) {
      return planned(`would ${label} all ${revisions.length} pending revision(s)`)
    }
    flushSync(() => {
      if (mode === 'accept') acceptAllRevisions(ctx.editor)
      else rejectAllRevisions(ctx.editor)
      const review = ctx.app?.review()
      if (review) review.dirtyRef.current = true
    })
    return success(
      `${mode === 'accept' ? 'Accepted' : 'Rejected'} all ${revisions.length} pending revision(s).`,
      `${label} all revisions`,
      { mode, revisions: revisions.length },
    )
  }

  const parsed = parseRevisionIds(input.revisionIds)
  if (parsed.error) return failure(parsed.error)
  const outOfRange = parsed.indexes.filter((index) => index >= revisions.length)
  if (outOfRange.length > 0) {
    return failure(
      `revision id(s) ${outOfRange.map((index) => `r${index}`).join(', ')} do not exist; `
      + `the document has ${revisions.length} pending revision(s) (ids r0..r${revisions.length - 1})`,
    )
  }
  const requested = [...parsed.indexes].sort((a, b) => b - a)
  if (ctx.dryRun) {
    return planned(
      `would ${label} ${requested.length} revision(s): ${requested.map((index) => `r${index}`).join(', ')}`,
      { mode, revisionIds: requested.map((index) => `r${index}`) },
    )
  }
  // Apply from the END of the document backwards: accepting/rejecting a later
  // revision never shifts the positions of earlier ones, so the ids resolved
  // against the pre-batch listing stay valid for the whole loop.
  let applied = 0
  for (const index of requested) {
    const current = collectRevisions(ctx.editor.state.doc)
    if (index >= current.length) break
    applyRevisionRanges(ctx.editor, [current[index]], mode)
    applied++
  }
  if (applied === 0) return failure(`no revision was ${label}ed (the document changed meanwhile)`)
  flushSync(() => {
    const review = ctx.app?.review()
    if (review) review.dirtyRef.current = true
  })
  return success(
    `${mode === 'accept' ? 'Accepted' : 'Rejected'} ${applied} revision(s): `
    + `${parsed.indexes.map((index) => `r${index}`).join(', ')}.`,
    `${label} revisions`,
    { mode, revisionIds: parsed.indexes.map((index) => `r${index}`) },
  )
}

function reviewAcceptOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  return runRevisionOp(input, ctx, 'accept')
}

function reviewRejectOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  return runRevisionOp(input, ctx, 'reject')
}

async function trackChangesOp(
  input: Record<string, unknown>,
  ctx: OperationContext,
): Promise<OperationOutcome> {
  if (typeof input.enabled !== 'boolean') return failure('enabled must be true or false')
  const enabled = input.enabled
  const app = ctx.app
  if (!app) return failure('review.trackChanges needs the app context (host call path)')
  const current = trackChangesEnabled(ctx)
  if (current === enabled) {
    return success(
      `Track changes is already ${enabled ? 'on' : 'off'}.`,
      'track changes',
      { enabled },
    )
  }
  if (ctx.dryRun) {
    return planned(`would turn track changes ${enabled ? 'on' : 'off'}`, { enabled })
  }
  flushSync(() => app.setTrackChanges(enabled))
  // the App effect writes editor.storage.trackChanges after the render; give
  // passive effects one microtask to land before verifying
  await Promise.resolve()
  if (trackChangesEnabled(ctx) !== enabled) {
    return failure('track changes could not be toggled (the editor tracker did not update)')
  }
  return success(
    `Turned track changes ${enabled ? 'on' : 'off'}; subsequent edits are recorded as revisions.`,
    'track changes',
    { enabled },
  )
}

async function protectOp(input: Record<string, unknown>, ctx: OperationContext): Promise<OperationOutcome> {
  if (input.confirm !== true) return failure(CONFIRM_REQUIRED)
  const mode = String(input.mode ?? '') as ProtectMode
  if (!PROTECT_MODES.includes(mode)) {
    return failure(`mode must be one of ${PROTECT_MODES.join('/')}`)
  }
  const app = ctx.app
  if (!app) return failure('review.protect needs the app context (host call path)')
  let password: string | null = null
  if (input.password !== undefined && input.password !== null) {
    password = String(input.password)
    if (!password) return failure('password must be a non-empty string when provided')
    if (password.length > 255) return failure('password must be at most 255 characters')
  }
  if (mode === 'readOnlyRecommended' && password !== null) {
    return failure('readOnlyRecommended has no password; pass the password only with readOnly/comments/forms/trackedChanges')
  }
  const describe =
    mode === 'none'
      ? 'clear the editing restriction'
      : mode === 'readOnlyRecommended'
        ? 'recommend read-only opening (write protection, no password)'
        : `restrict editing to "${mode}"${password !== null ? ' with a password (redacted)' : ''}`
  if (ctx.dryRun) return planned(`would ${describe}`, { mode, passwordSet: password !== null })
  // Hash only on the apply path (iterated SHA-512 is deliberately slow).
  let hashFields: { hash: string; salt: string; spinCount: number; algorithmSid: number } | null = null
  if (password !== null) {
    try {
      hashFields = await hashProtectionPassword(password)
    } catch {
      return failure('could not hash the protection password (internal error; the password was not logged)')
    }
  }
  const payload: AgentProtectResult =
    mode === 'none'
      ? { protection: null }
      : mode === 'readOnlyRecommended'
        ? { writeProtection: { recommended: true } }
        : {
            protection: {
              edit: mode,
              enforced: true,
              ...(hashFields
                ? {
                    hash: hashFields.hash,
                    salt: hashFields.salt,
                    spinCount: hashFields.spinCount,
                    algorithmSid: hashFields.algorithmSid,
                  }
                : {}),
            },
          }
  try {
    await app.applyProtection(payload)
  } catch {
    return failure('could not apply the protection settings (internal error; no password was logged)')
  }
  // persist the App-state dirty flags before the batch's silent save reads them
  flushSync(() => {})
  return success(
    `Protection updated: ${describe}.`,
    'protect document',
    { mode, passwordSet: password !== null },
  )
}

async function encryptOp(input: Record<string, unknown>, ctx: OperationContext): Promise<OperationOutcome> {
  if (input.confirm !== true) return failure(CONFIRM_REQUIRED)
  const app = ctx.app
  if (!app) return failure('review.encrypt needs the app context (host call path)')
  const password = input.password === null || input.password === undefined ? null : String(input.password)
  if (password !== null) {
    if (!password) return failure('password must be a non-empty string, or null to remove the open password')
    if (password.length > 255) return failure('password must be at most 255 characters')
  }
  const describe =
    password === null
      ? 'remove the open (encryption) password; the next save writes an unencrypted file'
      : 'set an open (encryption) password (redacted); the next save writes an encrypted file'
  if (ctx.dryRun) return planned(`would ${describe}`, { passwordSet: password !== null })
  let result: { ok: boolean; error?: string }
  try {
    result = await app.setOpenPassword(password)
  } catch {
    return failure('could not update the open password (internal error; no password was logged)')
  }
  if (!result.ok) {
    // never forward the main-process error verbatim: it could echo request data
    return failure('the main process rejected the open-password change; the document was not modified')
  }
  flushSync(() => {})
  return success(
    password === null
      ? 'Open password removed; the next save writes an unencrypted file.'
      : 'Open password updated (redacted); the next save writes an encrypted file.',
    'encrypt document',
    { passwordSet: password !== null },
  )
}

function findReplaceOp(input: Record<string, unknown>, ctx: OperationContext): OperationOutcome {
  const find = typeof input.find === 'string' ? input.find : ''
  if (!find) return failure('find must be a non-empty string')
  if (find.length > 1000) return failure('find must be at most 1000 characters')
  const replace = typeof input.replace === 'string' ? input.replace : ''
  if (replace.length > 10_000) return failure('replace must be at most 10000 characters')
  if (input.wholeWord !== undefined && typeof input.wholeWord !== 'boolean') {
    return failure('wholeWord must be a boolean')
  }
  if (input.all !== undefined && typeof input.all !== 'boolean') {
    return failure('all must be a boolean')
  }
  const wholeWord = input.wholeWord === true
  const firstOnly = input.all === false
  const matchCase = input.matchCase !== false
  if (input.target !== undefined) {
    const targetError = validateCommandTarget(input.target)
    if (targetError) return failure(targetError)
  }
  const doc = ctx.editor.state.doc
  const needle = matchCase ? find : find.toLowerCase()
  const isWordChar = (ch: string | undefined): boolean => !!ch && /[A-Za-z0-9_]/.test(ch)
  // Count exactly like the engine (whole-word: no ASCII word char on either
  // side; CJK has no word boundaries so the flag behaves like a plain match).
  const scan = (text: string): number => {
    if (!text) return 0
    const haystack = matchCase ? text : text.toLowerCase()
    let count = 0
    let at = haystack.indexOf(needle)
    while (at !== -1) {
      if (wholeWord && (isWordChar(text[at - 1]) || isWordChar(text[at + needle.length]))) {
        at = haystack.indexOf(needle, at + 1)
        continue
      }
      count++
      at = haystack.indexOf(needle, at + needle.length)
    }
    return count
  }
  const blocks = input.target
    ? resolveTargetBlocks(doc, input.target as Target, ctx.editor.state.selection)
    : null
  const total = (blocks ?? Array.from({ length: doc.childCount }, (_v, index) => ({ node: doc.child(index) })))
    .reduce((sum, block) => sum + scan(block.node.textContent ?? ''), 0)
  const plannedCount = firstOnly ? Math.min(1, total) : total
  if (ctx.dryRun) {
    return planned(
      plannedCount > 0
        ? `would replace ${plannedCount} occurrence(s) of "${find}"`
        : `no occurrence of "${find}" found in the target`,
      { matches: plannedCount },
    )
  }
  if (total === 0) {
    return failure(`no occurrence of "${find}" found in the target blocks; nothing was replaced`)
  }
  const outcome = executeCommands(ctx.editor, {
    commands: [
      {
        replaceAllText: {
          containsText: find,
          replaceText: replace,
          matchCase,
          ...(wholeWord ? { wholeWord: true } : {}),
          ...(firstOnly ? { replaceFirstOnly: true } : {}),
          ...(input.target !== undefined ? { target: input.target as Target } : {}),
        },
      },
    ],
  })
  if (!outcome.ok) return failure(outcome.error ?? 'find/replace failed')
  const changedBlocks = outcome.results[0]?.changed ?? 0
  if (changedBlocks === 0) {
    return failure(`no occurrence of "${find}" was replaced (protected blocks are skipped)`)
  }
  return success(
    `Replaced ${plannedCount} occurrence(s) of "${find}".`,
    'find and replace',
    { matches: plannedCount },
  )
}

export const reviewOperations: OperationDef[] = [
  {
    name: 'review.trackChanges',
    group: 'review',
    description:
      'review.trackChanges { enabled } — turn tracked changes on/off (on = subsequent edits are recorded as revisions)',
    inputSchema: {
      type: 'object',
      properties: { enabled: { type: 'boolean' } },
      required: ['enabled'],
    },
    handler: trackChangesOp,
  },
  {
    name: 'review.accept',
    group: 'review',
    description:
      'review.accept { revisionIds? | all?, confirm:true } — ACCEPT pending tracked revisions by [rN] ids from '
      + 'read_revisions, or every pending revision with all:true. Destructive: requires confirm:true',
    inputSchema: {
      type: 'object',
      properties: {
        revisionIds: {
          type: 'array',
          items: { type: 'string' },
          description: 'ids from read_revisions, e.g. ["r0","r2"]',
        },
        all: { type: 'boolean', description: 'true = accept every pending revision' },
        confirm: { type: 'boolean', description: 'must be true (destructive)' },
      },
      required: ['confirm'],
    },
    handler: reviewAcceptOp,
  },
  {
    name: 'review.reject',
    group: 'review',
    description:
      'review.reject { revisionIds? | all?, confirm:true } — REJECT pending tracked revisions by [rN] ids from '
      + 'read_revisions, or every pending revision with all:true. Destructive: requires confirm:true',
    inputSchema: {
      type: 'object',
      properties: {
        revisionIds: { type: 'array', items: { type: 'string' } },
        all: { type: 'boolean' },
        confirm: { type: 'boolean' },
      },
      required: ['confirm'],
    },
    handler: reviewRejectOp,
  },
  {
    name: 'review.protect',
    group: 'review',
    description:
      'review.protect { mode, password?, confirm:true } — set/clear document protection: mode none/readOnly/'
      + 'comments/forms/trackedChanges (optional editing password) or readOnlyRecommended (no password). '
      + 'Requires confirm:true; the password is never echoed',
    inputSchema: {
      type: 'object',
      properties: {
        mode: { type: 'string', enum: PROTECT_MODES },
        password: { type: 'string', maxLength: 255, description: 'optional editing password (redacted in results)' },
        confirm: { type: 'boolean', description: 'must be true (protective)' },
      },
      required: ['mode', 'confirm'],
    },
    handler: protectOp,
  },
  {
    name: 'review.encrypt',
    group: 'review',
    description:
      'review.encrypt { password, confirm:true } — set/remove the OPEN (encryption) password; applied on the next '
      + 'save. Requires confirm:true; the password is never echoed',
    inputSchema: {
      type: 'object',
      properties: {
        password: {
          type: ['string', 'null'],
          maxLength: 255,
          description: 'new open password, or null to remove it',
        },
        confirm: { type: 'boolean', description: 'must be true (encrypts the file)' },
      },
      required: ['password', 'confirm'],
    },
    handler: encryptOp,
  },
  {
    name: 'review.findReplace',
    group: 'review',
    description:
      'review.findReplace { find, replace, matchCase?, wholeWord?, all?, target? } — replace plain-text occurrences in '
      + 'the target blocks (apply_commands replaceAllText engine, formatting preserved). wholeWord = only matches not '
      + 'adjacent to an ASCII letter/digit/underscore (CJK has no word boundaries); all:false = first occurrence only',
    inputSchema: {
      type: 'object',
      properties: {
        find: { type: 'string', minLength: 1, maxLength: 1000 },
        replace: { type: 'string', maxLength: 10000 },
        matchCase: { type: 'boolean', description: 'default true' },
        wholeWord: { type: 'boolean', description: 'match whole words only (ASCII word boundaries)' },
        all: { type: 'boolean', description: 'default true; false = only the first occurrence' },
        target: {
          type: 'object',
          description: 'apply_commands target vocabulary (narrows the scanned blocks)',
          properties: {
            nodeType: { type: 'string', enum: ['docHeading', 'docParagraph', 'docListItem', 'image'] },
            headingLevel: { type: 'integer', minimum: 1, maximum: 6 },
            containsText: { type: 'string' },
            matchCase: { type: 'boolean' },
            blockIndexes: { type: 'array', items: { type: 'integer', minimum: 0 } },
            scope: { type: 'string', enum: ['selection', 'document'] },
          },
        },
      },
      required: ['find', 'replace'],
    },
    handler: findReplaceOp,
  },
]