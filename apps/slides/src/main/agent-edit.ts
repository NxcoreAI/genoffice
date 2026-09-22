/**
 * EverRoom Agent deck editing on a LIVE embed session: apply op transactions to
 * the session behind a webContents view, push the fresh render state to every
 * attached window (the agent's edits must be visible as they land), and persist
 * to disk so the host's fileSaved hook can re-import the version chain. Mirrors
 * the slides:apply-txn AI-panel surface (plan → history → journal → autofit
 * post-pass), plus the save the renderer path leaves to its user.
 */
import { webContents } from 'electron'

import { slideDurableId, materializeSlide } from '@genoffice/pptx-engine'

import { buildAgentDeckOutline } from '../shared/deck-outline'
import { opVocabulary } from '../shared/op-docs'
import { runTxn, type Op, type OpRecord } from './ops'
import {
  attachedIds,
  buildAllRenderSlides,
  journalOps,
  pushHistory,
  rebuildSlide,
  rebuildSlideWithReparse,
  sessions,
} from './session-state'
import { applyAutofitResize, persistSession, syncAutofitScale } from './slides-main'

const MAX_AGENT_OPS = 50

/** Ops whose executor changes live only in archive entries (notes/sections) or sldIdLst —
 *  the dedicated IPC shims mark the session meta-dirty for them, so the agent path does too. */
const ARCHIVE_ONLY_OPS = new Set([
  'setNotes',
  'setSections',
  'addSection',
  'renameSection',
  'removeSection',
  'moveSection',
  'moveSlide',
])

export interface AgentApplyResult {
  ok: boolean
  /** Host-level failure (no session / no path / bad request). */
  error?: string
  applied?: boolean
  dryRun?: boolean
  plan?: string[]
  records?: Array<{ op: string; target?: string; created?: string[] }>
  /** Per-op guided errors (with usage) — the txn result, not a host failure. */
  failures?: Array<{ index: number; error: string }>
  /** Silent save outcome; saveError means the deck changed but disk (and the version chain) did not. */
  saved?: boolean
  saveError?: string
  outline?: string
}

export function describeAgentDeck(wcId: number): { outline: string; opVocabulary: string } | null {
  const session = sessions.get(wcId)
  if (!session) return null
  return {
    outline: buildAgentDeckOutline(buildAllRenderSlides(session.opened, session.fitWidthPx)),
    opVocabulary: opVocabulary(),
  }
}

export async function applyAgentDeckOps(
  wcId: number,
  ops: unknown[],
  opts?: { dryRun?: boolean; isolation?: 'atomic' | 'per_op' },
): Promise<AgentApplyResult> {
  const session = sessions.get(wcId)
  if (!session) return { ok: false, error: 'no slides session for this view' }
  if (!session.path) return { ok: false, error: 'the deck has no file path yet' }
  if (!Array.isArray(ops) || ops.length === 0 || ops.length > MAX_AGENT_OPS) {
    return {
      ok: false,
      error: `ops must be a non-empty array (at most ${MAX_AGENT_OPS} per transaction).`,
    }
  }
  const isolation = opts?.isolation === 'per_op' ? ('per_op' as const) : ('atomic' as const)
  const txnOps = ops as Op[]
  const compact = (fails?: Array<{ index: number; error: string }>) =>
    fails?.map((f) => ({ index: f.index, error: f.error }))

  if (opts?.dryRun) {
    const r = runTxn(session.opened, { ops: txnOps, isolation, dryRun: true })
    return {
      ok: true,
      applied: false,
      dryRun: true,
      plan: r.plan ?? [],
      ...(r.failures?.length ? { failures: compact(r.failures) } : {}),
    }
  }

  // Plan before pushing history (a failed request must not clear the redo stack)
  const plan = runTxn(session.opened, { ops: txnOps, isolation, dryRun: true })
  const invalid = plan.failures?.length ?? 0
  if (isolation === 'atomic' ? invalid > 0 : invalid >= txnOps.length) {
    return { ok: true, applied: false, failures: compact(plan.failures) }
  }
  pushHistory(session)
  const r = runTxn(session.opened, { ops: txnOps, isolation })
  if (!r.applied) {
    session.undoStack.pop()
    return { ok: true, applied: false, failures: compact(r.failures) }
  }
  journalOps(session, 'batch', r.records ?? [])

  if ((r.records ?? []).some((rec) => ARCHIVE_ONLY_OPS.has(rec.op.op))) {
    session.metaDirty = true
  }

  // Autofit post-pass mirroring slides:apply-txn (render concerns outside the
  // executor). Slides are re-found by the executor-stamped durable id: a numeric
  // target.slide drifts when a later structural op shifts pages.
  const slideIdxOf = (rec: OpRecord): number => {
    if (rec.slideId)
      return session.opened.deck.slides.findIndex((s) => slideDurableId(s) === rec.slideId)
    return -1
  }
  const renderedByIdx = new Map<number, ReturnType<typeof rebuildSlide>>()
  for (const rec of r.records ?? []) {
    const o = rec.op
    const idx = slideIdxOf(rec)
    if (idx < 0) continue
    if (o.op === 'setTableStyle' || o.op === 'setChart') {
      rebuildSlideWithReparse(session, idx)
      renderedByIdx.delete(idx)
      continue
    }
    const id = o.target?.el
    if (!id || o.group) continue
    if (o.op !== 'setText' && o.op !== 'setFont' && o.op !== 'setParagraphFormat') continue
    if (o.op === 'setText' && (rec.after as { levelDirty?: boolean } | undefined)?.levelDirty)
      continue
    if (
      o.op === 'setParagraphFormat' &&
      (o.format as { indentDelta?: number } | undefined)?.indentDelta
    ) {
      materializeSlide(session.opened, idx)
      renderedByIdx.delete(idx)
      continue
    }
    let rendered = renderedByIdx.has(idx) ? renderedByIdx.get(idx)! : rebuildSlide(session, idx)
    rendered = applyAutofitResize(session, idx, id, rendered)
    rendered = syncAutofitScale(session, idx, id, rendered)
    renderedByIdx.set(idx, rendered)
  }

  // Push the fresh render state to every attached window — unconditionally,
  // unlike scheduleDeckBroadcast (a no-op with a single attached view, which is
  // exactly the embed case; the user must see each agent edit land).
  const slides = buildAllRenderSlides(session.opened, session.fitWidthPx)
  const payload = {
    slides,
    size: { cx: session.opened.deck.size.cx, cy: session.opened.deck.size.cy },
  }
  for (const id of attachedIds(session)) webContents.fromId(id)?.send('slides:deck-changed', payload)

  // Silent save: lands on disk and fires the host's fileSaved hook (edit-sync
  // re-import + version bump). A save failure does not undo the applied ops.
  let saved = true
  let saveError: string | undefined
  const wc = webContents.fromId(wcId)
  if (!wc) {
    saved = false
    saveError = 'the view was closed before the deck could be saved'
  } else {
    try {
      await persistSession(session, wc)
    } catch (err) {
      saved = false
      saveError = err instanceof Error ? err.message : String(err)
    }
  }

  return {
    ok: true,
    applied: true,
    records: (r.records ?? []).map((rec) => ({
      op: rec.op.op,
      ...(rec.op.target
        ? { target: `${rec.op.target.slide}${rec.op.target.el ? `/${rec.op.target.el}` : ''}` }
        : {}),
      ...(rec.created ? { created: rec.created } : {}),
    })),
    ...(r.failures?.length ? { failures: compact(r.failures) } : {}),
    saved,
    ...(saveError ? { saveError } : {}),
    outline: buildAgentDeckOutline(slides),
  }
}
