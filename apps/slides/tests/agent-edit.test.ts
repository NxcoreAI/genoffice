/**
 * Embed host agent-edit surface: runTxn over a real in-memory deck wired into a
 * real session-state session — outline reflects applied ops, the deck-changed
 * push fires for a single attached view, archive-only ops set metaDirty, and the
 * silent save (persistSession) is invoked. slides-main's save/autofit pieces are
 * mocked (they need the full app shell).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { HeuristicMetrics } from '@genoffice/pptx-render'

const { sends, persistCalls } = vi.hoisted(() => ({
  sends: [] as Array<{ wcId: number; channel: string; payload: unknown }>,
  persistCalls: [] as Array<{ path: string }>,
}))

vi.mock('electron', () => ({
  webContents: {
    fromId: (id: number) => ({
      id,
      send: (channel: string, payload: unknown) => sends.push({ wcId: id, channel, payload }),
    }),
  },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => new HeuristicMetrics(),
}))
vi.mock('../src/main/slides-main', () => ({
  persistSession: vi.fn(async (session: { path: string }) => {
    persistCalls.push({ path: session.path })
  }),
  applyAutofitResize: (_s: unknown, _i: number, _id: string, rendered: unknown) => rendered,
  syncAutofitScale: (_s: unknown, _i: number, _id: string, rendered: unknown) => rendered,
}))

import { applyAgentDeckOps, describeAgentDeck } from '../src/main/agent-edit'
import { sessions, type Session } from '../src/main/session-state'

const WC_ID = 41001

async function makeSession(): Promise<Session> {
  const opened = await openPptx(await createBlankPptx())
  const session: Session = {
    path: '/tmp/agent-edit-deck.pptx',
    opened,
    fitWidthPx: 1280,
    undoStack: [],
    redoStack: [],
  }
  sessions.set(WC_ID, session)
  return session
}

beforeEach(() => {
  sends.length = 0
  persistCalls.length = 0
  sessions.delete(WC_ID)
})

describe('describeAgentDeck', () => {
  it('returns null without a session, outline + vocabulary with one', async () => {
    expect(describeAgentDeck(WC_ID)).toBeNull()
    const session = await makeSession()
    const r = describeAgentDeck(WC_ID)!
    expect(r.outline).toContain('The presentation has 1 pages')
    expect(r.outline).toContain('Page 1 (slideIndex=0):')
    expect(r.opVocabulary).toContain('text:')
    sessions.delete(WC_ID)
    expect(session).toBeTruthy()
  })
})

describe('applyAgentDeckOps', () => {
  it('applies ops, broadcasts the fresh deck to the single attached view, saves, and returns a fresh outline', async () => {
    await makeSession()
    const r = await applyAgentDeckOps(WC_ID, [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: 457200, y: 457200, cx: 7315200, cy: 914400 },
        paragraphs: [{ runs: [{ text: 'Agent written title' }] }],
      },
    ])
    expect(r.ok).toBe(true)
    expect(r.applied).toBe(true)
    expect(r.saved).toBe(true)
    expect(r.outline).toContain('Agent written title')
    // The broadcast must fire even with ONE attached view (scheduleDeckBroadcast would skip it).
    expect(sends).toHaveLength(1)
    expect(sends[0]).toMatchObject({ wcId: WC_ID, channel: 'slides:deck-changed' })
    const payload = sends[0]!.payload as { slides: unknown[]; size: { cx: number } }
    expect(payload.slides.length).toBe(1)
    expect(payload.size.cx).toBeGreaterThan(0)
    expect(persistCalls).toEqual([{ path: '/tmp/agent-edit-deck.pptx' }])
    sessions.delete(WC_ID)
  })

  it('atomic validation failure applies nothing: no broadcast, no save, no undo entry', async () => {
    const session = await makeSession()
    const undoBefore = session.undoStack.length
    const r = await applyAgentDeckOps(WC_ID, [
      { op: 'setFill', target: { slide: 0, el: 'ghost' }, color: '#FF0000' },
    ])
    expect(r.ok).toBe(true)
    expect(r.applied).toBe(false)
    expect(r.failures?.[0]?.error).toBeTruthy()
    expect(sends).toHaveLength(0)
    expect(persistCalls).toHaveLength(0)
    expect(session.undoStack).toHaveLength(undoBefore)
    sessions.delete(WC_ID)
  })

  it('archive-only ops mark the session meta-dirty for the save', async () => {
    const session = await makeSession()
    const r = await applyAgentDeckOps(WC_ID, [
      { op: 'setNotes', target: { slide: 0 }, text: 'notes from the agent' },
    ])
    expect(r.applied).toBe(true)
    expect(session.metaDirty).toBe(true)
    expect(persistCalls).toHaveLength(1)
    sessions.delete(WC_ID)
  })

  it('dry run validates without touching deck, history, broadcast or save', async () => {
    const session = await makeSession()
    const r = await applyAgentDeckOps(
      WC_ID,
      [{ op: 'addElement', target: { slide: 0 }, kind: 'rect', offset: { x: 0, y: 0, cx: 1, cy: 1 } }],
      { dryRun: true },
    )
    expect(r).toMatchObject({ ok: true, applied: false, dryRun: true })
    expect(r.plan?.[0]).toContain('addElement')
    expect(sends).toHaveLength(0)
    expect(persistCalls).toHaveLength(0)
    expect(session.undoStack).toHaveLength(0)
    sessions.delete(WC_ID)
  })

  it('rejects host-level errors: no session, no path, empty ops', async () => {
    expect((await applyAgentDeckOps(WC_ID, [{ op: 'setNotes', target: { slide: 0 }, text: 'x' }])).ok).toBe(false)
    const session = await makeSession()
    session.path = ''
    const r2 = await applyAgentDeckOps(WC_ID, [{ op: 'setNotes', target: { slide: 0 }, text: 'x' }])
    expect(r2.ok).toBe(false)
    session.path = '/tmp/agent-edit-deck.pptx'
    const r3 = await applyAgentDeckOps(WC_ID, [])
    expect(r3.ok).toBe(false)
    sessions.delete(WC_ID)
  })
})
