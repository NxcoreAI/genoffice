import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

/**
 * pdf-main host annotation IPC (EverRoom embed): `pdf:agent-ask` forwards a
 * normalized selection note to the installed host hook, `pdf:ask-queue-action`
 * forwards queue actions, and `pushPdfAskQueue` sends the host snapshot to the
 * right view. Malformed payloads never reach the host.
 */

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown

const harness = vi.hoisted(() => {
  type FakeWebContents = {
    id: number
    listeners: Map<string, () => void>
    send: ReturnType<typeof vi.fn>
    once: (event: string, handler: () => void) => void
    on: (event: string, handler: () => void) => void
    isDestroyed: () => boolean
    setWindowOpenHandler: () => void
  }
  const handlers = new Map<string, IpcHandler>()
  const views = new Map<number, FakeWebContents>()
  let nextWcId = 1
  const make = (): FakeWebContents => {
    const listeners = new Map<string, () => void>()
    const wc: FakeWebContents = {
      id: nextWcId++,
      listeners,
      send: vi.fn(),
      once: (event, handler) => listeners.set(event, handler),
      on: (event, handler) => listeners.set(event, handler),
      isDestroyed: () => false,
      setWindowOpenHandler: () => {},
    }
    views.set(wc.id, wc)
    return wc
  }
  return { handlers, views, make }
})

vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: vi.fn(() => new Promise(() => {})) },
  dialog: {},
  shell: {},
  BrowserWindow: class {},
  WebContentsView: class {
    webContents = harness.make()
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => harness.handlers.set(channel, handler)),
    on: vi.fn((channel: string, handler: IpcHandler) => harness.handlers.set(channel, handler)),
    removeHandler: vi.fn(),
  },
  webContents: { fromId: (id: number) => harness.views.get(id) ?? null },
}))

import { PDF_CHANNELS } from '../src/shared/ipc'
import {
  createPdfView,
  pushPdfAskQueue,
  setPdfAgentAskHook,
  setPdfAskQueueActionHook,
} from '../src/main/pdf-main'
import type { PdfAgentAskOp, PdfAskQueueAction } from '../src/shared/ipc'

function makePdfFile(): string {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-host-ask-'))
  const path = join(dir, 'doc.pdf')
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  return path
}

/** Create one fake pdf view and return its webContents (IPC registration is idempotent). */
function makeView(): { id: number; send: ReturnType<typeof vi.fn> } {
  const view = createPdfView(makePdfFile()) as unknown as {
    webContents: { id: number; send: ReturnType<typeof vi.fn> }
  }
  return view.webContents
}

const quad = [10, 720, 110, 720, 10, 700, 110, 700]

function askEvent(wcId: number, payload?: unknown): Promise<unknown> {
  return Promise.resolve(
    harness.handlers.get(PDF_CHANNELS.agentAsk)?.({ sender: { id: wcId } }, payload),
  )
}

afterEach(() => {
  setPdfAgentAskHook(null)
  setPdfAskQueueActionHook(null)
})

describe('pdf:agent-ask', () => {
  it('forwards a normalized selection note to the host hook', async () => {
    const wc = makeView()
    const seen: Array<{ wcId: number; op: PdfAgentAskOp }> = []
    setPdfAgentAskHook(async (wcId, op) => {
      seen.push({ wcId, op })
      return { ok: true }
    })
    await expect(
      askEvent(wc.id, {
        instruction: '  tighten this  ',
        pageIndex: 3,
        quads: [quad],
        excerpt: '  selected\n text ',
        mode: 'queue',
      }),
    ).resolves.toEqual({ ok: true })
    expect(seen).toEqual([
      {
        wcId: wc.id,
        op: {
          instruction: 'tighten this',
          pageIndex: 3,
          quads: [quad],
          excerpt: 'selected text',
          mode: 'queue',
        },
      },
    ])
  })

  it('reports NO_HOST_AGENT when the host installed no hook', async () => {
    const wc = makeView()
    await expect(
      askEvent(wc.id, { instruction: 'note', pageIndex: 0, quads: [quad] }),
    ).resolves.toEqual({ ok: false, error: 'NO_HOST_AGENT' })
  })

  it('rejects malformed payloads without calling the host', async () => {
    const wc = makeView()
    const hook = vi.fn(async () => ({ ok: true as const }))
    setPdfAgentAskHook(hook)
    await expect(
      askEvent(wc.id, { instruction: '', pageIndex: 0, quads: [quad] }),
    ).resolves.toEqual({ ok: false, error: 'EMPTY_INSTRUCTION_OR_ANCHOR' })
    await expect(
      askEvent(wc.id, { instruction: 'note', pageIndex: -1, quads: [quad] }),
    ).resolves.toEqual({ ok: false, error: 'EMPTY_INSTRUCTION_OR_ANCHOR' })
    await expect(
      askEvent(wc.id, { instruction: 'note', pageIndex: 0, quads: [] }),
    ).resolves.toEqual({ ok: false, error: 'EMPTY_INSTRUCTION_OR_ANCHOR' })
    expect(hook).not.toHaveBeenCalled()
  })

  it('turns a throwing host hook into an error result', async () => {
    const wc = makeView()
    setPdfAgentAskHook(async () => {
      throw new Error('room is gone')
    })
    await expect(
      askEvent(wc.id, { instruction: 'note', pageIndex: 0, quads: [quad] }),
    ).resolves.toEqual({ ok: false, error: 'room is gone' })
  })
})

describe('pdf:ask-queue-action', () => {
  function sendAction(wcId: number, payload: unknown): void {
    harness.handlers.get(PDF_CHANNELS.askQueueAction)?.({ sender: { id: wcId } }, payload)
  }

  it('forwards the three host actions with the calling view id', () => {
    const wc = makeView()
    const seen: Array<{ wcId: number; action: PdfAskQueueAction }> = []
    setPdfAskQueueActionHook((wcId, action) => seen.push({ wcId, action }))
    sendAction(wc.id, { kind: 'send-all' })
    sendAction(wc.id, { kind: 'remove', draftId: 'd1' })
    sendAction(wc.id, { kind: 'update', draftId: 'd1', instruction: ' fix this ' })
    expect(seen).toEqual([
      { wcId: wc.id, action: { kind: 'send-all' } },
      { wcId: wc.id, action: { kind: 'remove', draftId: 'd1' } },
      { wcId: wc.id, action: { kind: 'update', draftId: 'd1', instruction: 'fix this' } },
    ])
  })

  it('drops malformed actions and tolerates a missing hook', () => {
    const wc = makeView()
    const hook = vi.fn()
    setPdfAskQueueActionHook(hook)
    sendAction(wc.id, { kind: 'update', draftId: 'd1', instruction: '  ' })
    sendAction(wc.id, { kind: 'remove' })
    sendAction(wc.id, { kind: 'explode' })
    expect(hook).not.toHaveBeenCalled()
    setPdfAskQueueActionHook(null)
    expect(() => sendAction(wc.id, { kind: 'send-all' })).not.toThrow()
  })
})

describe('pushPdfAskQueue', () => {
  it('sends the normalized snapshot to the target view', () => {
    const wc = makeView()
    pushPdfAskQueue(wc.id, [
      { id: 'd1', instruction: 'note', excerpt: '  excerpt ', pageIndex: 0, quads: [quad] },
      { id: 'broken', instruction: 'no anchor', pageIndex: 5, quads: [] },
    ] as never)
    expect(wc.send).toHaveBeenCalledWith(PDF_CHANNELS.askQueue, [
      { id: 'd1', instruction: 'note', excerpt: 'excerpt', pageIndex: 0, quads: [quad] },
    ])
  })

  it('is a no-op for an unknown or destroyed view', () => {
    const wc = makeView()
    expect(() => pushPdfAskQueue(99_999, [])).not.toThrow()
    const destroyed = { ...wc, isDestroyed: () => true }
    harness.views.set(wc.id, destroyed as never)
    pushPdfAskQueue(wc.id, [{ id: 'd1', instruction: 'x', pageIndex: 0, quads: [quad] }])
    expect(wc.send).not.toHaveBeenCalled()
  })
})
