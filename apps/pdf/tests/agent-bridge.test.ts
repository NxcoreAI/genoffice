import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

/**
 * pdf-main host agent executor (EverRoom): describe/apply are forwarded to
 * the renderer with a requestId and awaited via pdf:agent-tool-result; a
 * missing/late reply times out with { ok: false }; the waiters are released
 * when the view is destroyed. waitForAgentReady gates the host on the exact
 * working-copy path reported via pdf:agent-ready.
 */

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
  on: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
  isDestroyed: () => boolean
  listeners: Map<string, () => void>
}

let nextWcId = 1
let lastWebContents: FakeWebContents

function makeFakeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const wc: FakeWebContents = {
    id: nextWcId++,
    listeners,
    once: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    on: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    send: vi.fn(),
    setWindowOpenHandler: vi.fn(),
    isDestroyed: () => false,
  }
  lastWebContents = wc
  return wc
}

vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: vi.fn(() => new Promise(() => {})) },
  dialog: {},
  shell: {},
  BrowserWindow: class {},
  WebContentsView: class {
    webContents = makeFakeWebContents()
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    on: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    removeHandler: vi.fn(),
  },
}))

import { PDF_AGENT_CHANNELS } from '../src/shared/ipc'
import {
  applyAgentPdfTools,
  createPdfView,
  describeAgentPdf,
  waitForAgentReady,
} from '../src/main/pdf-main'

function makePdfFile(): string {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-agent-bridge-'))
  const path = join(dir, 'doc.pdf')
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  return path
}

/** Latest renderer-bound request sent to the fake webContents (channel-first arg). */
function lastAgentRequest(): { requestId: string; kind: string; calls?: unknown[] } {
  const calls = lastWebContents.send.mock.calls as unknown as Array<
    [string, Record<string, unknown>]
  >
  const hit = [...calls].reverse().find(([ch]) => ch === PDF_AGENT_CHANNELS.agentToolRequest)
  if (!hit) throw new Error('no agent-tool-request was sent')
  return hit[1] as { requestId: string; kind: string; calls?: unknown[] }
}

function replyFromRenderer(requestId: string, reply: unknown): void {
  handlers.get(PDF_AGENT_CHANNELS.agentToolResult)?.(
    { sender: { id: lastWebContents.id } },
    { requestId, reply },
  )
}

describe('pdf agent executor bridge', () => {
  it('forwards describe to the renderer and resolves the reply payload', async () => {
    createPdfView(makePdfFile())
    const pending = describeAgentPdf(lastWebContents)
    const request = lastAgentRequest()
    expect(request.kind).toBe('describe')
    replyFromRenderer(request.requestId, {
      ok: true,
      describe: { context: 'Current document: "doc.pdf", 2 pages.' },
    })
    await expect(pending).resolves.toEqual({
      context: 'Current document: "doc.pdf", 2 pages.',
    })
  })

  it('resolves describe errors instead of throwing', async () => {
    createPdfView(makePdfFile())
    const pending = describeAgentPdf(lastWebContents)
    replyFromRenderer(lastAgentRequest().requestId, { ok: false, error: 'no document is open' })
    await expect(pending).resolves.toEqual({ error: 'no document is open' })
  })

  it('sanitizes and forwards apply calls, rejecting an empty batch', async () => {
    createPdfView(makePdfFile())
    await expect(applyAgentPdfTools(lastWebContents, [])).resolves.toEqual({
      error: 'no valid tool calls',
    })
    const pending = applyAgentPdfTools(lastWebContents, [
      { id: 'c1', name: 'rotate_page', input: { page: 1, direction: 'right' } },
      { id: '', name: 'drop-me', input: {} },
      { id: 'c2', name: 'read_pages', input: {} },
    ])
    const request = lastAgentRequest()
    expect(request.kind).toBe('apply')
    expect(request.calls).toEqual([
      { id: 'c1', name: 'rotate_page', input: { page: 1, direction: 'right' } },
      { id: 'c2', name: 'read_pages', input: {} },
    ])
    replyFromRenderer(request.requestId, {
      ok: true,
      results: [
        { id: 'c1', name: 'rotate_page', output: 'Rotated page 1 (unsaved)', mutated: true },
        { id: 'c2', name: 'read_pages', output: 'page text…' },
      ],
    })
    await expect(pending).resolves.toEqual({
      results: [
        { id: 'c1', name: 'rotate_page', output: 'Rotated page 1 (unsaved)', mutated: true },
        { id: 'c2', name: 'read_pages', output: 'page text…' },
      ],
    })
  })
})

describe('pdf waitForAgentReady', () => {
  it('resolves true when the renderer reports the expected path', async () => {
    const path = makePdfFile()
    createPdfView(path)
    const pending = waitForAgentReady(lastWebContents, path, 500)
    handlers.get(PDF_AGENT_CHANNELS.agentReady)?.({ sender: { id: lastWebContents.id } }, {
      filePath: path,
    })
    await expect(pending).resolves.toBe(true)
  })

  it('resolves true for a report that arrived before the wait', async () => {
    const path = makePdfFile()
    createPdfView(path)
    handlers.get(PDF_AGENT_CHANNELS.agentReady)?.({ sender: { id: lastWebContents.id } }, {
      filePath: path,
    })
    await expect(waitForAgentReady(lastWebContents, path, 100)).resolves.toBe(true)
  })

  it('keeps waiting through a null (blank) report, then accepts the path', async () => {
    const path = makePdfFile()
    createPdfView(path)
    const pending = waitForAgentReady(lastWebContents, path, 1_000)
    handlers.get(PDF_AGENT_CHANNELS.agentReady)?.({ sender: { id: lastWebContents.id } }, {
      filePath: null,
    })
    handlers.get(PDF_AGENT_CHANNELS.agentReady)?.({ sender: { id: lastWebContents.id } }, {
      filePath: path,
    })
    await expect(pending).resolves.toBe(true)
  })

  it('ignores a report for a different path', async () => {
    const path = makePdfFile()
    createPdfView(path)
    const pending = waitForAgentReady(lastWebContents, path, 100)
    handlers.get(PDF_AGENT_CHANNELS.agentReady)?.({ sender: { id: lastWebContents.id } }, {
      filePath: '/tmp/some-other-file.pdf',
    })
    await expect(pending).resolves.toBe(false)
  })

  it('fails pending waits when the view is destroyed', async () => {
    const path = makePdfFile()
    createPdfView(path)
    const pending = waitForAgentReady(lastWebContents, path, 5_000)
    lastWebContents.listeners.get('destroyed')?.()
    await expect(pending).resolves.toBe(false)
  })
})
