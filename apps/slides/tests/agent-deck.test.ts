/**
 * Agent deck assembly: page specs → merged multi-slide pptx. Same true
 * roundtrip style as page-spec.test.ts (reopen the bytes with openPptx);
 * the electron/session-state deps of agent-deck are injected as mocks.
 */
import { describe, it, expect, vi } from 'vitest'
import { openPptx, type TextElement } from '@genoffice/pptx-engine'
import { HeuristicMetrics } from '@genoffice/pptx-render'

vi.mock('electron', () => ({
  nativeImage: { createFromBuffer: () => ({ getSize: () => ({ width: 0, height: 0 }) }) },
}))
vi.mock('../src/main/session-state', () => ({
  getFontMetrics: () => new HeuristicMetrics(),
}))
vi.mock('@genoffice/electron-utils', () => ({
  fetchRemoteImage: vi.fn(async () => null),
}))

import { buildAgentDeckPptx } from '../src/main/agent-deck'

const pageSpec = (text: string, background = '#FFFFFF') =>
  JSON.stringify({
    background,
    elements: [
      {
        type: 'text',
        x: 80,
        y: 60,
        w: 900,
        h: 90,
        paragraphs: [{ runs: [{ text, sizePt: 40, bold: true, color: '#112233' }] }],
      },
    ],
  })

describe('buildAgentDeckPptx', () => {
  it('assembles pages in array order into one reopenable deck', async () => {
    const r = await buildAgentDeckPptx([pageSpec('Cover', '#0E1A2B'), pageSpec('Agenda')])
    expect(r.ok).toBe(true)
    if (!r.ok) return
    const opened = await openPptx(r.deck.bytes)
    expect(opened.deck.slides).toHaveLength(2)
    const textOf = (index: number) => {
      const texts = opened.deck.slides[index]!.elements.filter((e): e is TextElement => e.type === 'text')
      return texts[0]!.text!.paragraphs[0]!.runs[0]!.text
    }
    expect(textOf(0)).toBe('Cover')
    expect(textOf(1)).toBe('Agenda')
    expect(r.deck.warnings).toEqual([])
    expect(r.deck.imageFailures).toEqual([])
  })

  it('rejects empty input, over-cap input, and page-scoped parse errors', async () => {
    expect((await buildAgentDeckPptx([])).ok).toBe(false)
    expect((await buildAgentDeckPptx(Array.from({ length: 25 }, () => pageSpec('x')))).ok).toBe(false)
    const bad = await buildAgentDeckPptx([pageSpec('ok'), 'no json here'])
    expect(bad.ok).toBe(false)
    if (bad.ok) return
    expect(bad.error).toContain('page 2')
  })

  it('surfaces tolerant-parser warnings per page', async () => {
    const spec = JSON.stringify({
      elements: Array.from({ length: 50 }, (_, i) => ({
        type: 'shape',
        shape: 'rect',
        x: i * 10,
        y: 0,
        w: 8,
        h: 8,
        fill: '#FF0000',
      })),
    })
    const r = await buildAgentDeckPptx([spec])
    expect(r.ok).toBe(true)
    if (!r.ok) return
    expect(r.deck.warnings).toHaveLength(1)
    expect(r.deck.warnings[0]!.page).toBe(1)
    expect(r.deck.warnings[0]!.messages.join(' ')).toContain('48')
    const opened = await openPptx(r.deck.bytes)
    expect(opened.deck.slides[0]!.elements).toHaveLength(48)
  })
})
