/**
 * EverRoom embed: host annotation (note) surfaces for the pdf viewer.
 *
 * The annotation queue is owned by the host (conversation side); this module
 * only draws it over the pages:
 * - `AskQueueAnchors`: the anchored passages of one page, tinted in the host
 *   note ink (distinct from the native markups) with a numbered badge.
 * - `AskQueueHoverCard`: the hover card (excerpt + note + edit/delete).
 * - `HostNotePopover`: the compact single-input popover opened from the
 *   selection bar (Enter queues the note) and from an existing anchor (edit).
 *
 * Coordinates: quads are PDF user space (`[x1,yMax,x2,yMax,x1,yMin,x2,yMin]`),
 * converted with the page geometry so zoom / rotation / page spread are handled
 * by the same helpers the native markups use.
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from './i18n/locale'
import { pdfRectToCss, quadToRect } from './annotations'
import type { PageGeom } from './annotations'
import type { AskAnchorRect } from './AiAskPopover'
import type { PdfAskQueueItem } from '../shared/ipc'

const POP_WIDTH = 260
const POP_GAP = 8
const POP_EDGE = 8
const POP_EST_HEIGHT = 84

const CARD_WIDTH = 240
const CARD_GAP = 6
const CARD_EDGE = 8
const CARD_EST_HEIGHT = 96
const CARD_HIDE_DELAY = 150

function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

/** Viewport rects are compared by value: recomputing on scroll must not re-render on sub-pixel jitter (React #185). */
function sameRect(a: AskAnchorRect | null, b: AskAnchorRect): boolean {
  return (
    a !== null &&
    a.left === b.left &&
    a.top === b.top &&
    a.right === b.right &&
    a.bottom === b.bottom
  )
}

/** Union of every anchor piece of `item` on this page, in viewport space. */
function anchorRect(id: string): AskAnchorRect | null {
  const nodes = document.querySelectorAll<HTMLElement>(`[data-pdf-ask-id="${CSS.escape(id)}"]`)
  if (nodes.length === 0) return null
  let left = Number.POSITIVE_INFINITY
  let top = Number.POSITIVE_INFINITY
  let right = Number.NEGATIVE_INFINITY
  let bottom = Number.NEGATIVE_INFINITY
  for (const node of nodes) {
    const r = node.getBoundingClientRect()
    left = Math.min(left, r.left)
    top = Math.min(top, r.top)
    right = Math.max(right, r.right)
    bottom = Math.max(bottom, r.bottom)
  }
  return { left, top, right, bottom }
}

/** Floating-box placement shared by the popover and the hover card: below the anchor, above when it does not fit. */
function placeBelow(
  rect: AskAnchorRect,
  width: number,
  height: number,
  gap: number,
  edge: number,
): { left: number; top: number } {
  const below = rect.bottom + gap
  const above = rect.top - gap - height
  const top =
    below + height <= window.innerHeight - edge
      ? below
      : above >= edge
        ? above
        : Math.max(edge, window.innerHeight - edge - height)
  const left = Math.min(
    Math.max(edge, (rect.left + rect.right) / 2 - width / 2),
    Math.max(edge, window.innerWidth - width - edge),
  )
  return { left, top }
}

/**
 * Compact host-mode popover: one input, no chips/excerpt/footer.
 * Enter (not while an IME composition is active) submits; Esc or an outside
 * click closes without submitting — the queue stays host-owned.
 */
export function HostNotePopover({
  rect,
  mode,
  initialText,
  onQueue,
  onUpdate,
  onDelete,
  onClose,
}: {
  rect: AskAnchorRect
  mode: 'new' | 'edit'
  initialText?: string
  onQueue?: (text: string) => void
  onUpdate?: (text: string) => void
  onDelete?: () => void
  onClose: () => void
}): ReactElement {
  const { t } = useI18n()
  const [text, setText] = useState(initialText ?? '')
  const [measuredH, setMeasuredH] = useState<number | null>(null)
  const boxRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const onCloseRef = useRef(onClose)
  onCloseRef.current = onClose

  useLayoutEffect(() => {
    setMeasuredH(boxRef.current?.offsetHeight ?? null)
  }, [])

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
  }, [])

  useEffect(() => {
    const onPointerDown = (e: PointerEvent) => {
      if (boxRef.current?.contains(e.target as Node)) return
      onCloseRef.current()
    }
    // capture phase: the pdf page's own handlers must not swallow the dismissal
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [])

  const canSubmit = text.trim().length > 0
  const submit = useCallback(() => {
    const value = text.trim()
    if (!value) return
    if (mode === 'edit') onUpdate?.(value)
    else onQueue?.(value)
  }, [text, mode, onQueue, onUpdate])

  const { left, top } = placeBelow(rect, POP_WIDTH, measuredH ?? POP_EST_HEIGHT, POP_GAP, POP_EDGE)

  return (
    <div
      ref={boxRef}
      className="ai-ask-pop ai-ask-pop-compact host-note-pop"
      style={{ left, top, width: POP_WIDTH }}
      role="dialog"
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.preventDefault()
          e.stopPropagation()
          onClose()
        }
      }}
    >
      <input
        ref={inputRef}
        className="ai-ask-pop-input host-note-pop-input"
        value={text}
        maxLength={2000}
        placeholder={t(mode === 'edit' ? 'aiNoteEditPlaceholder' : 'aiNotePlaceholder')}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
            e.preventDefault()
            submit()
          }
        }}
      />
      {mode === 'edit' && onDelete && (
        <button type="button" className="host-note-pop-delete" onClick={onDelete}>
          {t('aiNoteCardDelete')}
        </button>
      )}
    </div>
  )
}

/**
 * Anchored passages of one page. Each quad is an absolutely positioned box
 * inside the page; the first quad carries the queue-order badge. Hover opens
 * the host note card, click opens the edit popover.
 */
export function AskQueueAnchors({
  entries,
  geom,
  scale,
  editingId,
  onHover,
  onLeave,
  onOpen,
}: {
  entries: Array<{ item: PdfAskQueueItem; index: number }>
  geom: PageGeom
  scale: number
  editingId: string | null
  onHover: (id: string) => void
  onLeave: () => void
  onOpen: (item: PdfAskQueueItem, rect: AskAnchorRect) => void
}): ReactElement {
  const viewportRect = (el: HTMLElement): AskAnchorRect => {
    const r = el.getBoundingClientRect()
    return { left: r.left, top: r.top, right: r.right, bottom: r.bottom }
  }
  return (
    <>
      {entries.map(({ item, index }) => {
        const boxes = item.quads.map((q) => pdfRectToCss(geom, quadToRect(q), scale))
        const first = boxes[0]
        return (
          <div key={item.id} className="pdf-ask-anchor-set">
            {boxes.map((box, qi) => (
              <div
                key={qi}
                className={`pdf-ask-anchor${editingId === item.id ? ' is-editing' : ''}`}
                data-pdf-ask-id={item.id}
                style={box}
                onMouseEnter={() => onHover(item.id)}
                onMouseLeave={onLeave}
                onMouseDown={(e) => e.preventDefault()}
                onClick={(e) => {
                  e.stopPropagation()
                  const el = e.currentTarget as HTMLElement
                  onOpen(item, viewportRect(el))
                }}
              />
            ))}
            {first && (
              <button
                type="button"
                className="pdf-ask-anchor-badge"
                style={{ left: first.left + first.width - 4, top: Math.max(0, first.top - 8) }}
                onMouseEnter={() => onHover(item.id)}
                onMouseLeave={onLeave}
                onMouseDown={(e) => e.preventDefault()}
                onClick={(e) => {
                  e.stopPropagation()
                  const el = e.currentTarget as HTMLElement
                  // the badge sits at the anchor corner: anchor the popover on the
                  // detected passage, not on the badge itself
                  const rect = anchorRect(item.id) ?? viewportRect(el)
                  onOpen(item, rect)
                }}
              >
                {index + 1}
              </button>
            )}
          </div>
        )
      })}
    </>
  )
}

/**
 * Hover card over a host note anchor: excerpt + note + edit/delete. The card
 * follows its anchor through scroll/resize and closes when the anchor vanishes
 * (queue updated / page unmounted).
 */
export function AskQueueHoverCard({
  item,
  layoutKey,
  onEdit,
  onDelete,
  onMouseEnter,
  onMouseLeave,
}: {
  item: PdfAskQueueItem
  /** Changes whenever zoom / page rotation moved the page under the card. */
  layoutKey?: string
  onEdit: (item: PdfAskQueueItem, rect: AskAnchorRect) => void
  onDelete: (item: PdfAskQueueItem) => void
  onMouseEnter: () => void
  onMouseLeave: () => void
}): ReactElement | null {
  const { t } = useI18n()
  const [rect, setRect] = useState<AskAnchorRect | null>(null)
  const [measuredH, setMeasuredH] = useState<number | null>(null)
  const cardRef = useRef<HTMLDivElement>(null)

  useLayoutEffect(() => {
    setMeasuredH(cardRef.current?.offsetHeight ?? null)
  }, [])

  useLayoutEffect(() => {
    const update = () => {
      const next = anchorRect(item.id)
      setRect((prev) => (next === null ? null : sameRect(prev, next) ? prev : next))
    }
    update()
    const scroller = document.querySelector('.pdf-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [item.id, layoutKey])

  if (!rect) return null
  const { left, top } = placeBelow(
    rect,
    CARD_WIDTH,
    measuredH ?? CARD_EST_HEIGHT,
    CARD_GAP,
    CARD_EDGE,
  )

  return (
    <div
      ref={cardRef}
      className="host-note-card"
      style={{ left, top, width: CARD_WIDTH }}
      role="tooltip"
      onMouseEnter={onMouseEnter}
      onMouseLeave={onMouseLeave}
    >
      {item.excerpt ? (
        <div className="host-note-card-excerpt" title={item.excerpt}>
          {truncate(item.excerpt, 60)}
        </div>
      ) : null}
      <div className="host-note-card-note">{item.instruction}</div>
      <div className="host-note-card-foot">
        <button type="button" className="host-note-card-edit" onClick={() => onEdit(item, rect)}>
          {t('aiNoteCardEdit')}
        </button>
        <button type="button" className="host-note-card-del" onClick={() => onDelete(item)}>
          {t('aiNoteCardDelete')}
        </button>
      </div>
    </div>
  )
}

/** Hide-delay controller shared by the anchors and the hover card. */
export function useAskHover(): {
  hoveredId: string | null
  hover: (id: string) => void
  leave: () => void
  enterCard: () => void
  reset: () => void
} {
  const [hoveredId, setHoveredId] = useState<string | null>(null)
  const timerRef = useRef<number | null>(null)

  const cancel = useCallback(() => {
    if (timerRef.current !== null) {
      window.clearTimeout(timerRef.current)
      timerRef.current = null
    }
  }, [])

  const hover = useCallback(
    (id: string) => {
      cancel()
      setHoveredId(id)
    },
    [cancel],
  )

  const leave = useCallback(() => {
    cancel()
    timerRef.current = window.setTimeout(() => {
      timerRef.current = null
      setHoveredId(null)
    }, CARD_HIDE_DELAY)
  }, [cancel])

  const reset = useCallback(() => {
    cancel()
    setHoveredId(null)
  }, [cancel])

  useEffect(() => cancel, [cancel])

  return { hoveredId, hover, leave, enterCard: cancel, reset }
}
