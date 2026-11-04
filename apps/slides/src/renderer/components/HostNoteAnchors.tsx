/**
 * EverRoom embed: host annotation anchors on the canvas.
 *
 * The host owns the queue; this component only renders a snapshot. Each queued
 * note resolves to one or more elements on the current page, drawn as a dashed
 * outline plus a numbered badge. Hovering a badge opens a small card with the
 * captured excerpt, the note, and edit/delete actions — the same interaction
 * pattern as Word's host-note underline card.
 *
 * Positions are viewport rects, re-measured on scroll/resize/zoom by the app's
 * `measure` callback; the component itself never owns geometry, so a page with
 * a large zoom transform cannot distort the badges.
 */
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import type { SlidesAskQueueItem } from '../../shared/ipc'
import type { ElementPlacement } from '../ai/host-queue-anchors'
import type { AnchorRect } from './AiAskPopover'

export interface HostNoteAnchor {
  item: SlidesAskQueueItem
  /** 1-based position in the host queue (badge label) */
  order: number
  placements: ElementPlacement[]
}

interface Props {
  anchors: HostNoteAnchor[]
  /** Re-measure trigger: zoom level / page changes (scroll and resize are subscribed internally) */
  zoom: number
  /** Viewport rect of the placements (App measures against the live stage box) */
  measure: (placements: ElementPlacement[]) => AnchorRect | null
  onEdit: (item: SlidesAskQueueItem) => void
  onDelete: (item: SlidesAskQueueItem) => void
}

interface PlacedAnchor {
  anchor: HostNoteAnchor
  rect: AnchorRect
}

const CARD_WIDTH = 260
const CARD_GAP = 6
const EDGE = 8
/** Delay before the card hides, so the pointer can travel from badge to card */
const HIDE_DELAY = 150
/** Badge offset from the element's top-left corner */
const BADGE_OFFSET = 9

function sameRect(a: AnchorRect, b: AnchorRect): boolean {
  return a.left === b.left && a.top === b.top && a.right === b.right && a.bottom === b.bottom
}

function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

export function HostNoteAnchors({
  anchors,
  zoom,
  measure,
  onEdit,
  onDelete,
}: Props): React.JSX.Element | null {
  const { t } = useI18n()
  const [placed, setPlaced] = useState<PlacedAnchor[]>([])
  const [hoverId, setHoverId] = useState<string | null>(null)
  const cardRef = useRef<HTMLDivElement>(null)
  const hideTimer = useRef<number | null>(null)

  // Latest values without re-subscribing the scroll listener on every render
  const anchorsRef = useRef(anchors)
  anchorsRef.current = anchors
  const measureRef = useRef(measure)
  measureRef.current = measure

  const cancelHide = useCallback(() => {
    if (hideTimer.current !== null) {
      window.clearTimeout(hideTimer.current)
      hideTimer.current = null
    }
  }, [])

  const scheduleHide = useCallback(() => {
    cancelHide()
    hideTimer.current = window.setTimeout(() => setHoverId(null), HIDE_DELAY)
  }, [cancelHide])

  useEffect(() => cancelHide, [cancelHide])

  // Resolve viewport rects now and follow the canvas while it scrolls / resizes.
  // zoom is a dependency because a transform change fires no DOM event; the
  // observers cover layouts that move the stage without a window resize (AI
  // mode hides the ribbon/status bar, panes open/close) — a stale anchor would
  // otherwise hover over the wrong element.
  useLayoutEffect(() => {
    const update = () => {
      const next: PlacedAnchor[] = []
      for (const anchor of anchorsRef.current) {
        const rect = measureRef.current(anchor.placements)
        if (rect) next.push({ anchor, rect })
      }
      setPlaced((prev) => {
        if (prev.length !== next.length) return next
        for (let index = 0; index < next.length; index += 1) {
          const a = prev[index]
          const b = next[index]
          if (!a || !b) return next
          if (a.anchor.item.id !== b.anchor.item.id || !sameRect(a.rect, b.rect)) return next
        }
        return prev
      })
    }
    update()
    const scroller = document.querySelector('.stage-wrap')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    let observer: ResizeObserver | null = null
    if (typeof ResizeObserver !== 'undefined') {
      observer = new ResizeObserver(update)
      for (const target of [
        document.querySelector('.stage-wrap'),
        document.querySelector('.stage-rel'),
      ]) {
        if (target) observer.observe(target)
      }
    }
    return () => {
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
      observer?.disconnect()
    }
  }, [anchors, measure, zoom])

  // A removed / re-anchored note closes its card
  useEffect(() => {
    if (hoverId && !placed.some((entry) => entry.anchor.item.id === hoverId)) setHoverId(null)
  }, [placed, hoverId])

  if (placed.length === 0) return null

  const hovered = hoverId ? placed.find((entry) => entry.anchor.item.id === hoverId) : undefined
  const cardHeight = cardRef.current?.offsetHeight ?? 96
  const below = hovered ? hovered.rect.bottom + CARD_GAP : 0
  const above = hovered ? hovered.rect.top - CARD_GAP - cardHeight : 0
  const cardTop = !hovered
    ? 0
    : below + cardHeight <= window.innerHeight - EDGE
      ? below
      : above >= EDGE
        ? above
        : Math.max(EDGE, window.innerHeight - EDGE - cardHeight)
  const cardLeft = hovered
    ? Math.min(
        Math.max(EDGE, (hovered.rect.left + hovered.rect.right) / 2 - CARD_WIDTH / 2),
        Math.max(EDGE, window.innerWidth - CARD_WIDTH - EDGE),
      )
    : 0

  return (
    <>
      {placed.map(({ anchor, rect }) => (
        <React.Fragment key={anchor.item.id}>
          <div
            className="host-note-outline"
            style={{
              left: rect.left,
              top: rect.top,
              width: Math.max(0, rect.right - rect.left),
              height: Math.max(0, rect.bottom - rect.top),
            }}
            aria-hidden
          />
          <button
            type="button"
            className="host-note-badge"
            style={{
              left: Math.max(EDGE, rect.left - BADGE_OFFSET),
              top: Math.max(EDGE, rect.top - BADGE_OFFSET),
            }}
            onMouseEnter={() => {
              cancelHide()
              setHoverId(anchor.item.id)
            }}
            onMouseLeave={scheduleHide}
            onFocus={() => {
              cancelHide()
              setHoverId(anchor.item.id)
            }}
            onBlur={scheduleHide}
            onClick={() => onEdit(anchor.item)}
            aria-label={anchor.item.instruction}
          >
            {anchor.order}
          </button>
        </React.Fragment>
      ))}

      {hovered && (
        <div
          ref={cardRef}
          className="host-note-card"
          style={{ left: cardLeft, top: cardTop, width: CARD_WIDTH }}
          role="tooltip"
          onMouseEnter={cancelHide}
          onMouseLeave={scheduleHide}
        >
          {hovered.anchor.item.excerpt ? (
            <div className="host-note-card-excerpt" title={hovered.anchor.item.excerpt}>
              {truncate(hovered.anchor.item.excerpt, 60)}
            </div>
          ) : null}
          <div className="host-note-card-note">{hovered.anchor.item.instruction}</div>
          <div className="host-note-card-foot">
            <button
              type="button"
              className="host-note-card-edit"
              onClick={() => {
                setHoverId(null)
                onEdit(hovered.anchor.item)
              }}
            >
              {t('aiNoteCardEdit')}
            </button>
            <button
              type="button"
              className="host-note-card-del"
              onClick={() => {
                setHoverId(null)
                onDelete(hovered.anchor.item)
              }}
            >
              {t('aiNoteCardDelete')}
            </button>
          </div>
        </div>
      )}
    </>
  )
}
