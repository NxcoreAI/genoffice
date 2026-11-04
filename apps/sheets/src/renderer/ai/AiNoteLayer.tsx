/**
 * EverRoom embed: host annotation anchors on the grid.
 *
 * The host owns the annotation queue and pushes snapshots down over IPC; this
 * layer only renders them. Each noted A1 range gets a dashed box and a numbered
 * badge, hover shows a small card (original excerpt + note + edit/delete), and
 * the badge opens a compact editor whose Enter routes an `update` back to the
 * host. Deleting from the card routes a `remove`.
 *
 * Geometry is recomputed on grid scroll / zoom / sheet switch and on every
 * queue snapshot; anchors that leave the viewport (or whose sheet is not the
 * active one) simply drop out — the queue keeps the fact, the grid just cannot
 * show it right now.
 */
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'

import type { SheetsAskQueueAction, SheetsAskQueueItem } from '../../shared/desktop-api'
import { useI18n } from '../i18n/locale'
import type { UniverRuntime } from '../univer-state'
import { noteRangeRect, type NoteAnchorRect } from './note-anchors'

interface Props {
  runtime: UniverRuntime | null
  items: readonly SheetsAskQueueItem[]
  /** Update / remove actions are forwarded to the host as-is. */
  onAction: (action: SheetsAskQueueAction) => void
}

interface AnchorSlot {
  readonly item: SheetsAskQueueItem
  readonly rect: NoteAnchorRect
  readonly number: number
}

const HIDE_DELAY = 150
const CARD_WIDTH = 240
const COMPACT_WIDTH = 260
const GAP = 6
const EDGE = 8
const MAX_INSTRUCTION = 2000

function sameRect(a: NoteAnchorRect, b: NoteAnchorRect): boolean {
  return a.left === b.left && a.top === b.top && a.right === b.right && a.bottom === b.bottom
}

function sameSlots(a: readonly AnchorSlot[], b: readonly AnchorSlot[]): boolean {
  if (a.length !== b.length) return false
  for (let index = 0; index < a.length; index += 1) {
    const left = a[index]
    const right = b[index]
    if (!left || !right) return false
    if (left.item !== right.item || left.number !== right.number) return false
    if (!sameRect(left.rect, right.rect)) return false
  }
  return true
}

function slotAtPoint(slots: readonly AnchorSlot[], x: number, y: number): AnchorSlot | undefined {
  return slots.find(
    (slot) =>
      x >= slot.rect.left && x <= slot.rect.right && y >= slot.rect.top && y <= slot.rect.bottom,
  )
}

/** Place a floating box next to an anchor rect, flipping/clamping at the edges. */
function placeNear(
  rect: NoteAnchorRect,
  width: number,
  height: number,
): { left: number; top: number } {
  const below = rect.bottom + GAP
  const above = rect.top - GAP - height
  const top =
    below + height <= window.innerHeight - EDGE
      ? below
      : above >= EDGE
        ? above
        : Math.max(EDGE, window.innerHeight - EDGE - height)
  const left = Math.min(
    Math.max(EDGE, (rect.left + rect.right) / 2 - width / 2),
    Math.max(EDGE, window.innerWidth - width - EDGE),
  )
  return { left, top }
}

function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

export function AiNoteLayer({ runtime, items, onAction }: Props): React.JSX.Element | null {
  const { t } = useI18n()
  const [slots, setSlots] = useState<AnchorSlot[]>([])
  const [hoveredId, setHoveredId] = useState<string | null>(null)
  const [editing, setEditing] = useState<{ draftId: string; text: string } | null>(null)
  const itemsRef = useRef(items)
  itemsRef.current = items
  const slotsRef = useRef<readonly AnchorSlot[]>([])
  const hoveredRef = useRef<string | null>(null)
  hoveredRef.current = hoveredId
  const editingRef = useRef(editing)
  editingRef.current = editing
  const hideTimer = useRef<number | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const measure = useCallback(() => {
    const next: AnchorSlot[] = []
    for (const item of itemsRef.current) {
      const rect = noteRangeRect(runtime, item.sheetId, item.range)
      if (rect) next.push({ item, rect, number: next.length + 1 })
    }
    slotsRef.current = next
    setSlots((previous) => (sameSlots(previous, next) ? previous : next))
  }, [runtime])

  // Re-measure immediately whenever the queue or the workbook changes.
  useLayoutEffect(() => {
    measure()
  }, [measure, items])

  // Follow grid scroll / zoom / sheet switches and window resizes.
  useEffect(() => {
    if (!runtime) return
    const api = runtime.univerAPI
    const disposables = [
      api.addEvent(api.Event.Scroll, measure),
      api.addEvent(api.Event.SheetZoomChanged, measure),
      api.addEvent(api.Event.ActiveSheetChanged, measure),
    ]
    window.addEventListener('resize', measure)
    return () => {
      for (const disposable of disposables) disposable.dispose()
      window.removeEventListener('resize', measure)
    }
  }, [runtime, measure])

  // Keep the editing popover attached while the grid moves; a vanished anchor
  // (sheet switch, range scrolled away, queue update) closes it.
  useLayoutEffect(() => {
    if (!editing) return
    if (!slots.some((slot) => slot.item.id === editing.draftId)) setEditing(null)
  }, [editing, slots])

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus()
      inputRef.current?.select()
    }
  }, [editing])

  useEffect(() => {
    if (!editing) return
    const onPointerDown = (event: PointerEvent) => {
      if (
        event.target instanceof Element &&
        event.target.closest('.sheet-note-editor, .sheet-note-card')
      ) {
        return
      }
      setEditing(null)
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [editing])

  const cancelHide = useCallback(() => {
    if (hideTimer.current !== null) {
      window.clearTimeout(hideTimer.current)
      hideTimer.current = null
    }
  }, [])

  const scheduleHide = useCallback(() => {
    cancelHide()
    hideTimer.current = window.setTimeout(() => setHoveredId(null), HIDE_DELAY)
  }, [cancelHide])

  useEffect(() => cancelHide, [cancelHide])

  // Hover tracking lives on the document: the dashed box is click-through, so
  // the card follows the pointer over the anchored range without stealing the
  // grid's own cell interactions.
  useEffect(() => {
    if (!runtime) return
    const onMove = (event: PointerEvent) => {
      if (editingRef.current) return
      const target = event.target
      if (
        target instanceof Element &&
        target.closest('.sheet-note-card, .sheet-note-editor, .ai-ask-pop')
      ) {
        cancelHide()
        return
      }
      if (slotsRef.current.length === 0) {
        if (hoveredRef.current !== null) setHoveredId(null)
        return
      }
      const slot = slotAtPoint(slotsRef.current, event.clientX, event.clientY)
      if (slot) {
        cancelHide()
        setHoveredId((previous) => (previous === slot.item.id ? previous : slot.item.id))
        return
      }
      if (hoveredRef.current !== null) scheduleHide()
    }
    document.addEventListener('pointermove', onMove)
    return () => {
      document.removeEventListener('pointermove', onMove)
      cancelHide()
    }
  }, [runtime, cancelHide, scheduleHide])

  if (!runtime) return null

  const hoveredSlot = hoveredId ? (slots.find((slot) => slot.item.id === hoveredId) ?? null) : null
  const editingSlot = editing
    ? (slots.find((slot) => slot.item.id === editing.draftId) ?? null)
    : null

  const card = hoveredSlot ? (
    <div
      className="sheet-note-card"
      style={{ width: CARD_WIDTH, ...placeNear(hoveredSlot.rect, CARD_WIDTH, 96) }}
      role="tooltip"
      onMouseEnter={cancelHide}
      onMouseLeave={scheduleHide}
    >
      {hoveredSlot.item.excerpt ? (
        <div className="sheet-note-card-excerpt" title={hoveredSlot.item.excerpt}>
          {truncate(hoveredSlot.item.excerpt, 60)}
        </div>
      ) : null}
      <div className="sheet-note-card-note">{hoveredSlot.item.instruction}</div>
      <div className="sheet-note-card-foot">
        <button
          type="button"
          className="sheet-note-card-edit"
          onClick={() => {
            setHoveredId(null)
            setEditing({ draftId: hoveredSlot.item.id, text: hoveredSlot.item.instruction })
          }}
        >
          {t('aiNoteCardEdit')}
        </button>
        <button
          type="button"
          className="sheet-note-card-del"
          onClick={() => {
            setHoveredId(null)
            onAction({ kind: 'remove', draftId: hoveredSlot.item.id })
          }}
        >
          {t('aiNoteCardDelete')}
        </button>
      </div>
    </div>
  ) : null

  const editor =
    editing && editingSlot ? (
      <div
        className="ai-ask-pop ai-ask-pop-compact sheet-note-editor"
        style={{ width: COMPACT_WIDTH, ...placeNear(editingSlot.rect, COMPACT_WIDTH, 84) }}
        role="dialog"
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            setEditing(null)
          }
        }}
      >
        <input
          ref={inputRef}
          className="ai-ask-pop-input"
          value={editing.text}
          maxLength={MAX_INSTRUCTION}
          placeholder={t('aiNoteEditPlaceholder')}
          onChange={(event) =>
            setEditing((previous) => (previous ? { ...previous, text: event.target.value } : null))
          }
          onKeyDown={(event) => {
            if (event.key !== 'Enter' || event.nativeEvent.isComposing) return
            event.preventDefault()
            const instruction = editing.text.trim()
            if (!instruction) return
            onAction({ kind: 'update', draftId: editing.draftId, instruction })
            setEditing(null)
          }}
        />
        <button
          type="button"
          className="sheet-note-pop-delete"
          onClick={() => {
            onAction({ kind: 'remove', draftId: editing.draftId })
            setEditing(null)
          }}
        >
          {t('aiNoteCardDelete')}
        </button>
      </div>
    ) : null

  return (
    <div className="sheet-note-layer">
      {slots.map((slot) => (
        <div
          key={slot.item.id}
          className="sheet-note-anchor"
          style={{
            left: slot.rect.left,
            top: slot.rect.top,
            width: slot.rect.right - slot.rect.left,
            height: slot.rect.bottom - slot.rect.top,
          }}
        >
          <button
            type="button"
            className="sheet-note-badge"
            aria-label={slot.item.instruction}
            onClick={() => {
              setHoveredId(null)
              setEditing({ draftId: slot.item.id, text: slot.item.instruction })
            }}
          >
            {slot.number}
          </button>
        </div>
      ))}
      {card}
      {editor}
    </div>
  )
}
