/**
 * EverRoom embed: hover card over a host-note underline. The card shows the
 * note (plus the anchored excerpt) and offers edit / delete; hovering the card
 * itself keeps it open, leaving both hides it after a short delay — the
 * interaction pattern of mainstream comment UIs.
 *
 * The note text lives in the host's queue, so delete goes straight back over
 * IPC; edit re-dispatches AI_QUEUE_ANCHOR_CLICK, which AiAskPopover already
 * listens for to open in edit mode.
 */
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import {
  AI_QUEUE_ANCHOR_CLICK,
  HOST_NOTE_ANCHOR_CLASS,
  aiQueueAnchorsKey,
  queueAnchorRange,
} from '../editor/ai-queue-anchors'
import type { DocsAskQueueItem } from '../../shared/ipc'
import { rangeRect, type AnchorRect } from './AiAskPopover'

interface Props {
  editor: Editor
  items: DocsAskQueueItem[]
  onDelete: (item: DocsAskQueueItem) => void
}

const WIDTH = 240
const GAP = 6
const EDGE = 8
const HIDE_DELAY = 150

interface HoverState {
  qid: string
  rect: AnchorRect
}

export function HostNoteCard({ editor, items, onDelete }: Props): React.JSX.Element | null {
  const { t } = useI18n()
  const [hover, setHover] = useState<HoverState | null>(null)
  const cardRef = useRef<HTMLDivElement>(null)
  const hideTimer = useRef<number | null>(null)

  const cancelHide = useCallback(() => {
    if (hideTimer.current !== null) {
      window.clearTimeout(hideTimer.current)
      hideTimer.current = null
    }
  }, [])

  const scheduleHide = useCallback(() => {
    cancelHide()
    hideTimer.current = window.setTimeout(() => setHover(null), HIDE_DELAY)
  }, [cancelHide])

  useEffect(() => cancelHide, [cancelHide])

  /** hover target inside a host-note underline → that anchor's draft id */
  const qidAt = useCallback(
    (target: EventTarget | null): string | null => {
      const el = target instanceof Element ? target.closest(`.${HOST_NOTE_ANCHOR_CLASS}`) : null
      if (!el) return null
      const box = el.getBoundingClientRect()
      const hit = editor.view.posAtCoords({ left: (box.left + box.right) / 2, top: box.top + 1 })
      const pos = hit?.pos
      if (pos == null) return null
      const found = aiQueueAnchorsKey.getState(editor.state)?.find(pos, pos) ?? []
      return (found[0]?.spec as { qid?: string } | undefined)?.qid ?? null
    },
    [editor],
  )

  // document-level hover tracking (same pattern as the selection popover)
  useEffect(() => {
    const onOver = (e: MouseEvent) => {
      const qid = qidAt(e.target)
      if (!qid) return
      cancelHide()
      setHover((prev) => (prev?.qid === qid ? prev : { qid, rect: zeroRect() }))
    }
    const onOut = (e: MouseEvent) => {
      if (qidAt(e.target)) scheduleHide()
    }
    document.addEventListener('mouseover', onOver)
    document.addEventListener('mouseout', onOut)
    return () => {
      document.removeEventListener('mouseover', onOver)
      document.removeEventListener('mouseout', onOut)
    }
  }, [qidAt, cancelHide, scheduleHide])

  // resolve + follow the anchor rect while the card is visible; a vanished
  // anchor (text deleted / queue updated) closes the card
  useLayoutEffect(() => {
    if (!hover) return
    const update = () => {
      const rect = hostNoteRect(editor, hover.qid)
      if (rect) setHover({ qid: hover.qid, rect })
      else setHover(null)
    }
    update()
    const scroller = document.querySelector('.editor-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [hover, editor, items])

  const item = hover ? items.find((i) => i.id === hover.qid) : undefined
  if (!hover || !item) return null

  const height = cardRef.current?.offsetHeight ?? 96
  const below = hover.rect.bottom + GAP
  const above = hover.rect.top - GAP - height
  const top =
    below + height <= window.innerHeight - EDGE
      ? below
      : above >= EDGE
        ? above
        : Math.max(EDGE, window.innerHeight - EDGE - height)
  const left = Math.min(
    Math.max(EDGE, (hover.rect.left + hover.rect.right) / 2 - WIDTH / 2),
    Math.max(EDGE, window.innerWidth - WIDTH - EDGE),
  )

  return (
    <div
      ref={cardRef}
      className="host-note-card"
      style={{ left, top, width: WIDTH }}
      role="tooltip"
      onMouseEnter={cancelHide}
      onMouseLeave={scheduleHide}
    >
      {item.excerpt ? (
        <div className="host-note-card-excerpt" title={item.excerpt}>
          {truncate(item.excerpt, 60)}
        </div>
      ) : null}
      <div className="host-note-card-note">{item.instruction}</div>
      <div className="host-note-card-foot">
        <button
          type="button"
          className="host-note-card-edit"
          onClick={() => {
            setHover(null)
            window.dispatchEvent(
              new CustomEvent(AI_QUEUE_ANCHOR_CLICK, { detail: { qid: item.id } }),
            )
          }}
        >
          {t('aiNoteCardEdit')}
        </button>
        <button
          type="button"
          className="host-note-card-del"
          onClick={() => {
            setHover(null)
            onDelete(item)
          }}
        >
          {t('aiNoteCardDelete')}
        </button>
      </div>
    </div>
  )
}

function truncate(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max)}…` : text
}

function zeroRect(): AnchorRect {
  return { left: 0, top: 0, right: 0, bottom: 0 }
}

/** viewport rect of a host-note anchor; null once its range vanished */
function hostNoteRect(editor: Editor, qid: string): AnchorRect | null {
  // host-note decorations ride the shared anchor plugin (spec.qid = draft id)
  const range = queueAnchorRange(editor.state, qid)
  return range ? rangeRect(editor, range.from, range.to) : null
}
