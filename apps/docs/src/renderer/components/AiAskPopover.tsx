/**
 * Floating "Ask AI" entry on a text selection, plus the anchored popover that
 * captures one instruction — sent immediately (a normal selection-scoped run)
 * or queued as an anchored edit. Clicking a queued anchor in the document
 * re-opens the same popover in edit mode.
 *
 * Mounted at the app root with fixed positioning; ProseMirror's coordsAtPos
 * already speaks viewport coordinates. Positions refresh while the document
 * scrolls; outside clicks commit non-empty text into the queue (it stays
 * editable there) instead of discarding it.
 */
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { useI18n, type StringKey } from '../i18n/locale'
import { AI_QUEUE_ANCHOR_CLICK, queueAnchorRange } from '../editor/ai-queue-anchors'
import {
  EDIT_INSTRUCTION_MAX,
  EDIT_QUEUE_MAX,
  truncate,
  type DocsEditQueueItem,
} from '../ai/edit-queue'

interface AnchorRect {
  left: number
  top: number
  right: number
  bottom: number
}

/** viewport rect of a PM position range (shared with the host-note hover card) */
export function rangeRect(editor: Editor, from: number, to: number): AnchorRect | null {
  try {
    const a = editor.view.coordsAtPos(from)
    const b = editor.view.coordsAtPos(to)
    return {
      left: Math.min(a.left, b.left),
      top: Math.min(a.top, b.top),
      right: Math.max(a.right, b.right),
      bottom: Math.max(a.bottom, b.bottom),
    }
  } catch {
    return null // positions can be transiently unmappable mid-transaction
  }
}

export type { AnchorRect }

interface Props {
  editor: Editor
  queueFull: boolean
  getItem: (qid: string) => DocsEditQueueItem | undefined
  /** run the instruction right away as a normal selection-scoped request */
  onSendNow: (text: string) => void
  /** anchor the current selection and append to the queue */
  onQueueAdd: (text: string) => void
  onQueueUpdate: (qid: string, text: string) => void
  onQueueRemove: (qid: string) => void
  /**
   * Host-driven mode (EverRoom embed): there is no in-editor AI backend and no
   * anchor queue. The popover collapses to a compact single-line note input —
   * Enter hands the note to the host's queue, an outside click discards it.
   */
  hostMode?: boolean
  /** Host mode only: collect the note into the host's batch instead of sending now. */
  onHostQueue?: (text: string) => void
  /**
   * Host mode only: save an edit of an already-queued host note (IPC
   * askQueueAction 'update'), wired when the host queue is reachable.
   */
  onHostUpdate?: (draftId: string, text: string) => void
  /**
   * External open request (EverRoom embed): each new nonzero value opens the
   * popover on the live selection. Lets the selection format toolbar carry the
   * comment entry while this component keeps owning the input UI.
   */
  openNonce?: number
  /** Hide the floating trigger (the embed's format toolbar carries the entry). */
  hideTrigger?: boolean
}

type OpenState =
  | { mode: 'new' }
  | { mode: 'edit'; qid: string }
  /** host mode only: qid is the host draft id, not a local anchor id */
  | { mode: 'hostEdit'; qid: string }

const WIDTH = 340
const HOST_WIDTH = 260
const GAP = 8
const EDGE = 8
const EST_HEIGHT = 150
const HOST_EST_HEIGHT = 84

export type SelectionKind = 'image' | 'chart' | 'table' | 'text'

/** Suggestions match what the AI can actually do to the selected element */
const CHIP_KEYS: Record<SelectionKind, StringKey[]> = {
  image: ['aiChipReplaceImage', 'aiChipRegenImage', 'aiChipImageCaption'],
  chart: ['aiChipChartData', 'aiChipChartTitle'],
  table: ['aiChipTableEdit', 'aiChipPolish', 'aiChipFixGrammar'],
  text: ['aiChipPolish', 'aiChipShorten', 'aiChipExpand', 'aiChipFixGrammar'],
}

export function selectionKind(editor: Editor): SelectionKind {
  const sel = editor.state.selection
  if (sel instanceof CellSelection) return 'table'
  if (sel instanceof NodeSelection) {
    if (sel.node.type.name === 'docTable') return 'table'
    if (sel.node.type.name === 'docProtected') {
      // native docx charts are passthrough blocks + chartDisplay; AI/UI-created ones are blockType 'chart'
      if (sel.node.attrs.chartDisplay) return 'chart'
      const blockType = sel.node.attrs.blockType as string
      if (blockType === 'image') return 'image'
      if (blockType === 'chart') return 'chart'
      if (blockType === 'table') return 'table'
    }
  }
  return 'text'
}

export function AiAskPopover({
  editor,
  queueFull,
  getItem,
  onSendNow,
  onQueueAdd,
  onQueueUpdate,
  onQueueRemove,
  hostMode = false,
  onHostQueue,
  onHostUpdate,
  openNonce,
  hideTrigger = false,
}: Props): React.JSX.Element | null {
  const { t } = useI18n()
  const [trigger, setTrigger] = useState<{ left: number; top: number } | null>(null)
  const [open, setOpen] = useState<OpenState | null>(null)
  const [kind, setKind] = useState<SelectionKind>('text')
  const [text, setText] = useState('')
  const [rect, setRect] = useState<AnchorRect | null>(null)
  const boxRef = useRef<HTMLDivElement>(null)
  const triggerRef = useRef<HTMLButtonElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const openRef = useRef(open)
  openRef.current = open
  const textRef = useRef(text)
  textRef.current = text
  const queueFullRef = useRef(queueFull)
  queueFullRef.current = queueFull
  const hostModeRef = useRef(hostMode)
  hostModeRef.current = hostMode
  const onHostUpdateRef = useRef(onHostUpdate)
  onHostUpdateRef.current = onHostUpdate

  const anchorRect = useCallback((): AnchorRect | null => {
    const state = openRef.current
    if (!state) return null
    if (state.mode === 'edit') {
      const range = queueAnchorRange(editor.state, state.qid)
      return range ? rangeRect(editor, range.from, range.to) : null
    }
    if (state.mode === 'hostEdit') {
      // host-note anchors share the decoration plugin; the id IS the draft id
      const range = queueAnchorRange(editor.state, state.qid)
      return range ? rangeRect(editor, range.from, range.to) : null
    }
    const { from, to, empty } = editor.state.selection
    return empty ? null : rangeRect(editor, from, to)
  }, [editor])

  // floating trigger follows the live selection while the editor is focused
  useEffect(() => {
    const update = () => {
      if (openRef.current) {
        setTrigger(null)
        return
      }
      const { empty, to } = editor.state.selection
      if (empty || !editor.isFocused) {
        setTrigger(null)
        return
      }
      const c = rangeRect(editor, to, to)
      if (!c) {
        setTrigger(null)
        return
      }
      setTrigger({ left: c.right + 6, top: c.bottom + 4 })
    }
    const hide = () => setTrigger(null)
    editor.on('selectionUpdate', update)
    editor.on('focus', update)
    editor.on('blur', hide)
    const scroller = document.querySelector('.editor-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      editor.off('selectionUpdate', update)
      editor.off('focus', update)
      editor.off('blur', hide)
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [editor])

  // clicking a queued anchor in the document re-opens the popover on that item
  useEffect(() => {
    const onAnchorClick = (e: Event) => {
      const qid = (e as CustomEvent<{ qid: string }>).detail?.qid
      if (!qid) return
      if (hostModeRef.current) {
        // host mode: getItem returns DocsEditQueueItem-shaped host rows; the
        // popover edits in place through the host's update IPC
        const item = getItem(qid)
        if (!item || !onHostUpdateRef.current) return
        setText(item.instruction)
        setOpen({ mode: 'hostEdit', qid: item.qid })
        setTrigger(null)
        return
      }
      const item = getItem(qid)
      if (!item) return
      setText(item.instruction)
      setOpen({ mode: 'edit', qid: item.qid })
      setTrigger(null)
    }
    window.addEventListener(AI_QUEUE_ANCHOR_CLICK, onAnchorClick)
    return () => window.removeEventListener(AI_QUEUE_ANCHOR_CLICK, onAnchorClick)
  }, [getItem])

  // external open request (format toolbar's comment button): a fresh nonce
  // opens the popover on the live selection, same as the trigger click did
  useEffect(() => {
    if (!openNonce) return
    setText('')
    setKind(selectionKind(editor))
    setOpen({ mode: 'new' })
    setTrigger(null)
  }, [openNonce, editor])

  // popover position follows scroll/resize; a vanished anchor closes it
  useLayoutEffect(() => {
    if (!open) return
    const update = () => setRect(anchorRect())
    update()
    const scroller = document.querySelector('.editor-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [open, anchorRect])

  useEffect(() => {
    if (open) {
      inputRef.current?.focus()
      inputRef.current?.select()
    }
  }, [open])

  const close = useCallback(() => {
    setOpen(null)
    setText('')
  }, [])

  /** outside click keeps whatever was typed: a new note lands in the queue, an edit is saved */
  const commitOrClose = useCallback(() => {
    const state = openRef.current
    const value = textRef.current.trim()
    // Host mode has no anchor queue: an accidental outside click must not fire
    // a request at the host agent.
    if (hostModeRef.current) {
      close()
      return
    }
    if (state && value) {
      if (state.mode === 'edit') onQueueUpdate(state.qid, value)
      else if (!queueFullRef.current) onQueueAdd(value)
    }
    close()
  }, [onQueueAdd, onQueueUpdate, close])

  useEffect(() => {
    if (!open) return
    const onPointerDown = (e: PointerEvent) => {
      if (boxRef.current?.contains(e.target as Node)) return
      commitOrClose()
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [open, commitOrClose])

  const openFromSelection = () => {
    setText('')
    setKind(selectionKind(editor))
    setOpen({ mode: 'new' })
    setTrigger(null)
  }

  const canSubmit = text.trim().length > 0
  const editItem = open?.mode === 'edit' ? getItem(open.qid) : undefined

  const excerpt = (() => {
    if (!open) return ''
    if (open.mode === 'edit' || open.mode === 'hostEdit') {
      const range = queueAnchorRange(editor.state, open.qid)
      return range
        ? editor.state.doc.textBetween(range.from, range.to, ' ', ' ')
        : (editItem?.capturedText ?? '')
    }
    const { from, to } = editor.state.selection
    return editor.state.doc.textBetween(from, to, ' ', ' ')
  })()
    .replace(/\s+/g, ' ')
    .trim()

  let box: React.JSX.Element | null = null
  if (open && rect && hostMode) {
    // Compact host-mode popover: one short input, no title/excerpt/chips/footer.
    // Enter queues the note (Esc/outside closes) — the full form stays on the
    // standalone build.
    const width = HOST_WIDTH
    const height = boxRef.current?.offsetHeight ?? HOST_EST_HEIGHT
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
    const hostEdit = open.mode === 'hostEdit' ? open : null
    box = (
      <div
        ref={boxRef}
        className="ai-ask-pop ai-ask-pop-compact"
        style={{ left, top, width }}
        role="dialog"
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            e.preventDefault()
            e.stopPropagation()
            close()
          }
        }}
      >
        <input
          ref={inputRef}
          className="ai-ask-pop-input"
          value={text}
          maxLength={EDIT_INSTRUCTION_MAX}
          placeholder={t(hostEdit ? 'aiNoteEditPlaceholder' : 'aiNotePlaceholder')}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
              e.preventDefault()
              if (!canSubmit) return
              if (hostEdit) {
                onHostUpdate?.(hostEdit.qid, text.trim())
                close()
              } else if (onHostQueue) {
                onHostQueue(text.trim())
                close()
              }
            }
          }}
        />
        {hostEdit && (
          <button
            type="button"
            className="host-note-pop-delete"
            onClick={() => {
              onQueueRemove(hostEdit.qid)
              close()
            }}
          >
            {t('aiNoteCardDelete')}
          </button>
        )}
      </div>
    )
  } else if (open && rect) {
    const height = boxRef.current?.offsetHeight ?? EST_HEIGHT
    const below = rect.bottom + GAP
    const above = rect.top - GAP - height
    const top =
      below + height <= window.innerHeight - EDGE
        ? below
        : above >= EDGE
          ? above
          : Math.max(EDGE, window.innerHeight - EDGE - height)
    const left = Math.min(
      Math.max(EDGE, (rect.left + rect.right) / 2 - WIDTH / 2),
      Math.max(EDGE, window.innerWidth - WIDTH - EDGE),
    )
    box = (
      <div
        ref={boxRef}
        className="ai-ask-pop"
        style={{ left, top, width: WIDTH }}
        role="dialog"
        onKeyDown={(e) => {
          if (e.key === 'Escape') {
            e.preventDefault()
            e.stopPropagation()
            close()
          }
        }}
      >
        <button
          type="button"
          className="ai-ask-pop-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={close}
        >
          <svg viewBox="0 0 24 24" width="12" height="12" fill="none" aria-hidden>
            <path
              d="M6 6l12 12M18 6L6 18"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
            />
          </svg>
        </button>
        <div className="ai-ask-pop-title">
          {t(open.mode === 'edit' ? 'aiAskEditTitle' : 'aiAskTitle')}
        </div>
        {excerpt && <div className="ai-ask-pop-sub">{truncate(excerpt, 60)}</div>}
        <input
          ref={inputRef}
          className="ai-ask-pop-input"
          value={text}
          maxLength={EDIT_INSTRUCTION_MAX}
          placeholder={t('aiAskPlaceholder')}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
              e.preventDefault()
              if (!canSubmit) return
              if (open.mode === 'edit') {
                onQueueUpdate(open.qid, text.trim())
                close()
              } else if (hostMode) {
                onSendNow(text.trim())
                close()
              } else if (!queueFull) {
                onQueueAdd(text.trim())
                close()
              }
            }
          }}
        />
        {open.mode === 'new' && (
          <div className="ai-ask-pop-chips">
            {CHIP_KEYS[kind].map((key) => (
              <button
                key={key}
                className="ai-ask-chip"
                // filling rather than submitting: the shorthand almost always wants a qualifier
                onClick={() => {
                  setText(t(key))
                  inputRef.current?.focus()
                }}
              >
                {t(key)}
              </button>
            ))}
          </div>
        )}
        <div className="ai-ask-pop-foot">
          {open.mode === 'edit' ? (
            <>
              <button
                className="ai-ask-cancel"
                onClick={() => {
                  onQueueRemove(open.qid)
                  close()
                }}
              >
                {t('ribbonGroupDelete')}
              </button>
              <button
                className="ai-ask-confirm"
                disabled={!canSubmit}
                onClick={() => {
                  onQueueUpdate(open.qid, text.trim())
                  close()
                }}
              >
                {t('aiAskUpdate')}
              </button>
            </>
          ) : hostMode ? (
            <>
              {onHostQueue ? (
                <button
                  className="ai-ask-cancel"
                  disabled={!canSubmit}
                  onClick={() => {
                    onHostQueue(text.trim())
                    close()
                  }}
                >
                  {t('aiAskQueue')}
                </button>
              ) : null}
              <button
                className="ai-ask-confirm"
                disabled={!canSubmit}
                onClick={() => {
                  onSendNow(text.trim())
                  close()
                }}
              >
                {t('aiAskSendNow')}
              </button>
            </>
          ) : (
            <>
              <button
                className="ai-ask-cancel"
                disabled={!canSubmit}
                onClick={() => {
                  onSendNow(text.trim())
                  close()
                }}
              >
                {t('aiAskSendNow')}
              </button>
              <button
                className="ai-ask-confirm"
                disabled={!canSubmit || queueFull}
                data-tip={queueFull ? t('aiQueueFullNotice', { max: EDIT_QUEUE_MAX }) : undefined}
                onClick={() => {
                  onQueueAdd(text.trim())
                  close()
                }}
              >
                {t('aiAskQueue')}
              </button>
            </>
          )}
        </div>
      </div>
    )
  }

  return (
    <>
      {trigger && !open && !hideTrigger && (
        <button
          ref={triggerRef}
          className="ai-ask-trigger"
          style={{ left: trigger.left, top: trigger.top }}
          // keep the editor focused and the selection alive
          onMouseDown={(e) => e.preventDefault()}
          onClick={openFromSelection}
        >
          <svg viewBox="0 0 24 24" width="15" height="15" fill="none" aria-hidden>
            {hostMode ? (
              // pencil: annotating the selection, not asking for an AI rewrite
              <path
                d="M4 20h4L19.5 8.5a2.1 2.1 0 0 0-3-3L5 17v3zM14 7l3 3"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            ) : (
              <path
                d="M12 3l1.7 4.6L18 9.3l-4.3 1.7L12 15.6l-1.7-4.6L6 9.3l4.3-1.7L12 3zM19 15l.85 2.3L22 18.15l-2.15.85L19 21.3l-.85-2.3-2.15-.85 2.15-.85L19 15z"
                fill="currentColor"
              />
            )}
          </svg>
          {t(hostMode ? 'aiNoteBtn' : 'aiAskBtn')}
        </button>
      )}
      {box}
    </>
  )
}
