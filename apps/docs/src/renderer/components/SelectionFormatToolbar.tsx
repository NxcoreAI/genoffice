/**
 * Floating format toolbar on a text selection (WPS/Word-style): bold /
 * italic / underline / strikethrough, a font-color picker, and the host
 * "comment" pencil — one compact rounded bar that follows the live selection.
 *
 * EverRoom embed rationale: chromeHidden hides the ribbon, so without this
 * bar the selection had no formatting affordance at all. The bar replaces the
 * old standalone comment trigger (the pencil moved in here as one button);
 * clicking it opens the existing host-mode AiAskPopover on the same anchor.
 *
 * Positioning reuses the AiAskPopover trigger pattern: fixed coordinates from
 * coordsAtPos, refreshed on selectionUpdate / scroll / resize; mousedown is
 * prevented on the bar so the selection never dies while clicking a button.
 * Visibility keys off a non-empty selection only (not editor.isFocused):
 * clicking a bar button would otherwise hide it before the click lands, and
 * an unfocused-but-selected range (click away) keeps the bar reachable.
 */
import React, { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { ColorPicker } from '@genoffice/ui'
import { useI18n } from '../i18n/locale'

interface Props {
  editor: Editor
  /** open the host comment popover on the current selection (AiAskPopover) */
  onComment: () => void
}

const GAP = 6
const EDGE = 8

/** Live selection rect in viewport coordinates; null when nothing to anchor to. */
function selectionRect(editor: Editor): { left: number; right: number; bottom: number } | null {
  const { from, to, empty } = editor.state.selection
  if (empty) return null
  try {
    const a = editor.view.coordsAtPos(from)
    const b = editor.view.coordsAtPos(to)
    return {
      left: Math.min(a.left, b.left),
      right: Math.max(a.right, b.left),
      bottom: Math.max(a.bottom, b.bottom),
    }
  } catch {
    return null // positions can be transiently unmappable mid-transaction
  }
}

export function SelectionFormatToolbar({ editor, onComment }: Props): React.JSX.Element | null {
  const { t } = useI18n()
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null)
  const [colorOpen, setColorOpen] = useState(false)
  // mark commands (bold/…) change editor state without touching React: mirror
  // a counter off transactions so button active states re-render
  const [, setDocTick] = useState(0)
  const barRef = useRef<HTMLDivElement>(null)
  const popRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const onTransaction = () => setDocTick((tick) => tick + 1)
    editor.on('transaction', onTransaction)
    return () => {
      editor.off('transaction', onTransaction)
    }
  }, [editor])

  // keep the bar over the selection while it moves (caret drag, scroll, zoom)
  useLayoutEffect(() => {
    const update = () => {
      if (colorOpen) return // freeze while the palette is open
      const rect = selectionRect(editor)
      if (!rect) {
        setPos(null)
        return
      }
      const w = barRef.current?.offsetWidth ?? 260
      setPos({
        left: Math.min(
          Math.max(EDGE, (rect.left + rect.right) / 2 - w / 2),
          Math.max(EDGE, window.innerWidth - w - EDGE),
        ),
        top: Math.max(EDGE, rect.bottom + GAP),
      })
    }
    update()
    editor.on('selectionUpdate', update)
    editor.on('focus', update)
    const scroller = document.querySelector('.editor-scroll')
    scroller?.addEventListener('scroll', update, { passive: true })
    window.addEventListener('resize', update)
    return () => {
      editor.off('selectionUpdate', update)
      editor.off('focus', update)
      scroller?.removeEventListener('scroll', update)
      window.removeEventListener('resize', update)
    }
  }, [editor, colorOpen])

  // close the color palette on any outside click (capture: the editor
  // swallows pointer events)
  useEffect(() => {
    if (!colorOpen) return
    const onDown = (e: PointerEvent) => {
      if (popRef.current?.contains(e.target as Node)) return
      if (barRef.current?.contains(e.target as Node)) return
      setColorOpen(false)
    }
    document.addEventListener('pointerdown', onDown, true)
    return () => document.removeEventListener('pointerdown', onDown, true)
  }, [colorOpen])

  const canEdit = editor.isEditable
  // A click must run the mark command without blurring the editor
  const mark = useCallback(
    (name: 'bold' | 'italic' | 'underline' | 'strike') => {
      editor.chain().focus().toggleMark(name).run()
    },
    [editor],
  )

  if (!pos) return null

  const btn = (
    name: 'bold' | 'italic' | 'underline' | 'strike',
    label: React.ReactNode,
    tip: string,
    extraStyle?: React.CSSProperties,
  ) => {
    const active = editor.isActive(name)
    return (
      <button
        type="button"
        key={name}
        className={`sel-fmt-btn ${active ? 'active' : ''}`}
        disabled={!canEdit}
        data-tip={tip}
        aria-label={tip}
        aria-pressed={active}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => mark(name)}
        style={extraStyle}
      >
        {label}
      </button>
    )
  }

  const currentColor = (editor.getAttributes('docTextStyle').color as string | null) ?? null

  return (
    <div
      ref={barRef}
      className="sel-fmt-bar"
      style={{ left: pos.left, top: pos.top }}
      role="toolbar"
      aria-label={t('selFmtToolbarLabel')}
      onContextMenu={(e) => e.preventDefault()}
    >
      {btn('bold', <b>B</b>, t('ribbonBoldTip'))}
      {btn('italic', <i>I</i>, t('ribbonItalicTip'))}
      {btn('underline', <u>U</u>, t('ribbonUnderlineTip'))}
      {btn('strike', <s>ab</s>, t('ribbonStrikethrough'))}

      <span className="sel-fmt-sep" />

      {/* font color: swatch trigger + the shared Word palette pinned under it */}
      <div className="sel-fmt-color" style={{ position: 'relative' }}>
        <button
          type="button"
          className="sel-fmt-btn sel-fmt-color-btn"
          disabled={!canEdit}
          data-tip={t('ribbonFontColor')}
          aria-label={t('ribbonFontColor')}
          aria-expanded={colorOpen}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => setColorOpen((open) => !open)}
        >
          <span className="sel-fmt-color-glyph">A</span>
          <span
            className="sel-fmt-color-strip"
            style={{ background: currentColor ? `#${currentColor}` : 'var(--text)' }}
          />
        </button>
        {colorOpen && (
          <div ref={popRef} className="sel-fmt-color-pop" data-rb-panel="">
            <ColorPicker
              className="docs-color-pop"
              value={currentColor ? `#${currentColor}` : null}
              strings={{
                auto: t('ribbonNoColor'),
                themeColors: t('ribbonThemeColorsSection'),
                standardColors: t('ribbonStandardColors'),
                moreColors: t('ribbonMoreColors'),
                shadeTip: (r, c) => t('ribbonThemeColorShadeTip', { r, c }),
              }}
              onPick={(hex) => {
                editor
                  .chain()
                  .focus()
                  .setMark('docTextStyle', { color: hex ? hex.slice(1) : null })
                  .run()
                setColorOpen(false)
              }}
            />
          </div>
        )}
      </div>

      <span className="sel-fmt-sep" />

      {/* host comment (pencil) — same action as the old floating trigger */}
      <button
        type="button"
        className="sel-fmt-btn sel-fmt-comment"
        data-tip={t('aiNoteBtn')}
        aria-label={t('aiNoteBtn')}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onComment}
      >
        <svg viewBox="0 0 24 24" width="15" height="15" fill="none" aria-hidden>
          <path
            d="M4 20h4L19.5 8.5a2.1 2.1 0 0 0-3-3L5 17v3zM14 7l3 3"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
    </div>
  )
}
