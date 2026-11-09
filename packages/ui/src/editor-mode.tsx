import { useEffect, useRef, useState, type MouseEvent as ReactMouseEvent } from 'react'

import { useDismissablePopover } from './popover-dismiss'
import { IconModeAi, IconModeRedo, IconModeSave, IconModeUndo, IconPanelToggle } from './icons'

/**
 * Two-state operation panel: "normal edit mode" (full ribbon) and "AI edit
 * mode" (panel hidden, page fills the window).
 *
 * The toggle lives at the very front of each app's panel so it reads as the
 * panel's own power switch; AI mode keeps only the actions an AI-assisted user
 * still needs by hand — save, undo, redo and the app's one-click AI action —
 * on a quiet floating bar (see `editor-mode.css`).
 */

export interface EditorModeToggleProps {
  /** true = AI mode (panel hidden); flips the chevron direction */
  hidden: boolean
  /** tooltip/aria label of the state the click leads to */
  label: string
  onClick: () => void
  /** app-level button class (each app's `.qa-btn` row styling) */
  className?: string
}

/** Front-of-panel toggle. Renders as the app's own tool button so the row
 *  metrics stay native; `data-tip` feeds the shared screentip layer. */
export function EditorModeToggle({ hidden, label, onClick, className }: EditorModeToggleProps) {
  return (
    <button
      type="button"
      className={`editor-mode-toggle${className ? ` ${className}` : ''}`}
      data-tip={label}
      aria-label={label}
      aria-pressed={hidden}
      onClick={onClick}
    >
      <IconPanelToggle size={16} dir={hidden ? 'expand' : 'collapse'} />
    </button>
  )
}

export interface EditorModeBarLabels {
  /** leave AI mode and restore the full panel */
  restore: string
  save: string
  undo: string
  redo: string
  /** the app's one-click AI action, e.g. "AI 排版" */
  ai: string
  /** aria/text tip of the caret that opens the quick AI actions menu */
  aiMenuTip?: string
}

/**
 * One quick AI capability in the AI-mode bar's action menu (the WPS-style
 * shortcut list). The first action doubles as the primary button when the
 * host passes a single entry without a menu.
 */
export interface EditorModeAiAction {
  /** stable id for React keys and selection routing */
  id: string
  /** visible label, e.g. "AI 排版" */
  label: string
  /** ScreenTip / aria text when different from the label (defaults to label) */
  tip?: string
}

export interface EditorModeBarProps {
  labels: EditorModeBarLabels
  canSave: boolean
  canUndo: boolean
  canRedo: boolean
  /** false = this build has no AI surface (embed builds without an agent channel) */
  showAi: boolean
  aiDisabled?: boolean
  onRestore: () => void
  onSave: () => void
  onUndo: () => void
  onRedo: () => void
  onAi: () => void
  /**
   * Quick AI capabilities for the AI button's dropdown (WPS-style shortcuts:
   * 排版 / 润色 / 总结 …). When provided, the AI entry becomes a split button —
   * the primary action runs the first entry, the caret opens the menu. When
   * omitted, the single `onAi` button is rendered as before.
   */
  aiActions?: ReadonlyArray<EditorModeAiAction & { run: () => void }>
}

/** The only chrome left in AI mode. Fixed overlay so hiding the panel never
 *  reflows the canvas; deliberately quiet until hovered. */
export function EditorModeBar({
  labels,
  canSave,
  canUndo,
  canRedo,
  showAi,
  aiDisabled,
  onRestore,
  onSave,
  onUndo,
  onRedo,
  onAi,
  aiActions,
}: EditorModeBarProps) {
  // Pressing a bar button must not move focus out of the editor: the app-level
  // handlers (and native undo/redo) act on the current selection.
  const keepFocus = (event: ReactMouseEvent<HTMLButtonElement>) => event.preventDefault()
  return (
    <div className="editor-mode-bar" role="toolbar" aria-label={labels.restore}>
      <button
        type="button"
        className="editor-mode-btn editor-mode-restore"
        data-tip={labels.restore}
        aria-label={labels.restore}
        aria-pressed
        onMouseDown={keepFocus}
        onClick={onRestore}
      >
        <IconPanelToggle size={15} dir="expand" />
        <span className="editor-mode-restore-label">{labels.restore}</span>
      </button>
      <span className="editor-mode-sep" aria-hidden="true" />
      <button
        type="button"
        className="editor-mode-btn"
        data-tip={labels.save}
        aria-label={labels.save}
        disabled={!canSave}
        onMouseDown={keepFocus}
        onClick={onSave}
      >
        <IconModeSave />
      </button>
      <button
        type="button"
        className="editor-mode-btn"
        data-tip={labels.undo}
        aria-label={labels.undo}
        disabled={!canUndo}
        onMouseDown={keepFocus}
        onClick={onUndo}
      >
        <IconModeUndo />
      </button>
      <button
        type="button"
        className="editor-mode-btn"
        data-tip={labels.redo}
        aria-label={labels.redo}
        disabled={!canRedo}
        onMouseDown={keepFocus}
        onClick={onRedo}
      >
        <IconModeRedo />
      </button>
      {showAi && (aiActions?.length ? (
        <EditorModeAiMenu
          labels={labels}
          {...(aiDisabled !== undefined ? { aiDisabled } : {})}
          onAi={onAi}
          actions={aiActions}
        />
      ) : (
        <>
          <span className="editor-mode-sep" aria-hidden="true" />
          <button
            type="button"
            className="editor-mode-btn editor-mode-ai"
            data-tip={labels.ai}
            aria-label={labels.ai}
            disabled={aiDisabled}
            onMouseDown={keepFocus}
            onClick={onAi}
          >
            <IconModeAi />
            <span className="editor-mode-ai-label">{labels.ai}</span>
          </button>
        </>
      ))}
    </div>
  )
}

/**
 * Split AI entry: the primary side runs the first quick action (same behavior
 * as the old single button), the caret opens the full shortcut menu. Keyboard:
 * arrows move within the menu, Enter runs, Escape closes (focus returns to the
 * editor because the bar never takes focus away from it).
 */
function EditorModeAiMenu({
  labels,
  aiDisabled = false,
  onAi,
  actions,
}: {
  labels: EditorModeBarLabels
  aiDisabled?: boolean
  onAi: () => void
  actions: ReadonlyArray<EditorModeAiAction & { run: () => void }>
}) {
  const [open, setOpen] = useState(false)
  const [active, setActive] = useState(0)
  const wrapRef = useRef<HTMLSpanElement>(null)
  const listRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEffect(() => {
    if (!open) return
    setActive(0)
  }, [open])
  useEffect(() => {
    if (!open) return
    // jsdom has no scrollIntoView; guard so tests can render the menu
    listRef.current?.querySelectorAll('.editor-mode-ai-item')[active]?.scrollIntoView?.({ block: 'nearest' })
  }, [open, active])

  const primary = actions[0]!
  const primaryLabel = labels.ai || primary.label
  const primaryTip = primary.tip ?? primaryLabel
  // The primary side runs the first action when the host passed a real list;
  // onAi remains the fallback for single-action hosts that only set labels.ai.
  const runPrimary = actions.length > 1 ? primary.run : onAi

  const runAt = (index: number) => {
    const action = actions[index]
    if (!action) return
    setOpen(false)
    action.run()
  }

  return (
    <>
      <span className="editor-mode-sep" aria-hidden="true" />
      <span className="editor-mode-ai-split" ref={wrapRef}>
        <button
          type="button"
          className="editor-mode-btn editor-mode-ai"
          data-tip={primaryTip}
          aria-label={primaryTip}
          disabled={aiDisabled}
          onMouseDown={keepFocusStatic}
          onClick={() => runPrimary()}
        >
          <IconModeAi />
          <span className="editor-mode-ai-label">{primaryLabel}</span>
        </button>
        <button
          type="button"
          className={`editor-mode-btn editor-mode-ai-caret${open ? ' is-open' : ''}`}
          data-tip={labels.aiMenuTip ?? primaryLabel}
          aria-label={labels.aiMenuTip ?? primaryLabel}
          aria-haspopup="menu"
          aria-expanded={open}
          disabled={aiDisabled}
          onMouseDown={keepFocusStatic}
          onClick={() => setOpen((value) => !value)}
          onKeyDown={(event) => {
            if (!open && (event.key === 'ArrowDown' || event.key === 'ArrowUp')) {
              event.preventDefault()
              setOpen(true)
            }
          }}
        >
          <span className="editor-mode-caret" aria-hidden="true" />
        </button>
        {open && (
          <div
            className="editor-mode-ai-menu"
            role="menu"
            ref={listRef}
            aria-label={labels.aiMenuTip ?? primaryLabel}
            onMouseDown={(event) => event.stopPropagation()}
          >
            {actions.map((action, index) => (
              <button
                key={action.id}
                type="button"
                role="menuitem"
                className={`editor-mode-ai-item${index === active ? ' is-active' : ''}`}
                data-tip={action.tip ?? action.label}
                aria-label={action.tip ?? action.label}
                onMouseDown={keepFocusStatic}
                onMouseEnter={() => setActive(index)}
                onClick={() => runAt(index)}
                onKeyDown={(event) => {
                  if (event.key === 'ArrowDown') {
                    event.preventDefault()
                    setActive((index + 1) % actions.length)
                  } else if (event.key === 'ArrowUp') {
                    event.preventDefault()
                    setActive((index - 1 + actions.length) % actions.length)
                  } else if (event.key === 'Enter' || event.key === ' ') {
                    event.preventDefault()
                    runAt(index)
                  } else if (event.key === 'Escape') {
                    event.preventDefault()
                    setOpen(false)
                  }
                }}
              >
                <IconModeAi size={14} />
                <span className="editor-mode-ai-item-label">{action.label}</span>
              </button>
            ))}
          </div>
        )}
      </span>
    </>
  )
}

/** Module-level twin of the bar's keepFocus (the menu renders outside it). */
const keepFocusStatic = (event: ReactMouseEvent<HTMLButtonElement>) => event.preventDefault()
