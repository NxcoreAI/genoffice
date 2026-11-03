import type { MouseEvent as ReactMouseEvent } from 'react'

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
      {showAi && (
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
      )}
    </div>
  )
}
