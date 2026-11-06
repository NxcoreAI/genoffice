/**
 * Context menu — fixed positioning, kept within the viewport; closes on any click/Escape/scroll.
 * Items that are null render as separators.
 *
 * Long menus stay a fixed size: the list scrolls past MAX_VISIBLE_ITEMS rows
 * instead of growing off-screen. `children` on an item opens a hover submenu
 * (used to fold rare arrange/align groups out of the top level).
 */
import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { platformShortcuts } from '@genoffice/i18n'

export interface CtxItem {
  label: string
  /** Shortcut hint on the right */
  hint?: string
  disabled?: boolean
  danger?: boolean
  onClick?: () => void
  /** Row of color swatches rendered under the label instead of a clickable item ('none' = the clear square) */
  swatches?: string[]
  onSwatch?: (color: string) => void
  /** Nested items rendered in a hover submenu; replaces onClick */
  children?: Array<CtxItem | null>
}

interface Props {
  x: number
  y: number
  items: Array<CtxItem | null>
  onClose: () => void
}

/** Past this many rows the list scrolls instead of growing (fixed-size menu) */
const MAX_VISIBLE_ITEMS = 12
const ROW_HEIGHT = 32

export function ContextMenu({ x, y, items, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ x, y })
  const [submenu, setSubmenu] = useState<number | null>(null)

  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const { innerWidth, innerHeight } = window
    const r = el.getBoundingClientRect()
    setPos({
      x: Math.min(x, Math.max(0, innerWidth - r.width - 4)),
      y: Math.min(y, Math.max(0, innerHeight - r.height - 4)),
    })
  }, [x, y])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('mousedown', onDown)
    window.addEventListener('wheel', onClose, { once: true })
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('wheel', onClose)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  const swatchRow = (swatches: string[], onSwatch?: (color: string) => void) => (
    <div className="ctx-swatches-row">
      {swatches.map((c) => (
        <button
          key={c}
          className={`ctx-swatch ${c === 'none' ? 'ctx-swatch-none' : ''}`}
          style={c === 'none' ? undefined : { background: c }}
          data-tip={c}
          aria-label={c}
          onClick={() => {
            onClose()
            onSwatch?.(c)
          }}
        >
          {c === 'none' ? '✕' : ''}
        </button>
      ))}
    </div>
  )

  const row = (item: CtxItem, key: number, inSubmenu = false): React.JSX.Element => {
    if (item.children) {
      return (
        <div
          key={key}
          className="ctx-item-wrap"
          onMouseEnter={() => {
            if (!inSubmenu) setSubmenu(key)
          }}
          onMouseLeave={() => setSubmenu((current) => (current === key ? null : current))}
        >
          <button className="ctx-item" disabled={item.disabled}>
            <span>{item.label}</span>
            <span className="ctx-arrow">›</span>
          </button>
          {submenu === key && !item.disabled && (
            <div className="ctx-submenu">
              {item.children.map((child, i) =>
                child === null ? (
                  <div key={i} className="ctx-sep" />
                ) : (
                  row(child, i, true)
                ),
              )}
            </div>
          )}
        </div>
      )
    }
    return (
      <button
        key={key}
        className={`ctx-item ${item.danger ? 'danger' : ''}`}
        disabled={item.disabled}
        onClick={() => {
          onClose()
          item.onClick?.()
        }}
      >
        <span>{item.label}</span>
        {item.hint && <span className="ctx-hint">{platformShortcuts(item.hint)}</span>}
      </button>
    )
  }

  // separators and swatch rows don't count toward the visible-row budget;
  // a submenu parent weighs one row (its children render off to the side)
  const visibleWeight = items.reduce((n, item) => n + (item && !item.swatches ? 1 : 0), 0)
  const maxListHeight =
    visibleWeight > MAX_VISIBLE_ITEMS ? MAX_VISIBLE_ITEMS * ROW_HEIGHT + 8 : undefined

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.x, top: pos.y }}
      onContextMenu={(e) => e.preventDefault()}
    >
      <div
        className="ctx-list"
        style={maxListHeight ? { maxHeight: maxListHeight } : undefined}
      >
        {items.map((item, i) =>
          item === null ? (
            <div key={i} className="ctx-sep" />
          ) : item.swatches ? (
            <div key={i} className="ctx-swatches">
              <span className="ctx-swatches-label">{item.label}</span>
              {swatchRow(item.swatches, item.onSwatch)}
            </div>
          ) : (
            row(item, i)
          ),
        )}
      </div>
    </div>
  )
}
