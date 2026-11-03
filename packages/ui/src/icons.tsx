import type { ReactNode } from 'react'

export interface IconProps {
  size?: number
}

function Svg({ size = 16, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.2 8 13.8 2.6 11 13.4 7.6 9.6z" strokeLinejoin="round" />
      <path d="M7.6 9.6 13.8 2.6" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4" y="4" width="8" height="8" rx="1.5" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/** return/enter arrow (↵) for the icon-only send button */
export function IconEnter(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M13 3.5v4a2.5 2.5 0 0 1-2.5 2.5H3.5" />
      <path d="M6.5 7 3.5 10l3 3" />
    </Svg>
  )
}

/** Operation-panel toggle: a panel frame with a chevron that points at its edge.
 *  `dir='collapse'` signals "hide the panel" (normal → AI mode), `'expand'`
 *  signals "bring the full panel back" (AI → normal mode). */
export function IconPanelToggle({ size = 16, dir = 'collapse' }: IconProps & { dir?: 'collapse' | 'expand' }) {
  return (
    <Svg size={size}>
      <rect x="1.5" y="2.6" width="13" height="10.8" rx="1.6" />
      <path d="M5.6 2.6v10.8" />
      {dir === 'collapse' ? (
        <path d="M11 6.1 9.1 8 11 9.9" />
      ) : (
        <path d="M9.1 6.1 11 8l-1.9 1.9" />
      )}
    </Svg>
  )
}

/** Compact-bar save glyph (floppy outline, toolbars are 15–16px). */
export function IconModeSave({ size = 15 }: IconProps) {
  return (
    <Svg size={size}>
      <path d="M2.6 3.4a.8.8 0 0 1 .8-.8h7.4l2.6 2.6V12.6a.8.8 0 0 1-.8.8H3.4a.8.8 0 0 1-.8-.8z" />
      <path d="M5.4 2.6v3.2h4.2V2.6" />
      <path d="M5.4 13.4V9.6h5.2v3.8" />
    </Svg>
  )
}

/** Compact-bar undo glyph. */
export function IconModeUndo({ size = 15 }: IconProps) {
  return (
    <Svg size={size}>
      <path d="M5.2 4.2 2.4 7l2.8 2.8" />
      <path d="M2.4 7h6.2a3.4 3.4 0 0 1 0 6.8H6" />
    </Svg>
  )
}

/** Compact-bar redo glyph. */
export function IconModeRedo({ size = 15 }: IconProps) {
  return (
    <Svg size={size}>
      <path d="M10.8 4.2 13.6 7l-2.8 2.8" />
      <path d="M13.6 7H7.4a3.4 3.4 0 0 0 0 6.8H10" />
    </Svg>
  )
}

/** Compact-bar AI action glyph (four-point sparkle, matches the AI entry tone). */
export function IconModeAi({ size = 15 }: IconProps) {
  return (
    <Svg size={size}>
      <path d="M8 1.8 9.5 6.5 14.2 8 9.5 9.5 8 14.2 6.5 9.5 1.8 8 6.5 6.5z" fill="currentColor" stroke="none" />
      <path d="M12.9 1.6 13.5 3.3 15.2 3.9 13.5 4.5 12.9 6.2 12.3 4.5 10.6 3.9 12.3 3.3z" fill="currentColor" stroke="none" />
    </Svg>
  )
}
