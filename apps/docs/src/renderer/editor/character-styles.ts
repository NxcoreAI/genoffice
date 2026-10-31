import type { CSSProperties } from 'react'
import type { Editor } from '@tiptap/core'
import type { ThemeColors } from '@genoffice/docx-engine'
import type { StringKey } from '../i18n/locale'

/**
 * Character-style gallery definitions shared by the ribbon's style gallery and
 * the host agent's `format.applyCharacterStyle` operation (single source).
 *
 * The two built-in presets are shown when the document has no character
 * styles of its own: they are plain italic/bold marks plus the theme accent
 * color (no w:rStyle), which is exactly what the ribbon applies.
 */

export interface CharStylePreset {
  /** stable agent-facing id used in the host vocabulary */
  id: 'emphasis' | 'strong'
  /** sentinel styleId used by the ribbon gallery key (`char:__preset_emphasis`) */
  styleId: string
  labelKey: StringKey
  mark: 'italic' | 'bold'
  display: (accentHex: string) => CSSProperties
}

/** Theme accent used by the preset character styles when the doc theme has none */
export const DEFAULT_PRESET_ACCENT = '4472C4'

export const CHAR_STYLE_PRESETS: CharStylePreset[] = [
  {
    id: 'emphasis',
    styleId: '__preset_emphasis',
    labelKey: 'ribbonStyleEmphasis',
    mark: 'italic',
    display: (accentHex) => ({ fontStyle: 'italic', color: `#${accentHex}` }),
  },
  {
    id: 'strong',
    styleId: '__preset_strong',
    labelKey: 'ribbonStyleIntenseEmphasis',
    mark: 'bold',
    display: (accentHex) => ({ fontWeight: 'bold', color: `#${accentHex}` }),
  },
]

/** Accent color used by the preset marks (uppercased hex, no '#'). */
export function presetAccentOf(themeColors: ThemeColors | null | undefined): string {
  return themeColors?.accent1?.trim().toUpperCase() || DEFAULT_PRESET_ACCENT
}

export function charStylePresetById(id: string): CharStylePreset | null {
  return CHAR_STYLE_PRESETS.find((preset) => preset.id === id || preset.styleId === id) ?? null
}

/**
 * Whether the current selection already shows this character style, using the
 * same signals as the ribbon gallery (active mark + accent color for presets,
 * docTextStyle.styleId for document character styles).
 */
export function selectionHasCharStyle(
  editor: Editor,
  target: { preset?: CharStylePreset; styleId?: string; accent?: string },
): boolean {
  const styleAttrs = editor.getAttributes('docTextStyle') as { styleId?: string | null; color?: string | null }
  if (target.preset) {
    return (
      editor.isActive(target.preset.mark)
      && !styleAttrs.styleId
      && (styleAttrs.color ?? '').toUpperCase() === (target.accent ?? DEFAULT_PRESET_ACCENT)
    )
  }
  return !!target.styleId && styleAttrs.styleId === target.styleId
}

/**
 * Run the gallery's character-style application on the current selection.
 * Returns true when a mark/attribute chain ran.
 */
export function applyCharStyleToSelection(
  editor: Editor,
  target: { preset?: CharStylePreset; styleId?: string; accent?: string },
  mode: 'apply' | 'toggle',
): boolean {
  const active = selectionHasCharStyle(editor, target)
  if (target.preset) {
    const preset = target.preset
    if (mode === 'toggle' && active) {
      return editor.chain().focus().unsetMark(preset.mark).setMark('docTextStyle', { color: null }).run()
    }
    // switching presets drops the other preset's mark so the gallery highlight
    // cannot stick on two cards (same rule as the ribbon)
    let chain = editor.chain().focus()
    for (const other of CHAR_STYLE_PRESETS) {
      if (other !== preset) chain = chain.unsetMark(other.mark)
    }
    return chain
      .setMark(preset.mark)
      .setMark('docTextStyle', { color: target.accent ?? DEFAULT_PRESET_ACCENT })
      .run()
  }
  const styleId = target.styleId
  if (!styleId) return false
  if (mode === 'toggle' && active) {
    return editor.chain().focus().unsetMark('docTextStyle').run()
  }
  return editor.chain().focus().setMark('docTextStyle', { styleId }).run()
}