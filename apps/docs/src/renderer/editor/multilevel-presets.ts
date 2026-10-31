import type { CustomNumberingLevel } from '@genoffice/docx-engine'

/**
 * Multilevel-list gallery presets, shared by the ribbon's Multilevel List
 * dropdown and the host agent's `format.multilevelList` operation so both
 * produce exactly the same numbering definitions.
 *
 * Kept out of Ribbon.tsx (a heavy component module) so the operations layer
 * can import the presets without pulling in the whole ribbon.
 */

/** Bullet rotation used by the bullet-style multilevel preset (level 1..3 glyphs). */
function bulletPresetLevels(glyph: string): CustomNumberingLevel[] {
  const rotation = [glyph, '○', '■']
  return Array.from({ length: 9 }, (_, i) => ({
    numFmt: 'bullet',
    lvlText: rotation[i % 3],
    indentLeft: 720 * (i + 1),
    hanging: 360,
  }))
}

export const MULTILEVEL_LIBRARY: CustomNumberingLevel[][] = [
  // 1. / 1.1. / 1.1.1.
  Array.from({ length: 9 }, (_, i) => ({
    numFmt: 'decimal',
    lvlText: `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')}.`,
    indentLeft: 720 * (i + 1),
    hanging: 432,
  })),
  // Chinese official-document hierarchy: numeral + comma / parenthesized numeral / 1.
  Array.from({ length: 9 }, (_, i): CustomNumberingLevel => {
    if (i === 0)
      return { numFmt: 'chineseCountingThousand', lvlText: '%1、', indentLeft: 720, hanging: 425 }
    if (i === 1)
      return { numFmt: 'chineseCountingThousand', lvlText: '(%2)', indentLeft: 1440, hanging: 425 }
    return { numFmt: 'decimal', lvlText: `%${i + 1}.`, indentLeft: 720 * (i + 1), hanging: 360 }
  }),
  // • / ○ / ■
  bulletPresetLevels('•'),
]

/** Agent-facing preset ids for the gallery entries (index into MULTILEVEL_LIBRARY). */
export const MULTILEVEL_PRESETS: Array<{ id: string; name: string }> = [
  { id: 'decimal', name: '1. / 1.1. / 1.1.1.' },
  { id: 'chinese', name: '一、 / （一） / 1.' },
  { id: 'bullet', name: '• / ○ / ■' },
]

/** Resolve an agent preset id (or index) to the gallery level definitions. */
export function multilevelLevelsFor(
  preset: unknown,
): { levels: CustomNumberingLevel[]; id: string } | { error: string } {
  let index = 0
  if (preset !== undefined && preset !== null && preset !== '') {
    if (typeof preset === 'number') index = preset
    else {
      const byId = MULTILEVEL_PRESETS.findIndex((candidate) => candidate.id === String(preset))
      const byIndex = Number(preset)
      index = byId >= 0 ? byId : Number.isInteger(byIndex) ? byIndex : -1
    }
  }
  if (!Number.isInteger(index) || index < 0 || index >= MULTILEVEL_LIBRARY.length) {
    return {
      error: `preset must be one of ${MULTILEVEL_PRESETS.map((candidate) => candidate.id).join('/')}`,
    }
  }
  return { levels: MULTILEVEL_LIBRARY[index], id: MULTILEVEL_PRESETS[index].id }
}