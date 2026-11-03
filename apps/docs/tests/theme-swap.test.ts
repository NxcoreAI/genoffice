import { describe, expect, it } from 'vitest'
import type { ParsedDocFull, StyleInfo, ThemeColors, ThemeFonts } from '@genoffice/docx-engine'
import { applyPaletteToHeadingStyles, swapThemeInParsed } from '../src/renderer/theme-swap'

const OLD_FONTS: ThemeFonts = {
  major: '等线 Light',
  minor: 'Times New Roman',
  eastAsia: '宋体',
  majorEastAsia: '等线 Light',
}
const NEW_FONTS: ThemeFonts = {
  major: '黑体',
  minor: 'Arial',
  eastAsia: '微软雅黑',
  majorEastAsia: '黑体',
}
const OLD_COLORS: ThemeColors = { name: 'Office', accent1: '4472C4', dk2: '44546A' }
const NEW_COLORS: ThemeColors = { name: 'Facet', accent1: '90C226', dk2: '3E3D2D' }

function style(styleId: string, display: StyleInfo['display'], headingLevel?: number): StyleInfo {
  return { styleId, name: styleId, type: 'paragraph', display, ...(headingLevel ? { headingLevel } : {}) }
}

function parsedFixture(): ParsedDocFull {
  return {
    styles: new Map<string, StyleInfo>([
      ['Normal', style('Normal', { font: '宋体', fontAscii: 'Times New Roman' })],
      ['Heading1', style('Heading1', { font: '等线 Light', fontAscii: '等线 Light', color: '4472C4' }, 1)],
      ['Heading2', style('Heading2', { color: '000000' }, 2)],
      ['Title', style('Title', { color: '000000' })],
      ['Quote', style('Quote', { font: '仿宋', fontAscii: 'Georgia' })],
    ]),
    docDefaults: { asciiFont: 'Times New Roman', eastAsiaFont: '宋体' },
  } as unknown as ParsedDocFull
}

describe('swapThemeInParsed', () => {
  it('moves theme faces of styles and defaults to the new theme', () => {
    const parsed = parsedFixture()
    const next = swapThemeInParsed(parsed, OLD_FONTS, NEW_FONTS, null, null)

    const normal = next.styles.get('Normal')!.display!
    expect(normal.font).toBe('微软雅黑')
    expect(normal.fontAscii).toBe('Arial')

    const heading = next.styles.get('Heading1')!.display!
    expect(heading.font).toBe('黑体')
    expect(heading.fontAscii).toBe('黑体')

    // explicit, non-theme formatting never moves
    const quote = next.styles.get('Quote')!.display!
    expect(quote.font).toBe('仿宋')
    expect(quote.fontAscii).toBe('Georgia')

    expect(next.docDefaults?.asciiFont).toBe('Arial')
    expect(next.docDefaults?.eastAsiaFont).toBe('微软雅黑')
  })

  it('maps theme palette slots by key', () => {
    const parsed = parsedFixture()
    const next = swapThemeInParsed(parsed, null, null, OLD_COLORS, NEW_COLORS)
    expect(next.styles.get('Heading1')!.display!.color).toBe('90C226')
  })

  it('returns the same parsed object when nothing references the old theme', () => {
    const parsed = parsedFixture()
    const untouched = swapThemeInParsed(
      parsed,
      { major: 'Calibri Light', minor: 'Calibri' },
      { major: '黑体', minor: 'Arial' },
      null,
      null,
    )
    expect(untouched).toBe(parsed)
  })
})

describe('applyPaletteToHeadingStyles', () => {
  it('pins heading-level styles to the new accent1 even when they were black', () => {
    const parsed = parsedFixture()
    const next = applyPaletteToHeadingStyles(parsed, { accent1: '90C226' })
    expect(next.styles.get('Heading2')!.display!.color).toBe('90C226')
    expect(next.styles.get('Title')!.display!.color).toBe('90C226')
    // body / non-heading styles keep their own color
    expect(next.styles.get('Normal')!.display!.color).toBeUndefined()
    expect(next.styles.get('Quote')!.display!.color).toBeUndefined()
  })

  it('returns the same parsed object when headings already use the accent', () => {
    const parsed = parsedFixture()
    const once = applyPaletteToHeadingStyles(parsed, { accent1: '90C226' })
    expect(applyPaletteToHeadingStyles(once, { accent1: '90C226' })).toBe(once)
    expect(applyPaletteToHeadingStyles(parsed, null)).toBe(parsed)
  })
})
