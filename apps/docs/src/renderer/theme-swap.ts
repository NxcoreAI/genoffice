import type {
  DocDefaults,
  ParsedDocFull,
  StyleDisplay,
  StyleInfo,
  ThemeColors,
  ThemeFonts,
} from '@genoffice/docx-engine'

/**
 * Design ▸ Themes / Fonts / Colors 的“编辑器内即时生效”。
 *
 * 主题字体/主题色在解析 docx 时就已经落进 styles.xml 的 display 提示里
 * （Normal 的正文字体、Heading 的标题字体、引用主题色的样式颜色），页面上的
 * 正文/标题 CSS 全部读这份 display 数据。只改 `themeFonts` / `themeColors`
 * 状态的话，画布不会变——用户看到的就是“点了主题/字体/颜色没有任何反应”。
 *
 * 这里按 Word 的语义做“样式级换肤”：把 display 里等于旧主题字体/旧主题色的
 * 值替换成新主题的对应值（major→major、minor→minor、accent1→accent1 …），
 * 显式格式（run 自己的直接格式）与主题无关，保持不动。
 *
 * 只动显示层（display），保存路径仍然只写 theme1.xml，styles.xml 不重写。
 */

const normFace = (value: string | undefined | null): string => (value ?? '').trim().toLowerCase()

const normHex = (value: string | undefined | null): string =>
  (value ?? '').replace(/^#/, '').trim().toUpperCase()

/** 旧主题字体 → 新主题字体；不在旧主题里的字体（显式格式）原样返回 */
function swapFace(
  face: string | undefined,
  from: ThemeFonts | null | undefined,
  to: ThemeFonts | null | undefined,
): string | undefined {
  if (!face || !from || !to) return face
  const current = normFace(face)
  const pairs: Array<[string | undefined, string | undefined]> = [
    [from.major, to.major],
    [from.minor, to.minor],
    [from.majorEastAsia, to.majorEastAsia],
    [from.eastAsia, to.eastAsia],
    [from.majorCs, to.majorCs],
    [from.minorCs, to.minorCs],
  ]
  for (const [oldFace, newFace] of pairs) {
    if (oldFace && newFace && normFace(oldFace) === current) return newFace
  }
  return face
}

const COLOR_KEYS: Array<keyof ThemeColors> = [
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]

/** 旧主题色 → 新主题色（按同一个槽位）；非主题色原样返回 */
function swapColor(
  color: string | undefined,
  from: ThemeColors | null | undefined,
  to: ThemeColors | null | undefined,
): string | undefined {
  if (!color || !from || !to) return color
  const current = normHex(color)
  for (const key of COLOR_KEYS) {
    const oldColor = from[key]
    const newColor = to[key]
    if (typeof oldColor === 'string' && typeof newColor === 'string' && normHex(oldColor) === current) {
      return newColor
    }
  }
  return color
}

function swapDisplay(
  display: StyleDisplay,
  fromFonts: ThemeFonts | null | undefined,
  toFonts: ThemeFonts | null | undefined,
  fromColors: ThemeColors | null | undefined,
  toColors: ThemeColors | null | undefined,
): StyleDisplay {
  const font = swapFace(display.font, fromFonts, toFonts)
  const fontAscii = swapFace(display.fontAscii, fromFonts, toFonts)
  const csFont = swapFace(display.csFont, fromFonts, toFonts)
  const color = swapColor(display.color, fromColors, toColors)
  if (
    font === display.font &&
    fontAscii === display.fontAscii &&
    csFont === display.csFont &&
    color === display.color
  ) {
    return display
  }
  return { ...display, font, fontAscii, csFont, color }
}

/**
 * 把已解析文档的样式提示按旧主题 → 新主题换一遍。
 * 没有任何命中的样式保持原对象引用，整体无变化时返回原 `parsed`。
 */
export function swapThemeInParsed(
  parsed: ParsedDocFull,
  fromFonts: ThemeFonts | null | undefined,
  toFonts: ThemeFonts | null | undefined,
  fromColors: ThemeColors | null | undefined,
  toColors: ThemeColors | null | undefined,
): ParsedDocFull {
  if (!fromFonts && !fromColors) return parsed
  let stylesChanged = false
  const styles = new Map<string, StyleInfo>()
  for (const [styleId, style] of parsed.styles) {
    if (!style.display) {
      styles.set(styleId, style)
      continue
    }
    const display = swapDisplay(style.display, fromFonts, toFonts, fromColors, toColors)
    if (display === style.display) {
      styles.set(styleId, style)
      continue
    }
    styles.set(styleId, { ...style, display })
    stylesChanged = true
  }

  let defaults = parsed.docDefaults
  let defaultsChanged = false
  if (defaults) {
    const next: DocDefaults = { ...defaults }
    const asciiFont = swapFace(defaults.asciiFont, fromFonts, toFonts)
    const eastAsiaFont = swapFace(defaults.eastAsiaFont, fromFonts, toFonts)
    const color = swapColor(defaults.color, fromColors, toColors)
    if (asciiFont !== defaults.asciiFont) {
      next.asciiFont = asciiFont
      defaultsChanged = true
    }
    if (eastAsiaFont !== defaults.eastAsiaFont) {
      next.eastAsiaFont = eastAsiaFont
      defaultsChanged = true
    }
    if (color !== defaults.color) {
      next.color = color
      defaultsChanged = true
    }
    if (defaultsChanged) defaults = next
  }

  if (!stylesChanged && !defaultsChanged) return parsed
  return { ...parsed, styles, docDefaults: defaults }
}

/** Heading-ish styles whose color the palette should drive (Word's own templates
 *  color Heading 1–9 / Title / Subtitle with accent1). */
function isHeadingStyle(style: StyleInfo): boolean {
  if ((style.headingLevel ?? 0) > 0) return true
  return style.styleId === 'Title' || style.styleId === 'Subtitle'
    || style.name === 'Title' || style.name === 'Subtitle'
}

/**
 * Design ▸ 颜色 的可见效果：把标题级样式（Heading 1–9 / Title / Subtitle）的
 * 文字颜色统一指向新调色板的 accent1。
 *
 * Word 只改「引用主题色」的文字，而中文模板里的标题常被写成固定的黑色，
 * 于是切换调色板在画面上毫无变化（用户会以为按钮坏了）。这里按用户预期把
 * 标题级样式直接绑到新主题的强调色上——只动显示层，正文与其他样式不动。
 */
export function applyPaletteToHeadingStyles(
  parsed: ParsedDocFull,
  colors: ThemeColors | null | undefined,
): ParsedDocFull {
  const accent = colors?.accent1
  if (!accent) return parsed
  let changed = false
  const styles = new Map<string, StyleInfo>()
  for (const [styleId, style] of parsed.styles) {
    if (!isHeadingStyle(style)) {
      styles.set(styleId, style)
      continue
    }
    if (style.display?.color === accent) {
      styles.set(styleId, style)
      continue
    }
    styles.set(styleId, { ...style, display: { ...(style.display ?? {}), color: accent } })
    changed = true
  }
  return changed ? { ...parsed, styles } : parsed
}
