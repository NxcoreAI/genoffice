/**
 * 论文装订引擎：源论文 docx + 格式卡（ThesisFormatCard）+ 装订规格
 * （ThesisBindingSpec）→ 完整装订版 docx。
 *
 * 结构：封面节（无页眉页脚）→ 前置节（承诺书/中英摘要/目录，罗马页码）
 * → 正文节（章节/附录/参考文献/致谢，阿拉伯页码从 1 重编）。
 * 正文内容块尽量走 original（字节保真：表格/图片/公式原样复用），
 * 文本段落重建为 generated 块以套卡内字体字号行距；目录为带缓存条目的
 * TOC 域（w:dirty，Word 打开自动刷新）。
 */

import { generateTocFieldXml } from './generate.js'
import { parseDocx } from './parse.js'
import { saveDocx, type SaveBlock, type SaveOptions, type StyleUpsert } from './patch.js'
import { readSections } from './section.js'
import { PAGE_MARK, type GeneratedBlock, type ParaFormat, type Run, type SectionSettings } from './types.js'
import { loadDocxZip } from './zip-load.js'

const MM_TO_TWIPS = 56.6929
const PT_TO_TWIPS = 20

export interface ThesisFontCard {
  eastAsia: string
  ascii: string
  size: number
  bold: boolean
}

export interface ThesisFormatCardInput {
  page: {
    widthMm: number
    heightMm: number
    orientation: 'portrait' | 'landscape'
    marginTopMm: number
    marginBottomMm: number
    marginLeftMm: number
    marginRightMm: number
    gutterMm: number
  }
  base: {
    font: ThesisFontCard
    paragraph: {
      align: 'left' | 'center' | 'right' | 'justify'
      firstLineIndentChars: number
      lineSpacingMultiple?: number
      spaceBeforePt?: number
      spaceAfterPt?: number
    }
  }
  header: {
    enabled: boolean
    template: string
    font: ThesisFontCard
    underline: boolean
  }
  footer: {
    enabled: boolean
    font: ThesisFontCard
    frontMatterFormat: 'roman' | 'decimal'
  }
  headings: Array<{
    level: 1 | 2 | 3 | 4
    font: ThesisFontCard
    align: 'left' | 'center'
    numbering: 'numeric-tree' | 'none'
    paragraph?: { spaceBeforePt?: number; spaceAfterPt?: number; lineSpacingMultiple?: number }
  }>
  abstract: {
    titleFont: ThesisFontCard
    bodyFont: ThesisFontCard
    bodyParagraph: {
      align: 'left' | 'center' | 'right' | 'justify'
      firstLineIndentChars: number
      lineSpacingMultiple?: number
    }
    keywordPrefix: string
    keywordPrefixBold: boolean
    keywordSeparator: string
  }
  toc: { enabled: boolean; levels: string; titleFont: ThesisFontCard }
  references: {
    listFont: ThesisFontCard
    listParagraph: {
      align: 'left' | 'center' | 'right' | 'justify'
      lineSpacingMultiple?: number
    }
  }
  cover: {
    schoolName: string
    degreeLine: string
    yearLine: string
    fields: Array<{ label: string; key: string }>
    footerLine: string
    schoolFont: ThesisFontCard
    fieldLabelFont: ThesisFontCard
  }
  fixedTexts: { commitment: string[] }
}

export interface ThesisBindingSpecInput {
  title: string
  coverFields: Record<string, string>
  abstractZh: string
  keywordsZh: string[]
  abstractEn: string
  keywordsEn: string[]
  chapters: Array<{ level: 1 | 2 | 3 | 4; headingText: string; blockStart: number; blockEnd: number }>
  references: string[]
  appendixBlocks: { start: number; end: number } | null
  acknowledgementParagraphs: string[]
  achievements: string[]
}

function mm(mmValue: number): number {
  return Math.round(mmValue * MM_TO_TWIPS)
}

function pt(ptValue: number): number {
  return Math.round(ptValue * PT_TO_TWIPS)
}

function halfPoints(sizePt: number): number {
  return Math.round(sizePt * 2)
}

function textRun(card: ThesisFontCard, text: string, bold = card.bold): Run {
  return {
    text,
    font: card.eastAsia,
    fontAscii: card.ascii,
    sizeHalfPoints: halfPoints(card.size),
    bold,
  }
}

function bodyFormat(card: ThesisFormatCardInput): ParaFormat {
  const paragraph = card.base.paragraph
  const format: ParaFormat = { align: paragraph.align }
  if (paragraph.lineSpacingMultiple && paragraph.lineSpacingMultiple > 0) {
    format.lineSpacing = paragraph.lineSpacingMultiple
  }
  const chars = paragraph.firstLineIndentChars
  if (chars > 0) format.indentFirstLine = Math.round(chars * card.base.font.size * PT_TO_TWIPS)
  if (paragraph.spaceBeforePt) format.spaceBefore = pt(paragraph.spaceBeforePt)
  if (paragraph.spaceAfterPt) format.spaceAfter = pt(paragraph.spaceAfterPt)
  return format
}

function generated(runs: Run[], format: ParaFormat, styleId?: string): SaveBlock {
  const block: GeneratedBlock = { type: 'paragraph', runs, format }
  if (styleId) block.styleId = styleId
  return { kind: 'generated', block }
}

function centered(runs: Run[], spaceBeforePt = 0, spaceAfterPt = 0): SaveBlock {
  return generated(runs, {
    align: 'center',
    ...(spaceBeforePt ? { spaceBefore: pt(spaceBeforePt) } : {}),
    ...(spaceAfterPt ? { spaceAfter: pt(spaceAfterPt) } : {}),
  })
}

/** 分节段：空段落的 pPr 内嵌完整 sectPr（页尺寸/页边距/装订线）。 */
function sectionBreakParagraph(page: ThesisFormatCardInput['page']): string {
  return (
    '<w:p><w:pPr>' +
    '<w:sectPr><w:type w:val="nextPage"/>' +
    `<w:pgSz w:w="${mm(page.widthMm)}" w:h="${mm(page.heightMm)}"/>` +
    `<w:pgMar w:top="${mm(page.marginTopMm)}" w:right="${mm(page.marginRightMm)}" ` +
    `w:bottom="${mm(page.marginBottomMm)}" w:left="${mm(page.marginLeftMm)}" ` +
    `w:header="851" w:footer="992" w:gutter="${mm(page.gutterMm)}"/>` +
    '</w:sectPr>' +
    '</w:pPr></w:p>'
  )
}

function sectionSettingsOf(card: ThesisFormatCardInput): SectionSettings {
  const page = card.page
  return {
    pageWidth: mm(page.widthMm),
    pageHeight: mm(page.heightMm),
    orientation: page.orientation,
    marginTop: mm(page.marginTopMm),
    marginRight: mm(page.marginRightMm),
    marginBottom: mm(page.marginBottomMm),
    marginLeft: mm(page.marginLeftMm),
    pageBorder: false,
    columns: 1,
  }
}

function partTitle(runs: Run[]): SaveBlock {
  return generated(runs, { align: 'center', spaceBefore: pt(12), spaceAfter: pt(12) })
}

function chapterNumberText(
  counters: [number, number, number, number],
  level: 1 | 2 | 3 | 4,
): string {
  const parts: string[] = []
  for (let i = 0; i < level; i += 1) parts.push(String(counters[i]))
  return parts.join('.')
}

/** 源块 → 装订块：文本段重建套样式；表格/图片/公式等 passthrough 保真。 */
function adaptBodyBlock(
  block: { docxIndex: number | null; type: string; level?: number; runs?: Run[] },
  card: ThesisFormatCardInput,
): SaveBlock | null {
  if (block.docxIndex === null) return null
  if (block.type === 'paragraph' && block.runs) {
    const runs = block.runs.map((run) =>
      run.image || run.math || run.ruby || run.noteRef || run.link
        ? run
        : {
          ...run,
          font: card.base.font.eastAsia,
          fontAscii: card.base.font.ascii,
          sizeHalfPoints: halfPoints(card.base.font.size),
        },
    )
    return { kind: 'generated', block: { type: 'paragraph', runs, format: bodyFormat(card) } }
  }
  return { kind: 'original', docxIndex: block.docxIndex }
}

export async function describeThesisBlocks(bytes: Uint8Array): Promise<
  Array<{ index: number; kind: 'heading' | 'paragraph' | 'table' | 'image' | 'list'; level?: number; text: string }>
> {
  const parsed = await parseDocx(bytes)
  return parsed.blocks
    .filter((block) => !block.hidden)
    .map((block, index) => ({
      index,
      kind: block.type === 'listItem' ? 'list' : block.type === 'passthrough' ? 'table' : block.type,
      ...(block.level !== undefined ? { level: block.level } : {}),
      text: (block.runs?.map((run) => run.text).join('') ?? block.previewText ?? block.label ?? '')
        .slice(0, 200)
        .trim(),
    }))
}

export async function buildThesisDocx(
  sourceBytes: Uint8Array,
  card: ThesisFormatCardInput,
  spec: ThesisBindingSpecInput,
): Promise<Uint8Array> {
  const parsed = await parseDocx(sourceBytes)
  const sections = readSections(parsed)
  if (sections.length > 1) throw new Error('thesis_source_multisection_unsupported')

  const visible = parsed.blocks.filter((block) => !block.hidden)
  const settings = sectionSettingsOf(card)
  const title = spec.title || spec.coverFields.title || ''
  const coverTitle = card.header.template.includes('{title}') && title
    ? card.header.template.replace('{title}', title)
    : card.header.template

  const finalBlocks: SaveBlock[] = []

  // ---- 章节号预计算（目录缓存条目与标题块共用同一编号）----
  const counters: [number, number, number, number] = [0, 0, 0, 0]
  const tocEntries: Array<{ text: string; level: number }> = []
  const headingStyleIds = ['Heading1', 'Heading2', 'Heading3', 'Heading4']
  const headingByLevel = new Map(card.headings.map((heading) => [heading.level, heading]))
  for (const chapter of spec.chapters) {
    const heading = headingByLevel.get(chapter.level) ?? headingByLevel.get(1)!
    if (heading.numbering === 'numeric-tree') {
      const idx = chapter.level - 1
      counters[idx] += 1
      for (let i = idx + 1; i < 4; i += 1) counters[i] = 0
    }
    const numberText = heading.numbering === 'numeric-tree'
      ? `${chapterNumberText(counters, chapter.level)} `
      : ''
    tocEntries.push({ text: `${numberText}${chapter.headingText}`, level: Math.min(chapter.level, 3) })
  }
  const tocMaxLevel = Number.parseInt(card.toc.levels.replace(/[^0-9]/g, '').slice(-1), 10)
  const tocFieldXml = card.toc.enabled && tocEntries.length > 0
    ? generateTocFieldXml(
      tocEntries,
      Number.isFinite(tocMaxLevel) && tocMaxLevel >= 1 && tocMaxLevel <= 9 ? tocMaxLevel : undefined,
    ).join('')
    : ''

  // ---- 封面节 ----
  finalBlocks.push(centered([textRun(card.cover.schoolFont, card.cover.schoolName)], 6, 12))
  if (card.cover.degreeLine) {
    finalBlocks.push(centered([textRun(card.cover.schoolFont, card.cover.degreeLine)], 0, 6))
  }
  if (card.cover.yearLine) {
    finalBlocks.push(centered([textRun(card.cover.fieldLabelFont, card.cover.yearLine)], 0, 24))
  }
  for (const field of card.cover.fields) {
    const value = spec.coverFields[field.key] ?? ''
    finalBlocks.push(centered([
      textRun(card.cover.fieldLabelFont, `${field.label}：`),
      textRun(card.cover.fieldLabelFont, value),
    ], 6, 6))
  }
  if (card.cover.footerLine) {
    finalBlocks.push(centered([textRun(card.cover.fieldLabelFont, card.cover.footerLine)], 36, 0))
  }
  finalBlocks.push({ kind: 'xml', xml: sectionBreakParagraph(card.page) })

  // ---- 前置节：承诺书 → 中文摘要 → 英文摘要 → 目录 ----
  if (card.fixedTexts.commitment.length > 0) {
    finalBlocks.push(partTitle([textRun(card.abstract.titleFont, '毕业论文（设计）诚信承诺书')]))
    for (const paragraph of card.fixedTexts.commitment) {
      finalBlocks.push(generated([textRun(card.base.font, paragraph)], bodyFormat(card)))
    }
  }
  if (spec.abstractZh.trim() || spec.keywordsZh.length > 0) {
    finalBlocks.push(partTitle([textRun(card.abstract.titleFont, '摘要')]))
    if (spec.abstractZh.trim()) {
      finalBlocks.push(generated([textRun(card.abstract.bodyFont, spec.abstractZh.trim())], {
        align: card.abstract.bodyParagraph.align,
        ...(card.abstract.bodyParagraph.lineSpacingMultiple
          ? { lineSpacing: card.abstract.bodyParagraph.lineSpacingMultiple }
          : {}),
        indentFirstLine: Math.round(
          card.abstract.bodyParagraph.firstLineIndentChars * card.abstract.bodyFont.size * PT_TO_TWIPS,
        ),
      }))
    }
    if (spec.keywordsZh.length > 0) {
      finalBlocks.push(generated([
        textRun(card.abstract.bodyFont, card.abstract.keywordPrefix, card.abstract.keywordPrefixBold),
        textRun(card.abstract.bodyFont, spec.keywordsZh.join(card.abstract.keywordSeparator)),
      ], { ...bodyFormat(card), indentFirstLine: 0 }))
    }
  }
  if (spec.abstractEn.trim() || spec.keywordsEn.length > 0) {
    if (title) {
      finalBlocks.push(partTitle([textRun(card.abstract.titleFont, title)]))
    }
    finalBlocks.push(partTitle([textRun(card.abstract.titleFont, 'Abstract')]))
    if (spec.abstractEn.trim()) {
      finalBlocks.push(generated([textRun(card.abstract.bodyFont, spec.abstractEn.trim())], {
        align: card.abstract.bodyParagraph.align,
        ...(card.abstract.bodyParagraph.lineSpacingMultiple
          ? { lineSpacing: card.abstract.bodyParagraph.lineSpacingMultiple }
          : {}),
      }))
    }
    if (spec.keywordsEn.length > 0) {
      finalBlocks.push(generated([
        textRun(card.abstract.bodyFont, 'Keywords: ', card.abstract.keywordPrefixBold),
        textRun(card.abstract.bodyFont, spec.keywordsEn.join(', ')),
      ], { ...bodyFormat(card), indentFirstLine: 0 }))
    }
  }
  if (card.toc.enabled) {
    finalBlocks.push(partTitle([textRun(card.toc.titleFont, '目录')]))
    finalBlocks.push({ kind: 'xml', xml: tocFieldXml })
  }
  finalBlocks.push({ kind: 'xml', xml: sectionBreakParagraph(card.page) })

  // ---- 正文节 ----
  for (const [chapterIndex, chapter] of spec.chapters.entries()) {
    const heading = headingByLevel.get(chapter.level) ?? headingByLevel.get(1)!
    const headingText = tocEntries[chapterIndex]!.text
    const headingFormat: ParaFormat = {
      align: heading.align,
      ...(heading.paragraph?.lineSpacingMultiple
        ? { lineSpacing: heading.paragraph.lineSpacingMultiple }
        : {}),
      ...(heading.paragraph?.spaceBeforePt ? { spaceBefore: pt(heading.paragraph.spaceBeforePt) } : {}),
      ...(heading.paragraph?.spaceAfterPt ? { spaceAfter: pt(heading.paragraph.spaceAfterPt) } : {}),
    }
    finalBlocks.push({
      kind: 'generated',
      block: {
        type: 'heading',
        level: chapter.level,
        styleId: headingStyleIds[chapter.level - 1],
        runs: [textRun(heading.font, headingText)],
        format: headingFormat,
      },
    })
    for (let i = chapter.blockStart + 1; i < chapter.blockEnd; i += 1) {
      const block = visible[i]
      if (!block) continue
      if (block.originalXml?.includes('<w:sectPr')) continue
      const adapted = adaptBodyBlock(block, card)
      if (adapted) finalBlocks.push(adapted)
    }
  }

  // ---- 附录（复用源块）----
  if (spec.appendixBlocks) {
    finalBlocks.push(partTitle([textRun(card.abstract.titleFont, '附录')]))
    for (let i = spec.appendixBlocks.start; i < spec.appendixBlocks.end; i += 1) {
      const block = visible[i]
      if (!block || block.docxIndex === null) continue
      if (block.originalXml?.includes('<w:sectPr')) continue
      finalBlocks.push({ kind: 'original', docxIndex: block.docxIndex })
    }
  }

  // ---- 参考文献 ----
  finalBlocks.push(partTitle([textRun(card.abstract.titleFont, '参考文献')]))
  const refFormat: ParaFormat = { align: card.references.listParagraph.align }
  if (card.references.listParagraph.lineSpacingMultiple) {
    refFormat.lineSpacing = card.references.listParagraph.lineSpacingMultiple
  }
  for (const entry of spec.references) {
    finalBlocks.push(generated([textRun(card.references.listFont, entry)], refFormat))
  }

  // ---- 致谢 / 研究成果 ----
  finalBlocks.push(partTitle([textRun(card.abstract.titleFont, '致谢')]))
  for (const paragraph of spec.acknowledgementParagraphs) {
    finalBlocks.push(generated([textRun(card.base.font, paragraph)], bodyFormat(card)))
  }
  if (spec.achievements.length > 0) {
    finalBlocks.push(partTitle([textRun(card.abstract.titleFont, '攻读学位期间的研究成果')]))
    for (const entry of spec.achievements) {
      finalBlocks.push(generated([textRun(card.references.listFont, entry)], refFormat))
    }
  }

  // ---- 保存选项 ----
  const styleUpserts: StyleUpsert[] = []
  for (const heading of card.headings) {
    styleUpserts.push({
      styleId: headingStyleIds[heading.level - 1],
      type: 'paragraph',
      name: `heading ${heading.level}`,
      basedOn: 'Normal',
      rPr: {
        bold: heading.font.bold,
        font: heading.font.eastAsia,
        sizeHalfPoints: halfPoints(heading.font.size),
      },
      pPr: {
        align: heading.align,
        ...(heading.paragraph?.lineSpacingMultiple
          ? { lineSpacing: heading.paragraph.lineSpacingMultiple }
          : {}),
      },
    })
  }
  for (let level = 1; level <= 3; level += 1) {
    styleUpserts.push({
      styleId: `TOC${level}`,
      type: 'paragraph',
      name: `toc ${level}`,
      basedOn: 'Normal',
      rPr: { font: card.base.font.eastAsia, sizeHalfPoints: halfPoints(card.base.font.size) },
    })
  }

  const options: SaveOptions = {
    section: settings,
    pgNumType: { fmt: 'decimal', start: 1 },
    hfAllSections: true,
    styleUpserts,
  }
  if (card.header.enabled) {
    options.header = {
      text: coverTitle,
      paras: [{
        align: 'center',
        ...(card.header.underline ? { borders: 'b', borderStyle: { szEighths: 4 } } : {}),
        runs: [textRun(card.header.font, coverTitle)],
      }],
    }
  }
  if (card.footer.enabled) {
    options.footer = {
      text: '',
      pageNumber: true,
      paras: [{
        align: 'center',
        runs: [{
          text: PAGE_MARK,
          font: card.footer.font.eastAsia,
          fontAscii: card.footer.font.ascii,
          sizeHalfPoints: halfPoints(card.footer.font.size),
          bold: card.footer.font.bold,
        }],
      }],
    }
  }

  const saved = await saveDocx(parsed, finalBlocks, options)
  return postProcessSections(saved, card)
}

/**
 * 后处理（saveDocx 之后的字符串手术）：
 * 1. 封面节（第一个内嵌 sectPr）剥掉页眉页脚引用——封面无页眉页脚。
 * 2. 前置节（第二个内嵌 sectPr）注入罗马页码格式。
 * 3. TOC 域保持 w:dirty（generateTocFieldXml 已带），Word 打开自动刷新。
 */
async function postProcessSections(bytes: Uint8Array, card: ThesisFormatCardInput): Promise<Uint8Array> {
  const zip = await loadDocxZip(bytes)
  const documentFile = zip.file('word/document.xml')
  if (!documentFile) return bytes
  const xml = await documentFile.async('string')
  const sectMatches = [...xml.matchAll(/<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>/g)]
  if (sectMatches.length < 3) return bytes

  const coverSect = sectMatches[0]![0]
  const frontSect = sectMatches[1]![0]

  let next = xml
  // 封面：剥页眉页脚引用
  const coverStripped = coverSect.replace(/<w:headerReference[^>]*\/>|<w:footerReference[^>]*\/>/g, '')
  if (coverStripped !== coverSect) next = next.replace(coverSect, coverStripped)

  // 前置节：罗马页码
  if (card.footer.frontMatterFormat === 'roman') {
    const frontWithNum = frontSect.includes('<w:pgNumType')
      ? frontSect.replace(/<w:pgNumType[^>]*\/>/, '<w:pgNumType w:fmt="upperRoman"/>')
      : frontSect.replace('</w:sectPr>', '<w:pgNumType w:fmt="upperRoman"/></w:sectPr>')
    if (frontWithNum !== frontSect) next = next.replace(frontSect, frontWithNum)
  }

  if (next === xml) return bytes
  zip.file('word/document.xml', next)
  return zip.generateAsync({ type: 'nodebuffer' })
}
