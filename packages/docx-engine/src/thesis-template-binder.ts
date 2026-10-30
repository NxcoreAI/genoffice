/**
 * 论文装订（模板骨架版 / v3）：**以模版 docx 为底**，只替换内容槽位。
 *
 * 与 v2 的 `buildThesisDocx`（源文档为底、封面/目录/正文全部重建）的根本区别：
 * 模版里的一切——封面、主标题区、核心信息区（冒号对齐/下划线填空）、logo 的绝对
 * 定位、格式卡、复杂表格、固定的教师评价表单、目录域——都按**原 docxIndex 原样
 * 透传**，一个字节都不重写。内容只做两件事：
 *
 *   1. 把「槽位标题」的文字换成源文档对应章节的标题（保留模版标题的 pPr/rPr、
 *      自动编号与制表位，所以「一、」这类编号仍由模版决定）；
 *   2. 把标题之后到下一个标题之间的**占位内容**换成源文档该章节的正文。
 *
 * 槽位标题之外的块（前置表格、评价表、目录、章节之间的固定表单）一律原样保留，
 * 因此不会出现「模版里的表格被搬到目录后面」这类错位——位置从来没人动过。
 *
 * 源内容的表格/图片按需把 media 与 rels 搬进模版包（rId 重映射），
 * 文本段落则套模版正文字体字号重建（避免把源文档的版式带进模版）。
 */

import { generateTocFieldXml } from './generate.js'
import { parseDocx } from './parse.js'
import { saveDocx, type SaveBlock } from './patch.js'
import { patchParagraphTexts } from './text-patch.js'
import { loadDocxZip } from './zip-load.js'
import {
  headingNumberPrefix,
  type ThesisBindingSpecInput,
  type ThesisFormatCardInput,
  type ThesisNumberingScheme,
} from './thesis-binder.js'
import type { Block, ParaFormat, Run } from './types.js'

const PT_TO_TWIPS = 20


/** 删除表格第 columnIndex 列：网格列 + 每行对应的单元格（不处理合并单元格，调用方需先排除）。 */
function removeTableColumn(tableXml: string, columnIndex: number): string {
  const removals: Array<{ start: number; end: number }> = []
  // 1) w:tblGrid 里的第 n 个 gridCol
  const gridOpen = tableXml.indexOf('<w:tblGrid>')
  const gridClose = tableXml.indexOf('</w:tblGrid>', gridOpen + 1)
  if (gridOpen >= 0 && gridClose > gridOpen) {
    const slice = tableXml.slice(gridOpen, gridClose)
    const colRe = /<w:gridCol\b[^>]*\/>/g
    const cols: number[] = []
    let m: RegExpExecArray | null
    while ((m = colRe.exec(slice))) cols.push(gridOpen + m.index)
    const target = cols[columnIndex]
    if (target !== undefined) {
      const end = tableXml.indexOf('/>', target) + 2
      removals.push({ start: target, end })
    }
  }
  // 2) 每行的第 n 个 w:tc（行内没有嵌套表格时才安全，调用方已保证）
  const rowRe = /<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g
  let row: RegExpExecArray | null
  while ((row = rowRe.exec(tableXml))) {
    const rowXml = row[0]
    const cellRe = /<w:tc\b|<\/w:tc>/g
    let cell: RegExpExecArray | null
    let depth = 0
    let seen = 0
    let cellStart = -1
    while ((cell = cellRe.exec(rowXml))) {
      if (cell[0] === '</w:tc>') {
        depth -= 1
        if (depth === 0 && seen === columnIndex + 1 && cellStart >= 0) {
          removals.push({ start: row.index + cellStart, end: row.index + cell.index + cell[0].length })
          break
        }
      } else {
        if (depth === 0) {
          seen += 1
          if (seen === columnIndex + 1) cellStart = cell.index
        }
        depth += 1
      }
    }
  }
  if (removals.length === 0) return tableXml
  removals.sort((a, b) => b.start - a.start)
  let out = tableXml
  for (const removal of removals) out = out.slice(0, removal.start) + out.slice(removal.end)
  return out
}

/** 单元格纯文本（表格拆列时用来判断这一列讲的是什么）。 */
function cellText(cell: { paras?: string[] } | undefined): string {
  return (cell?.paras ?? []).join(' ').trim()
}

/** 模版剖面里渲染器需要的部分（其余字段不参与装订）。 */
export interface ThesisTemplateProfileInput {
  slots: Array<{ id: string; title: string }>
  /** 模版优先模式：目录/标题/模版文字原样保留，内容按小节语义归类。 */
  mode?: string
  /** 模版里的二级小节（1.1 背景…），供面板展示与归属参考。 */
  subSlots?: Array<{ id: string; parentId: string; title: string; order: number }>
}

export interface ThesisTemplateBinderInput {
  templateBytes: Uint8Array
  sourceBytes: Uint8Array
  card: ThesisFormatCardInput & { profile?: ThesisTemplateProfileInput }
  spec: ThesisBindingSpecInput
}

/** 装订诊断：面板据此告诉用户「这次是按模板排的」还是「正文只能堆到文末」。 */
export interface ThesisTemplateBinderDiagnostics {
  /** 模版剖面里的槽位数量（0 = 模版章节骨架没解析出来）。 */
  slots: number
  /** 模版可见块数量。 */
  blocks: number
  /** 实际匹配到的锚点数量。 */
  anchors: number
  /** 正文区起点（目录之后；0 = 没找到目录标题）。 */
  bodyStart: number
  /** 待落位的论文章节数量。 */
  chapters: number
  /** true = 一个章节锚点都没匹配上，正文已整体追加到文末。 */
  appendedOnly: boolean
  /** true = 模版正文认出的章节锚点太少（如「只有目录」的模版），骨架排版降级为内容重建。 */
  templateFirstDowngraded?: boolean
}

export interface ThesisTemplateBinderResult {
  bytes: Uint8Array
  diagnostics: ThesisTemplateBinderDiagnostics
}

/** 装订目标：正文槽位，或模版固定部件（摘要/参考文献/致谢/附录）。 */
type Target =
  | { kind: 'slot'; slotId: string; title: string }
  | { kind: 'abstract_zh' | 'abstract_en' | 'references' | 'acknowledgement' | 'appendix'; title: string }

interface Anchor {
  target: Target
  /** 模版可见块序号（不是 docxIndex）。 */
  blockIndex: number
}

const SPECIAL_TARGETS: Array<{ kind: 'abstract_zh' | 'abstract_en' | 'references' | 'acknowledgement' | 'appendix'; keywords: string[] }> = [
  { kind: 'abstract_zh', keywords: ['摘要', '中文摘要'] },
  { kind: 'abstract_en', keywords: ['abstract', '英文摘要'] },
  { kind: 'references', keywords: ['参考文献', '参考资料'] },
  { kind: 'acknowledgement', keywords: ['致谢', '謝辞'] },
  { kind: 'appendix', keywords: ['附录', '附錄'] },
]

function pt(value: number): number {
  return Math.round(value * PT_TO_TWIPS)
}

/**
 * 标题是不是「模版固定部件」（摘要/参考文献/致谢/附录/目录）。
 * 占位区里的示例子标题（「1.1 背景」）不是固定部件，要随示例内容一起丢掉；
 * 这几个则必须留着——它们是模版结构本身。
 */
function isFixedPartTitle(text: string): boolean {
  const key = norm(text)
  if (!key) return false
  if (/^(目录|目次|contents)$/.test(key)) return true
  return SPECIAL_TARGETS.some((entry) => entry.keywords.some((keyword) => {
    const target = norm(keyword)
    return target.length >= 2 && (key === target || key.startsWith(target))
  }))
}

function halfPoints(sizePt: number): number {
  return Math.round(sizePt * 2)
}

function textRun(card: ThesisFormatCardInput, text: string, bold = card.base.font.bold): Run {
  return {
    text,
    font: card.base.font.eastAsia,
    fontAscii: card.base.font.ascii,
    sizeHalfPoints: halfPoints(card.base.font.size),
    bold,
  }
}

function spacingOf(paragraph: { lineSpacingMultiple?: number; lineSpacingExactPt?: number }): Pick<ParaFormat, 'lineSpacing' | 'lineRule' | 'lineRawTwips'> {
  if (paragraph.lineSpacingExactPt && paragraph.lineSpacingExactPt > 0) {
    return { lineRule: 'exact', lineRawTwips: pt(paragraph.lineSpacingExactPt) }
  }
  if (paragraph.lineSpacingMultiple && paragraph.lineSpacingMultiple > 0) {
    return { lineSpacing: paragraph.lineSpacingMultiple }
  }
  return {}
}

function bodyFormat(card: ThesisFormatCardInput): ParaFormat {
  const paragraph = card.base.paragraph
  const format: ParaFormat = { align: paragraph.align, ...spacingOf(paragraph) }
  const chars = paragraph.firstLineIndentChars ?? 0
  if (chars > 0) format.indentFirstLine = Math.round(chars * card.base.font.size * PT_TO_TWIPS)
  if (paragraph.spaceBeforePt) format.spaceBefore = pt(paragraph.spaceBeforePt)
  if (paragraph.spaceAfterPt) format.spaceAfter = pt(paragraph.spaceAfterPt)
  return format
}

function generated(runs: Run[], format: ParaFormat, styleId?: string): SaveBlock {
  const block = { type: 'paragraph' as const, runs, format }
  return { kind: 'generated', block: styleId ? { ...block, styleId } : block }
}

function original(block: Block): SaveBlock | null {
  if (block.docxIndex === null) return null
  return { kind: 'original', docxIndex: block.docxIndex }
}

function textOf(block: Block): string {
  return (block.runs?.map((run) => run.text).join('') ?? block.previewText ?? '').trim()
}

/** 去空白与常见标点：模版标题里的「一、」「（一）」和空格不参与匹配。 */
function norm(value: string): string {
  return value.replace(/[\s\u3000]+/g, '').replace(/[、，,．.：:；;（）()【】\[\]“”"'《》<>—－-]/g, '').toLowerCase()
}

/** 模版标题自带的编号前缀（手工编号时替换文字要保留它）。 */
function leadingNumberPrefix(text: string): string {
  const match = /^\s*([（(]?[一二三四五六七八九十百0-9]+[）)]?[、.．]?\s*)/.exec(text)
  if (!match) return ''
  const prefix = match[1]!
  // 只有「数字/中文数字 + 分隔符」才算编号；纯数字标题（如「2024」）不剥
  return /[、.．）)]/.test(prefix) ? prefix : ''
}

/**
 * 去掉开头的编号（「一、」「（二）」「1.」「第一章」）。
 *
 * 入参是 `norm()` 之后的文本——标点已被剥离，所以这里只认数字/中文数字本体。
 * 两边都剥，用于「同一条标题在模版里带编号、在槽位里不带（或反过来）」的对齐。
 */
function stripLeadingNumber(text: string): string {
  return text
    .replace(/^第[一二三四五六七八九十百零0-9]+[章节部分篇]/, '')
    .replace(/^[（(]?[一二三四五六七八九十百零0-9]{1,4}[）)]?/, '')
    .replace(/^[0-9]+(?:\.[0-9]+)*/, '')
    .trim()
}

/** 标题段落是否用 Word 自动编号（numPr）：有则编号由 Word 生成，替换文字不该补前缀。 */
function hasAutoNumbering(block: Block): boolean {
  return /<w:numPr[\s>]/.test(block.originalXml ?? '')
}

/** 模版里是否已有目录域（TOC field）。 */
function hasTocField(blocks: Block[]): boolean {
  return blocks.some((block) => /TOC\s*\\/.test(block.originalXml ?? ''))
}

/**
 * 目录项判定：模版目录若也套 Heading 样式，会被误当成章节锚点，把正文整段插进目录里。
 * 这类行不参与槽位匹配（块本身照旧原样透传）。
 */
function isTocEntry(block: Block): boolean {
  const xml = block.originalXml ?? ''
  if (/<w:instrText[^>]*>\s*TOC\b/i.test(xml) || /<w:fldSimple[^>]*\bTOC\b/i.test(xml)) return true
  // Word 生成的目录项是指向书签的超链接
  if (/<w:hyperlink[^>]*\bw:anchor=/i.test(xml)) return true
  if (/<w:pStyle w:val="[^"]*[Tt][Oo][Cc][^"]*"/.test(xml)) return true
  // 手工目录：「一、项目任务……1」这类引导点 + 页码收尾
  return /[.．·…]{2,}\s*\d{1,4}\s*$/.test(textOf(block))
}

function contentBlocksOfChapter(spec: ThesisBindingSpecInput, start: number, end: number): number[] {
  const out: number[] = []
  for (let index = start + 1; index < end; index += 1) out.push(index)
  return out
}

/** 该块是否要原样搬进模版包（表格/图片/其它受保护块）。 */
function isPassthroughBlock(block: Block): boolean {
  return block.type === 'table' || block.type === 'image' || block.type === 'passthrough'
}

function maxRelId(relsXml: string): number {
  let max = 0
  for (const match of relsXml.matchAll(/\bId="rId(\d+)"/g)) {
    const value = Number(match[1])
    if (Number.isFinite(value) && value > max) max = value
  }
  return max
}

/** 稳定的内容指纹（无 crypto 依赖：渲染端也会打包本模块）。 */
function fnv1a(bytes: Uint8Array): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < bytes.length; index += 1) {
    hash ^= bytes[index]!
    hash = Math.imul(hash, 0x01000193) >>> 0
  }
  return hash.toString(16).padStart(8, '0')
}

const MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  svg: 'image/svg+xml',
}

interface SourceMediaImport {
  bytes: Uint8Array
  /** 把源 XML 里的 rId 换成模版包内的新 rId。 */
  rewrite(xml: string): string
}

/**
 * 把源内容引用到的 media / 外部链接搬进模版包，并给它们分配不与模版冲突的 rId。
 *
 * 不这么做的话，注入的表格/图片会指向源包里根本不存在的 rId —— Word 打开就是
 * 一个红叉。这里只搬「真的被注入块引用到」的那几个文件，不整包复制。
 */
async function importSourceMedia(
  templateBytes: Uint8Array,
  sourceBytes: Uint8Array,
  injectedXml: string[],
): Promise<SourceMediaImport> {
  const identity: SourceMediaImport = { bytes: templateBytes, rewrite: (xml) => xml }
  const used = new Set<string>()
  for (const xml of injectedXml) {
    for (const match of xml.matchAll(/r:(?:embed|id|link)="([^"]+)"/g)) used.add(match[1]!)
  }
  if (used.size === 0) return identity

  const templateZip = await loadDocxZip(templateBytes)
  const sourceZip = await loadDocxZip(sourceBytes)
  const relsPath = 'word/_rels/document.xml.rels'
  const sourceRels = (await sourceZip.file(relsPath)?.async('string')) ?? ''
  const relById = new Map<string, { target: string; type: string; external: boolean }>()
  for (const rel of sourceRels.match(/<Relationship\b[^>]*>/g) ?? []) {
    const id = /\bId="([^"]+)"/.exec(rel)?.[1]
    const target = /\bTarget="([^"]*)"/.exec(rel)?.[1]
    const type = /\bType="([^"]*)"/.exec(rel)?.[1]
    if (!id || !target || !type) continue
    relById.set(id, { target, type, external: /TargetMode="External"/i.test(rel) })
  }

  let templateRels = (await templateZip.file(relsPath)?.async('string'))
    ?? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let nextRel = maxRelId(templateRels) + 1
  const idMap = new Map<string, string>()
  const appended: string[] = []
  const newExtensions = new Set<string>()

  for (const relId of used) {
    const rel = relById.get(relId)
    if (!rel) continue
    if (rel.external) {
      const newId = `rId${nextRel++}`
      appended.push(
        `<Relationship Id="${newId}" Type="${rel.type}" Target="${rel.target.replace(/"/g, '&quot;')}" TargetMode="External"/>`,
      )
      idMap.set(relId, newId)
      continue
    }
    const targetPath = rel.target.startsWith('/')
      ? rel.target.slice(1)
      : `word/${rel.target.replace(/^\.\//, '')}`
    const file = sourceZip.file(targetPath)
    if (!file) continue
    const bytes = await file.async('uint8array')
    const extension = (targetPath.split('.').pop() ?? 'png').toLowerCase()
    const mediaName = `tpl-src-${fnv1a(bytes)}.${extension}`
    const mediaPath = `word/media/${mediaName}`
    if (!templateZip.file(mediaPath)) {
      templateZip.file(mediaPath, bytes)
      newExtensions.add(extension)
    }
    const newId = `rId${nextRel++}`
    appended.push(`<Relationship Id="${newId}" Type="${rel.type}" Target="media/${mediaName}"/>`)
    idMap.set(relId, newId)
  }

  if (idMap.size === 0) return identity
  templateZip.file(relsPath, templateRels.replace('</Relationships>', `${appended.join('')}</Relationships>`))

  // 新扩展名要登记 Default，否则 Word 认为包损坏
  const contentTypes = await templateZip.file('[Content_Types].xml')?.async('string')
  if (contentTypes) {
    const missing = [...newExtensions].filter(
      (extension) => MIME_BY_EXT[extension] && !new RegExp(`Extension="${extension}"`, 'i').test(contentTypes),
    )
    if (missing.length > 0) {
      const added = missing.map((extension) => `<Default Extension="${extension}" ContentType="${MIME_BY_EXT[extension]}"/>`).join('')
      templateZip.file('[Content_Types].xml', contentTypes.replace('</Types>', `${added}</Types>`))
    }
  }

  const bytes = new Uint8Array(await templateZip.generateAsync({ type: 'arraybuffer' }))
  return {
    bytes,
    rewrite: (xml) => xml.replace(/r:(embed|id|link)="([^"]+)"/g, (whole, attr: string, relId: string) => {
      const mapped = idMap.get(relId)
      return mapped ? `r:${attr}="${mapped}"` : whole
    }),
  }
}

/**
 * 正文区起点：目录标题（连同其后的目录项）之后。
 *
 * 模版前部是封面、核心信息区、指导老师评分表、项目经理评价表这类**前置表单**，
 * 它们的标题长得和章节标题一样（短、带编号），如果参与锚点匹配就会把正文插进去
 * ——用户看到的「表格跑到正文第一章下面」正是这么来的。目录之后才是正文骨架。
 */
function bodyRegionStart(blocks: Block[]): number {
  const tocTitle = blocks.findIndex((block) => /^(目录|目次|contents)$/.test(norm(textOf(block))))
  if (tocTitle < 0) return 0
  // 「同名块在后面还出现一次」= 当前这个多半是目录行（正文里那个才是真标题）
  const lastIndexOfBare = new Map<string, number>()
  for (let index = tocTitle + 1; index < blocks.length; index += 1) {
    const bare = stripLeadingNumber(norm(textOf(blocks[index]!)))
    if (!bare) continue
    lastIndexOfBare.set(bare, index)
  }
  let index = tocTitle + 1
  while (index < blocks.length) {
    const block = blocks[index]!
    // 带引导点+页码 / TOC 域 / 超链接：确定的目录行
    if (isTocEntry(block)) {
      index += 1
      continue
    }
    // 真标题才是正文区起点
    if (block.type === 'heading' || block.type === 'table') break
    // 目录条目常常是域/文本框（passthrough），既没有引导点也没有页码
    if (block.type === 'passthrough') {
      index += 1
      continue
    }
    const bare = stripLeadingNumber(norm(textOf(block)))
    if (!bare) {
      index += 1
      continue
    }
    // 后面还有同名块 → 这是目录里的回显行，不是真标题
    if ((lastIndexOfBare.get(bare) ?? index) > index) {
      index += 1
      continue
    }
    break
  }
  return index
}

/** 模版标题块 ↔ 装订目标：按标题文本匹配，保持模版顺序。 */
function resolveAnchors(blocks: Block[], card: ThesisTemplateBinderInput['card'], spec: ThesisBindingSpecInput): Anchor[] {
  const slots = card.profile?.slots ?? []
  const pendingSlots: Target[] = slots.map((slot) => ({ kind: 'slot', slotId: slot.id, title: slot.title }))
  const specials: Target[] = []
  if (spec.abstractZh.trim() || spec.keywordsZh.length > 0) specials.push({ kind: 'abstract_zh', title: '摘要' })
  if (spec.abstractEn.trim() || spec.keywordsEn.length > 0) specials.push({ kind: 'abstract_en', title: 'Abstract' })
  if (spec.references.length > 0) specials.push({ kind: 'references', title: '参考文献' })
  if (spec.acknowledgementParagraphs.length > 0) specials.push({ kind: 'acknowledgement', title: '致谢' })
  if (spec.appendixBlocks) specials.push({ kind: 'appendix', title: '附录' })

  const bodyStart = bodyRegionStart(blocks)

  // 候选标题：真标题，或短文本块。
  // 中文论文模版大量用「普通段落 + 手工编号」模拟标题（没有 Heading 样式），
  // 只认 heading 会让锚点一个都匹配不上，整篇回落成「照内容重建」。
  // 注意：这里从文档开头收集——摘要/参考文献可能在目录之前（摘要必然在）。
  // 「不能把前置表单当章节」这条约束由 matchSlot 里的 bodyStart 保证。
  const candidates: Array<{ index: number; text: string; bare: string; isHeading: boolean }> = []
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]!
    if (isPassthroughBlock(block)) continue
    // 目录项不是正文锚点（否则正文会被插进目录区）
    if (isTocEntry(block)) continue
    const text = norm(textOf(block))
    if (!text) continue
    const isHeading = block.type === 'heading'
    // 正文长段落不可能是章节标题，只对短文本放开
    if (!isHeading && text.length > 40) continue
    candidates.push({ index, text, bare: stripLeadingNumber(text), isHeading })
  }

  const anchors: Anchor[] = []
  const takenSlots = new Set<number>()
  const takenSpecials = new Set<string>()
  const anchored = new Set<number>()

  const matchSlot = (candidate: (typeof candidates)[number], strict: boolean): number => (
    pendingSlots.findIndex((target, slotIndex) => {
      if (takenSlots.has(slotIndex)) return false
      // 正文骨架只在目录之后：封面/核心信息区/指导老师评分表/项目经理评价表
      // 这些前置表单的标题长得和章节标题一样，参与匹配就会把正文插进去。
      if (candidate.index < bodyStart) return false
      const key = norm(target.title)
      if (key.length < 2) return false
      const bareKey = stripLeadingNumber(key)
      // 去掉编号后同名：模版写「一、项目任务」、槽位写「项目任务」也能对上
      if (candidate.bare.length >= 2 && candidate.bare === bareKey) return true
      if (candidate.text === key) return true
      // 包含匹配只给真标题：正文段落里出现同样的词不算标题
      if (!strict && candidate.isHeading) {
        return candidate.text.includes(key) || key.includes(candidate.text)
      }
      return false
    })
  )

  // 第一遍严格匹配（去编号后同名）：手工目录行「一、项目任务……1」抢不走真标题，
  // 因为它的 bare 是「项目任务1」，与槽位的「项目任务」不等。
  for (const candidate of candidates) {
    const slotIndex = matchSlot(candidate, true)
    if (slotIndex < 0) continue
    takenSlots.add(slotIndex)
    anchored.add(candidate.index)
    anchors.push({ target: pendingSlots[slotIndex]!, blockIndex: candidate.index })
  }
  // 第二遍宽松匹配（仅真标题）：兜住「模版标题比槽位标题多几个字」的情况
  for (const candidate of candidates) {
    if (anchored.has(candidate.index)) continue
    const slotIndex = matchSlot(candidate, false)
    if (slotIndex < 0) continue
    takenSlots.add(slotIndex)
    anchored.add(candidate.index)
    anchors.push({ target: pendingSlots[slotIndex]!, blockIndex: candidate.index })
  }
  // 固定部件：摘要/参考文献/致谢/附录
  for (const candidate of candidates) {
    if (anchored.has(candidate.index)) continue
    const specialIndex = specials.findIndex((target) => {
      if (takenSpecials.has(target.kind)) return false
      const keywords = SPECIAL_TARGETS.find((entry) => entry.kind === target.kind)?.keywords ?? [target.title]
      return keywords.some((keyword) => {
        const key = norm(keyword)
        return candidate.text === key
          || candidate.bare === stripLeadingNumber(key)
          || candidate.text.startsWith(key)
      })
    })
    if (specialIndex < 0) continue
    const target = specials[specialIndex]!
    takenSpecials.add(target.kind)
    anchored.add(candidate.index)
    anchors.push({ target, blockIndex: candidate.index })
  }
  // 兜底：模版一个槽位都没解析出来时（普通段落标题 + 卡片是旧版本解析的，
  // profile.slots 是空数组且不会自动重算），按正文区里「短标题」的顺序对应论文章节。
  // 至少保证正文落在模版的章节位置，而不是整体回落 v2 被搬走。
  if (!anchors.some((anchor) => anchor.target.kind === 'slot') && spec.chapters.length > 0) {
    // 找不到「目录」标题时正文区起点是 0，封面文字也会进候选池——那种情况只认带编号的，
    // 否则「XX大学」这类封面行会被当第一章。找到目录之后才放开无编号标题。
    const requireNumbering = bodyStart === 0
    // 目录行即使没被 isTocEntry 认出来（缺引导点/页码），也会在正文里再出现一次。
    // 同名取最后一次出现 = 真标题；取第一次就会把正文塞进目录区。
    const lastIndexOfBare = new Map<string, number>()
    for (const candidate of candidates) {
      lastIndexOfBare.set(candidate.bare, Math.max(lastIndexOfBare.get(candidate.bare) ?? -1, candidate.index))
    }
    const sequential = candidates.filter((candidate) => (
      candidate.index >= bodyStart
      && !anchored.has(candidate.index)
      && lastIndexOfBare.get(candidate.bare) === candidate.index
      && candidate.bare.length >= 2
      && candidate.bare.length <= 30
      // 去编号后还得是中文短语：挡掉封面上的「2024年6月」「2021001」这类日期/学号
      && /[\u4e00-\u9fa5]/.test(candidate.bare)
      && !/^[0-9]{4}/.test(candidate.text)
      // 模版占位说明「（此处填写项目任务）」不是章节标题
      && !/(此处|填写|略$|示例)/.test(candidate.text)
      // 有编号前缀 → 一定是标题；没编号 → 只在目录之后的正文区认，且要够短
      // （中文模版大量用「普通段落 + 手工编号」甚至完全无编号排标题）
      && (
        candidate.bare !== candidate.text
        || (!requireNumbering && (candidate.isHeading || candidate.bare.length <= 20))
      )
    ))
    sequential.slice(0, spec.chapters.length).forEach((candidate, position) => {
      const chapter = spec.chapters[position]!
      anchors.push({
        target: { kind: 'slot', slotId: chapter.slotId ?? `__seq-${position}`, title: chapter.headingText },
        blockIndex: candidate.index,
      })
      anchored.add(candidate.index)
    })
  }

  return anchors.sort((left, right) => left.blockIndex - right.blockIndex)
}

export async function buildThesisDocxOnTemplate(input: ThesisTemplateBinderInput): Promise<ThesisTemplateBinderResult> {
  const { card, spec } = input
  /** 模版优先模式：目录/标题/示例文字全部按模版保留，内容只往小节里填。 */
  const templateFirst = (card.profile as { mode?: string } | undefined)?.mode === 'template-first'
  const template = await parseDocx(input.templateBytes, { patternHeadings: true })
  const source = await parseDocx(input.sourceBytes, { patternHeadings: true })
  const templateBlocks = template.blocks.filter((block) => !block.hidden)
  const sourceBlocks = source.blocks.filter((block) => !block.hidden)

  const bodyStart = bodyRegionStart(templateBlocks)
  const anchors = resolveAnchors(templateBlocks, card, spec)
  // 一个锚点都没有时**不抛错、也不让上层回落「照内容重建」**：
  // 重建会把模版里的表格/评价表按「前置/附录」搬走，比「正文堆到文末」糟得多。
  // 这里的降级是「模版整体原样透传 + 正文追加到文末未归类区」，位置一定不会错。
  const appendedOnly = !anchors.some((anchor) => anchor.target.kind === 'slot')
  const diagnostics: ThesisTemplateBinderDiagnostics = {
    slots: card.profile?.slots?.length ?? 0,
    blocks: templateBlocks.length,
    anchors: anchors.length,
    bodyStart,
    chapters: spec.chapters.length,
    appendedOnly,
  }

  // ---- 章节 → 槽位 ----
  // 优先按 slotId 精确配对；对不上时（老卡片没有 slotId，或顺序兜底生成的 __seq-*）
  // 按锚点顺序依次分配。不能只走 slotId 一条路：网关下发的 slotId 与模版锚点
  // 一旦不同源，整篇章节都会「无槽位」→ 全部变成文末未归类内容。
  const chaptersBySlot = new Map<string, ThesisBindingSpecInput['chapters']>()
  const slotAnchors = anchors.filter((anchor) => anchor.target.kind === 'slot')
  const anchorSlotIds = new Set(
    slotAnchors.flatMap((anchor) => (anchor.target.kind === 'slot' ? [anchor.target.slotId] : [])),
  )
  // 只有当「网关下发的 slotId 与模版锚点一个都对不上」时才允许按序硬配：
  // 这说明两边根本不是同一套槽位（老卡片 / 顺序兜底的 __seq-*）。
  // 若已有一一对应的，说明同源——对不上的章节是真缺锚点，必须走文末未归类兜底，
  // 硬塞进别的槽位会把两章内容混在一起，比标红更难发现。
  const slotIdAligned = spec.chapters.some((chapter) => chapter.slotId && anchorSlotIds.has(chapter.slotId))
  const usedSlotIds = new Set<string>()
  let orderPointer = 0
  for (const chapter of spec.chapters) {
    let slotId: string | undefined
    if (chapter.slotId && anchorSlotIds.has(chapter.slotId)) slotId = chapter.slotId
    if (!slotId && !slotIdAligned) {
      while (orderPointer < slotAnchors.length) {
        const candidate = slotAnchors[orderPointer]!
        orderPointer += 1
        if (candidate.target.kind !== 'slot') continue
        if (usedSlotIds.has(candidate.target.slotId)) continue
        slotId = candidate.target.slotId
        break
      }
    }
    if (!slotId) continue
    usedSlotIds.add(slotId)
    const list = chaptersBySlot.get(slotId) ?? []
    list.push(chapter)
    chaptersBySlot.set(slotId, list)
  }

  // ---- 先收集要注入的受保护块 XML，把 media/rels 搬进模版包 ----
  const injectedXml: string[] = []
  const collect = (indexes: number[]): void => {
    for (const index of indexes) {
      const block = sourceBlocks[index]
      if (!block || !isPassthroughBlock(block) || !block.originalXml) continue
      if (block.originalXml.includes('<w:sectPr')) continue
      injectedXml.push(block.originalXml)
    }
  }
  for (const chapter of spec.chapters) {
    collect(contentBlocksOfChapter(spec, chapter.blockStart, chapter.blockEnd))
  }
  if (spec.appendixBlocks) {
    const indexes: number[] = []
    for (let index = spec.appendixBlocks.start; index < spec.appendixBlocks.end; index += 1) indexes.push(index)
    collect(indexes)
  }
  for (const range of spec.unmatchedBlocks ?? []) {
    const indexes: number[] = []
    for (let index = range.start; index < range.end; index += 1) indexes.push(index)
    collect(indexes)
  }
  const media = await importSourceMedia(input.templateBytes, input.sourceBytes, injectedXml)
  // 重新解析（搬运了源文档媒体后）也要开手打编号识别：否则「六、实验分析与总结」
  // 这类没用标题样式的章节会变成普通段落，既不是归类目标、也拿不到「本节待补充」提示。
  const base = media.bytes === input.templateBytes ? template : await parseDocx(media.bytes, { patternHeadings: true })
  const baseBlocks = base.blocks.filter((block) => !block.hidden)

  // ---- 逐块装配：模版块原样透传，锚点处替换 ----
  const finalBlocks: SaveBlock[] = []
  const anchorByBlock = new Map<number, Anchor>()
  for (const anchor of anchors) anchorByBlock.set(anchor.blockIndex, anchor)

  const numberingScheme: ThesisNumberingScheme = card.numbering?.scheme
    ?? (card.headings.some((heading) => heading.numbering === 'numeric-tree') ? 'decimal-tree' : 'none')
  const counters: [number, number, number, number] = [0, 0, 0, 0]
  const headingStyleIds = ['Heading1', 'Heading2', 'Heading3', 'Heading4']
  const headingByLevel = new Map(card.headings.map((heading) => [heading.level, heading]))
  const tocEntries: Array<{ text: string; level: number }> = []

  // ---- 目录 ----
  // 三种情况：
  // 1) 模版有 TOC 域：原样保留（Word 打开自动刷新）；
  // 2) 模版有「目录」标题但条目是手工敲的（没有域）：条目会过期——把标题后的条目区
  //    丢掉，换成用真实章节生成的 TOC 域；
  // 3) 模版没有目录：按格式卡开关补一个。
  // 目录在正文之前，条目要等章节装配完才知道，所以先在 finalBlocks 里占位、收尾再填。
  let tocSlotIndex: number | null = null
  const templateHasToc = hasTocField(templateBlocks)
  const tocTitleIndex = templateBlocks.findIndex((block) => /^(目录|目次|contents)$/.test(norm(textOf(block))))
  // 目录标题到正文区起点之间就是「目录条目区」——不管模版是手工敲的还是 TOC 域的缓存结果，
  // 里面的条目都是模版自己的章节，跟本次论文无关，必须换掉（否则目录列的是别人的章节）。
  const tocRegion = tocTitleIndex >= 0 && bodyStart > tocTitleIndex + 1
    ? { start: tocTitleIndex + 1, end: bodyStart }
    : null
  const replaceTocEntries = !templateFirst && tocRegion !== null && card.toc.enabled

  // ---- 模版优先模式（template-first）----
  // 目录、章节标题、模版自带正文都按模版原样保留；论文内容按「小节语义」归类插到
  // 模版标题下面。归类规则是确定性的（不依赖模型）：把论文的每个小节标题和模版的
  // 每个章节/小节标题都拆成「概念词」，同一概念组里命中越多分越高，取最高分的槽位；
  // 一个词都对不上的，接到上一段已落位内容的末尾（用户确认的「归到最接近小节末尾」）。
  interface FlatSection { heading?: string; blocks: number[]; bullets?: string[] }
  const contentAt = new Map<number, FlatSection[]>()
  /** 拆列后的表格 XML（按源块下标索引）；渲染时用它替掉原表。 */
  const tableXmlOverride = new Map<number, string>()

  /**
   * 模版标题 → 「哪些词算讲这件事」。这是把「学校规定的 7 个点」翻译成论文常用词：
   * 标题用词不同（论文写「验证与里程碑」、模版写「项目计划」）也能对上。
   * 顺序即优先级：先命中的规则赢（如「概要设计」要先于「系统设计」）。
   */
  const KEYWORD_RULES: Array<{ when: RegExp; keys: string[] }> = [
    { when: /背景/, keys: ['背景', '痛点', '政策', '现状', '问题', '场景'] },
    { when: /目标/, keys: ['目标', '思路', '创新', '意义', '目的'] },
    { when: /分工/, keys: ['分工', '成员', '组织', '人员'] },
    { when: /计划/, keys: ['计划', '里程碑', '进度', '排期'] },
    { when: /提交|交付/, keys: ['提交', '交付', '成果', '验收'] },
    { when: /概要/, keys: ['模块', '划分', '概要', '组成'] },
    { when: /详细/, keys: ['时序', '流程', '详细', '实现', '步骤'] },
    { when: /需求/, keys: ['需求', '功能', '范围', '角色', '用例'] },
    { when: /系统设计|系统/, keys: ['架构', '技术', '选型', '数据流', '决策', '接口'] },
    { when: /测试/, keys: ['测试', '验证', '用例', '缺陷', '结果'] },
    { when: /总结|展望/, keys: ['总结', '展望', '结论'] },
  ]
  const keysOf = (title: string): string[] => {
    const flat = norm(title)
    for (const rule of KEYWORD_RULES) {
      if (rule.when.test(flat)) return rule.keys
    }
    // 没有专属规则：用标题自身的词
    return flat.replace(/^[一二三四五六七八九十]+[、.\s]*/, '').replace(/[0-9.\s]/g, '') ? [flat] : []
  }
  /** 小节标题与模版标题的契合度：命中词越长越具体，分越高。 */
  const semanticScore = (sectionTitle: string, targetTitle: string): number => {
    const flat = norm(sectionTitle)
    if (!flat) return 0
    let score = 0
    for (const key of keysOf(targetTitle)) if (flat.includes(norm(key))) score += norm(key).length
    return score
  }

  /**
   * 模版正文里的归类目标：章标题、小节标题，以及「1.1.1」这类三级小标题。
   * 三级标题也要能接内容、能出「本节待补充」提示——学校模版经常一路细到 1.1.1。
   */
  const targets: Array<{ blockIndex: number; title: string }> = []
  for (let index = bodyStart; index < baseBlocks.length; index += 1) {
    const block = baseBlocks[index]!
    // 模版的一级标题有两种写法：heading 样式，或「列表项 + 自动编号」（四、/五、这种）。
    const autoHeading = block.type === 'listItem' && /<w:numPr/.test(block.originalXml ?? '')
    if (block.type !== 'heading' && !autoHeading) continue
    const level = block.level ?? 1
    if (level > 3) continue
    const title = norm(textOf(block))
    if (!title) continue
    // 目录/摘要/参考文献/致谢/附录这类部件不接收正文；
    // 封面表单行（题目/姓名/学号/学院…）也不是章节——不排除的话，
    // 没有目录标题的模版会把封面当作第一个归类目标，正文落进封面。
    if (/^(目录|目次|contents|摘要|abstract|参考文献|致谢|附录|封面)/.test(title)) continue
    if (/^(课程名称|课程设计|题目|姓名|学号|学院|专业|指导教师|完成时间|承诺|声明|签名|签字|日期|班级|年级)/.test(title)) continue
    targets.push({ blockIndex: index, title })
  }

  /** 论文小节：[标题, 下一同级或更高级标题) 的区间；二级标题只当容器时也保留自身内容。 */
  const sourceSections: Array<{ title: string; level: number; blocks: number[]; from: number }> = []
  for (let index = 0; index < sourceBlocks.length; index += 1) {
    const block = sourceBlocks[index]!
    if (block.type !== 'heading') continue
    const level = block.level ?? 9
    if (level < 2) continue
    // 只收「到下一个标题之前」的自有内容：子标题的内容归子节自己，
    // 否则父节会把整章正文吞进来，子节再吐一遍（重复 + 顺序错乱）。
    let stop = sourceBlocks.length
    for (let next = index + 1; next < sourceBlocks.length; next += 1) {
      const candidate = sourceBlocks[next]!
      if (candidate.type === 'heading' || (candidate.type === 'listItem' && candidate.level)) {
        stop = next
        break
      }
    }
    const blocks: number[] = []
    for (let inner = index + 1; inner < stop; inner += 1) blocks.push(inner)
    sourceSections.push({ title: textOf(block), level, blocks, from: index })
  }

  // 章级正文（一级标题下面、第一个子标题之前的导语）也要参与归类：
  // 只靠 sourceSections（level>=2）会漏掉它，「第一章 前言」正文整段消失。
  // 作为「伪小节」进同一套语义规则，标题相同/语义相近时自然落到对应章节。
  for (const chapter of spec.chapters) {
    const leadIn: number[] = []
    for (const index of contentBlocksOfChapter(spec, chapter.blockStart, chapter.blockEnd)) {
      const block = sourceBlocks[index]
      if (!block) continue
      const autoHeading = block.type === 'listItem' && /<w:numPr/.test(block.originalXml ?? '')
      if (block.type === 'heading' || autoHeading) break
      leadIn.push(index)
    }
    if (leadIn.length === 0) continue
    sourceSections.push({ title: chapter.headingText, level: 1, blocks: leadIn, from: chapter.blockStart })
  }

  /** AI 补写内容：按小节标题（去编号）索引。 */
  /** 小节标题比对键：去掉「七、」「1.2」这类编号与空白，模型省编号时也能对上。 */
  const sectionKey = (value: string): string =>
    norm(value).replace(/^[一二三四五六七八九十百]+[、.．]?/, '').replace(/^\d+(\.\d+)*[、.．]?/, '')
  let currentChapterTitle = ""
  const aiTextFor = (title: string): string | undefined => {
    const withContext = currentChapterTitle ? generatedByContext.get(norm(`${currentChapterTitle}｜${title}`)) : undefined
    return withContext ?? generatedByTitle.get(sectionKey(title))
  }
  const generatedByTitle = new Map<string, string>()
  /**
   * 「章｜小节」精确索引：同名小节（8 个实验都有「实验步骤」）必须各写各的。
   * 网关下发时把上下文拼进标题（`实验二｜实验步骤`），这里按整串匹配；
   * 退路仍是只按小节名匹配（老数据/模型省了上下文时仍能补上）。
   */
  const generatedByContext = new Map<string, string>()
  for (const entry of spec.generatedSections ?? []) {
    const text = entry.text?.trim()
    if (!text) continue;
    const raw = norm(entry.title ?? "");
    if (raw.includes("｜")) {
      generatedByContext.set(raw, text);
      continue;
    }
    const key = sectionKey(entry.title)
    if (key) generatedByTitle.set(key, text)
  }
  /** 模版里的归类目标集合（只有这些块才可能接收论文内容）。 */
  const targetIndexSet = new Set(targets.map((target) => target.blockIndex))
  /** 空小节提示开关：card.profile.placeholders === false 时关闭。 */
  const placeholdersEnabled = (card.profile as { placeholders?: boolean } | undefined)?.placeholders !== false
  /** 模版这个标题下面本来就有正文（示例文字/说明）时，不算空节。 */
  const templateHasOwnBody = (index: number): boolean => {
    for (let next = index + 1; next < baseBlocks.length; next += 1) {
      const block = baseBlocks[next]!
      const autoHeading = block.type === 'listItem' && /<w:numPr/.test(block.originalXml ?? '')
      if (block.type === 'heading' || autoHeading) return false
      if (textOf(block).trim()) return true
    }
    return false
  }

  // 骨架模式的可用性闸门：模版正文至少要认出 2 个归类锚点、且覆盖过半槽位。
  // 「只有目录、正文没有章节标题」的模版（学校常见写法）在这里会被拦下，
  // 回落到内容重建——否则整篇内容会挤进唯一认出的那一节里。
  const templateFirstUsable = targets.length >= 2
    && (!card.profile?.slots?.length || targets.length >= Math.ceil(card.profile.slots.length / 2))
  if (!templateFirstUsable) {
    diagnostics.templateFirstDowngraded = true
  }
  /**
   * 挂起的内容：模版章节标题下面常常自带固定块（实验日期、指导老师、实验评分表、
   * ★预习提示…）。这些必须紧跟标题，不能把论文正文插在它们前面。
   * 所以内容先挂起，等这一节的固定块都落位、遇到下一个标题时再插入。
   */
  let pendingSectionContent: FlatSection[] | null = null
  const flushPendingSectionContent = (): void => {
    if (!pendingSectionContent) return
    pushFlatContent(pendingSectionContent)
    pendingSectionContent = null
  }

  // 模版一个可归类标题都没有时，只做「模版原样输出」，不能崩
  if (templateFirst && templateFirstUsable && targets.length > 0) {
    let lastTarget = -1
    for (const section of sourceSections) {
      const usable = section.blocks.filter((index) => {
        const block = sourceBlocks[index]!
        if (!block.originalXml && (block.runs ?? []).length === 0) return false
        if (isPassthroughBlock(block)) return Boolean(block.originalXml)
        const text = textOf(block)
        return Boolean(text) && !/^(---|\/|graph\s+(TD|LR)|flowchart\b)/i.test(text)
      })
      if (usable.length === 0) continue
      let bestTarget = -1
      let bestScore = 0
      for (let order = 0; order < targets.length; order += 1) {
        const score = semanticScore(section.title, targets[order]!.title)
        if (score > bestScore) {
          bestScore = score
          bestTarget = order
        }
      }
      const targetOrder = bestTarget >= 0 ? bestTarget : (lastTarget >= 0 ? lastTarget : 0)
      lastTarget = targetOrder
      // ---- 拆分发牌：表格里有一列讲的是另一个小节（如「验证结果」→ 测试结果）----
      // 只拆一列、只拆结构简单的表（无合并单元格/嵌套表），把那一列搬成对方小节的条目，
      // 原表同时删掉该列——搬走而不是复制，避免同一批数据在两处出现。
      for (const blockIndex of usable) {
        const tableBlock = sourceBlocks[blockIndex]!
        if (tableBlock.type !== 'table' || !tableBlock.table || !tableBlock.originalXml) continue
        if (/<w:tbl[\s>]/.test(tableBlock.originalXml.replace(/^[\s\S]*?<w:tbl[\s>]/, ''))) continue // 嵌套表不拆
        const rows = tableBlock.table.rows
        if (rows.length < 2 || rows[0].length < 2) continue
        const unsafe = rows.some((row) => row.some((cell) => (cell.colSpan ?? 1) > 1 || cell.vMerge || cell.gridGap))
        if (unsafe) continue
        let bestColumn = -1
        let bestColumnTarget = -1
        let bestColumnScore = 0
        for (let column = 1; column < rows[0].length; column += 1) {
          const header = cellText(rows[0][column])
          if (!header) continue
          for (let order = 0; order < targets.length; order += 1) {
            if (order === targetOrder) continue
            const score = semanticScore(header, targets[order]!.title)
            if (score > bestColumnScore) {
              bestColumnScore = score
              bestColumn = column
              bestColumnTarget = order
            }
          }
        }
        if (bestColumn > 0 && bestColumnTarget >= 0 && bestColumnScore >= 3) {
          const values = rows.slice(1)
            .map((row) => {
              const label = cellText(row[0])
              const value = cellText(row[bestColumn])
              if (!value) return ''
              return label ? `${label}：${value}` : value
            })
            .filter(Boolean)
          if (values.length > 0) {
            tableXmlOverride.set(blockIndex, removeTableColumn(tableBlock.originalXml, bestColumn))
            const splitAt = targets[bestColumnTarget]!.blockIndex
            const splitList = contentAt.get(splitAt) ?? []
            splitList.push({ blocks: [], bullets: values })
            contentAt.set(splitAt, splitList)
          }
        }
      }
      const blockIndex = targets[targetOrder]!.blockIndex
      const list = contentAt.get(blockIndex) ?? []
      list.push({
        ...(section.level >= 3 ? { heading: section.title } : {}),
        blocks: usable,
      })
      contentAt.set(blockIndex, list)
    }
  }

  const pushTocPlaceholder = (): void => {
    if (tocSlotIndex !== null) return
    tocSlotIndex = finalBlocks.length
    finalBlocks.push({ kind: 'xml', xml: '' })
  }

  let tocInserted = false
  const ensureToc = (): void => {
    if (tocInserted || !card.toc.enabled) return
    if (templateHasToc) return
    tocInserted = true
    // 模版自己就有「目录」标题：条目位置已经由 replaceTocEntries 分支占好了
    if (tocTitleIndex >= 0) return
    finalBlocks.push(generated([textRun(card, '目录')], { align: 'center', spaceBefore: pt(12), spaceAfter: pt(12) }))
    pushTocPlaceholder()
  }

  /** 源内容块 → 装订块：文本套模版正文样式，表格/图片原样搬（rId 已重映射）。 */
  const pushSourceBlocks = (indexes: number[]): void => {
    for (const index of indexes) {
      const block = sourceBlocks[index]
      if (!block) continue
      if (block.originalXml?.includes('<w:sectPr')) continue
      if (isPassthroughBlock(block)) {
        if (!block.originalXml) continue
        const overridden = tableXmlOverride.get(index) ?? block.originalXml
        finalBlocks.push({ kind: 'xml', xml: media.rewrite(overridden) })
        continue
      }
      if (block.type === 'paragraph' || block.type === 'heading' || block.type === 'listItem') {
        const runs = (block.runs ?? []).map((run) =>
          run.image || run.math || run.ruby || run.noteRef || run.link
            ? run
            : {
              ...run,
              font: card.base.font.eastAsia,
              fontAscii: card.base.font.ascii,
              sizeHalfPoints: halfPoints(card.base.font.size),
            },
        )
        if (runs.length === 0) continue
        finalBlocks.push(generated(runs, bodyFormat(card)))
      }
    }
  }

  /**
   * 模版优先模式的内容推送：源小节标题 → 粗体段落（去掉自带编号），
   * 正文段落套模版正文样式，表格原样搬；分隔线/图代码这类噪声丢弃。
   */
  const pushFlatContent = (sections: FlatSection[]): void => {
    for (const section of sections) {
      // 源文档的小节标题 → 粗体小标题（编号已去掉），方便对照原文找回
      const heading = section.heading
        ? section.heading.replace(/^\s*\d+(\.\d+)*\s*/, '').trim()
        : ''
      if (heading) {
        finalBlocks.push(generated([textRun(card, heading, true)], {
          ...bodyFormat(card),
          indentFirstLine: 0,
          spaceBefore: pt(6),
        }))
      }
      pushSectionBlocks(section.blocks)
      for (const line of section.bullets ?? []) {
        finalBlocks.push(generated([textRun(card, line)], { ...bodyFormat(card), indentFirstLine: 0 }))
      }
    }
  }

  const pushSectionBlocks = (indexes: number[]): void => {
    for (const index of indexes) {
      const block = sourceBlocks[index]
      if (!block) continue
      if (block.originalXml?.includes('<w:sectPr')) continue
      if (isPassthroughBlock(block)) {
        if (!block.originalXml) continue
        const overridden = tableXmlOverride.get(index) ?? block.originalXml
        finalBlocks.push({ kind: 'xml', xml: media.rewrite(overridden) })
        continue
      }
      const text = textOf(block)
      if (!text || /^(---|\/|graph\s+(TD|LR)|flowchart\b)/i.test(text)) continue
      if (block.type === 'heading') continue
      finalBlocks.push(generated([textRun(card, text)], bodyFormat(card)))
    }
  }

  /** 章节标题：首个章节沿用模版标题的版式，其余按模版标题层级生成。 */
  const pushChapterHeading = (chapter: ThesisBindingSpecInput['chapters'][number], templateHeading: Block | null): void => {
    const level = chapter.level
    if (numberingScheme !== 'none') {
      const slot = level - 1
      counters[slot] += 1
      for (let index = slot + 1; index < 4; index += 1) counters[index] = 0
    }
    const numberText = headingNumberPrefix(numberingScheme, counters, level)
    // 模版标题的编号可能是三种来源：段落里的 numPr、样式自带的自动编号（引擎会把它
    // 归成 listItem）、或者纯手工敲在文字里。自动编号不能补前缀，否则 Word 里会出现
    // 「四、四、系统架构图」这种重号。
    const autoNumbered = Boolean(templateHeading && (hasAutoNumbering(templateHeading) || templateHeading.type === 'listItem'))
    const templatePrefix = templateHeading?.originalXml && !autoNumbered
      ? leadingNumberPrefix(textOf(templateHeading))
      : ''
    const prefix = autoNumbered ? '' : (templatePrefix || numberText)
    const emittedText = `${prefix}${chapter.headingText}`
    // 目录条目必须和「读者看到的标题」一致：自动编号的由 Word 生成，目录里用算出来的编号补齐。
    tocEntries.push({ text: autoNumbered ? `${numberText}${chapter.headingText}` : emittedText, level: Math.min(level, 3) })

    if (templateHeading?.originalXml) {
      const patched = patchParagraphTexts(templateHeading.originalXml, emittedText)
      if (patched) {
        finalBlocks.push({ kind: 'xml', xml: patched })
        return
      }
    }
    const heading = headingByLevel.get(level) ?? headingByLevel.get(1) ?? { level: 1, font: card.base.font, align: 'center' as const, numbering: 'none' as const }
    finalBlocks.push({
      kind: 'generated',
      block: {
        type: 'heading',
        level,
        styleId: headingStyleIds[level - 1]!,
        runs: [{
          text: emittedText,
          font: heading.font.eastAsia,
          fontAscii: heading.font.ascii,
          sizeHalfPoints: halfPoints(heading.font.size),
          bold: heading.font.bold,
        }],
        format: { align: heading.align, ...spacingOf(heading.paragraph ?? {}) },
      },
    })
  }

  let skippingPlaceholder = false
  // 源文档开头的块（题目/说明行，区间从 0 起）不是「孤儿」，它们本来就该在文档开头。
  // 放到文末标红是错位的：这里改成第一个章节锚点之前原样落位，且不加警示标记。
  const allUnmatched = spec.unmatchedBlocks ?? []
  const frontRanges = allUnmatched.filter((range) => range.start === 0)
  const tailRanges = allUnmatched.filter((range) => range.start !== 0)
  let frontEmitted = false
  const flushFront = (): void => {
    if (frontEmitted) return
    frontEmitted = true
    for (const range of frontRanges) {
      const indexes: number[] = []
      for (let index = range.start; index < range.end; index += 1) indexes.push(index)
      pushSourceBlocks(indexes)
    }
  }
  /** 已落位的章节：装配结束后据此做守恒检查，漏掉的章节文末补出，绝不静默丢弃。 */
  const emittedChapters = new Set<ThesisBindingSpecInput['chapters'][number]>()
  for (let index = 0; index < baseBlocks.length; index += 1) {
    const block = baseBlocks[index]!
    const anchor = anchorByBlock.get(index)

    // 分节符段落承载模版自己的页面设置：页边距、页眉页脚、页码方案（封面节 → 正文节）。
    // 它常常正好夹在「目录旧条目区」的末尾（目录标题之后、正文第一章之前）——
    // 上面的目录清理和下面的占位示例清理都会把它当成垃圾一并吞掉；封面一旦并进
    // 正文节，页边距/页眉全变，核心信息区就会被挤错位、拆到第二页。任何路径都保它。
    if (!anchor && block.originalXml?.includes('<w:sectPr')) {
      const kept = original(block)
      if (kept) finalBlocks.push(kept)
      continue
    }

    // 模版优先模式：模版内容整篇原样走一遍；轮到归类目标时把论文内容接在标题后面
    if (templateFirst && templateFirstUsable) {
      // 进入下一节之前，把上一节挂起的内容落下去（这样模版固定块仍紧跟标题）
      if ((block.type === 'heading' || hasAutoNumbering(block)) && pendingSectionContent) {
        flushPendingSectionContent()
      }
      const kept = original(block)
      if (kept) finalBlocks.push(kept)
      // 记下当前章标题（一级标题），用于「章｜小节」精确定位补写内容
      if ((block.type === 'heading' || hasAutoNumbering(block)) && (block.level ?? 1) <= 1 && textOf(block)) {
        currentChapterTitle = textOf(block)
      }
      const stashed = contentAt.get(index)
      if (stashed) pendingSectionContent = stashed
      // 论文没内容、但用户开了 AI 补写且这一节有生成结果 → 插入带标记的补写段落
      else if (aiTextFor(textOf(block))) {
        finalBlocks.push(generated(
          [{
            text: 'AI补写，请认真核对',
            font: card.base.font.eastAsia,
            fontAscii: card.base.font.ascii,
            sizeHalfPoints: halfPoints(card.base.font.size),
            italic: true,
            color: '808080',
          }],
          { ...bodyFormat(card), indentFirstLine: 0 },
        ))
        for (const line of aiTextFor(textOf(block))!.split(/\n+/)) {
          const trimmed = line.trim()
          if (trimmed) finalBlocks.push(generated([textRun(card, trimmed)], bodyFormat(card)))
        }
      }
      // 模版有这一节、论文没内容、模版自己也没写正文 → 留一行浅灰提示，
      // 让读者知道「这里要写什么」而不是以为排版漏了。
      else if (
        placeholdersEnabled
        && targetIndexSet.has(index)
        && !/^(参考资料|参考文献|附录|致谢)/.test(norm(textOf(block)))
        && !templateHasOwnBody(index)
      ) {
        finalBlocks.push(generated(
          [{
            text: `本节待补充：${textOf(block)}`,
            font: card.base.font.eastAsia,
            fontAscii: card.base.font.ascii,
            sizeHalfPoints: halfPoints(card.base.font.size),
            italic: true,
            color: '808080',
          }],
          { ...bodyFormat(card), indentFirstLine: 0 },
        ))
      }
      continue
    }

    if (replaceTocEntries && tocRegion && index >= tocRegion.start && index < tocRegion.end) {
      continue
    }

    if (anchor) {
      ensureToc()
      flushFront()
      const target = anchor.target
      if (target.kind === 'slot') {
        const chapters = chaptersBySlot.get(target.slotId) ?? []
        if (chapters.length === 0) {
          // 模版有这一章节位、但本次论文没有对应内容：保留模版标题（用户还要往里写），
          // 并且必须登记目录条目——否则正文多一章、目录少一条，两边对不上。
          // 标题下的示例正文照样跳过：留一个空标题供填写，不留模版的说明文字。
          const level = Math.min((block.type === 'heading' ? block.level ?? 1 : 1), 3) as 1 | 2 | 3
          const headingText = textOf(block)
          if (headingText) {
            const autoNumbered = hasAutoNumbering(block) || block.type === 'listItem'
            let numberText = ''
            if (autoNumbered && numberingScheme !== 'none') {
              const slot = level - 1
              counters[slot] += 1
              for (let index = slot + 1; index < 4; index += 1) counters[index] = 0
              numberText = headingNumberPrefix(numberingScheme, counters, level)
            }
            // 自动编号的标题正文里没有编号文字（Word 渲染时才出现），目录用算出来的补；
            // 手工编号的标题文字自带「七、」，原样登记即可。
            tocEntries.push({ text: autoNumbered ? `${numberText}${headingText}` : headingText, level })
          }
          const kept = original(block)
          if (kept) finalBlocks.push(kept)
          skippingPlaceholder = true
          continue
        }
        chapters.forEach((chapter, position) => {
          emittedChapters.add(chapter)
          pushChapterHeading(chapter, position === 0 ? block : null)
          pushSourceBlocks(contentBlocksOfChapter(spec, chapter.blockStart, chapter.blockEnd))
        })
      } else {
        const kept = original(block)
        if (kept) finalBlocks.push(kept)
        if (target.kind === 'abstract_zh') {
          if (spec.abstractZh.trim()) {
            finalBlocks.push(generated([textRun(card, spec.abstractZh.trim())], bodyFormat(card)))
          }
          if (spec.keywordsZh.length > 0) {
            finalBlocks.push(generated([
              textRun(card, card.abstract.keywordPrefix, card.abstract.keywordPrefixBold),
              textRun(card, spec.keywordsZh.join(card.abstract.keywordSeparator)),
            ], { ...bodyFormat(card), indentFirstLine: 0 }))
          }
        } else if (target.kind === 'abstract_en') {
          if (spec.abstractEn.trim()) {
            finalBlocks.push(generated([textRun(card, spec.abstractEn.trim())], bodyFormat(card)))
          }
          if (spec.keywordsEn.length > 0) {
            finalBlocks.push(generated([
              textRun(card, 'Keywords: ', card.abstract.keywordPrefixBold),
              textRun(card, spec.keywordsEn.join(', ')),
            ], { ...bodyFormat(card), indentFirstLine: 0 }))
          }
        } else if (target.kind === 'references') {
          for (const entry of spec.references) {
            finalBlocks.push(generated([textRun(card, entry)], {
              align: card.references.listParagraph.align,
              ...spacingOf(card.references.listParagraph),
            }))
          }
        } else if (target.kind === 'acknowledgement') {
          for (const paragraph of spec.acknowledgementParagraphs) {
            finalBlocks.push(generated([textRun(card, paragraph)], bodyFormat(card)))
          }
        } else if (target.kind === 'appendix' && spec.appendixBlocks) {
          const indexes: number[] = []
          for (let i = spec.appendixBlocks.start; i < spec.appendixBlocks.end; i += 1) indexes.push(i)
          pushSourceBlocks(indexes)
        }
      }
      skippingPlaceholder = true
      continue
    }

    // 模版目录的标题：保留标题，条目位置换成真实章节生成的 TOC 域
    if (replaceTocEntries && index === tocTitleIndex) {
      const kept = original(block)
      if (kept) finalBlocks.push(kept)
      pushTocPlaceholder()
      continue
    }

    // 标题是否「模版固定部件」（参考文献/致谢/附录/摘要…）：这类标题在占位区里也要留。
    const fixedHeading = block.type === 'heading' && isFixedPartTitle(textOf(block))

    // 任何固定标题都结束上一段占位区：模版自己的「致谢」等不会被误删
    if (block.type === 'heading') {
      // 占位区里的「1.1 背景」「1.4 项目计划」这类示例子标题要丢掉——
      // 章节内容已经由源论文提供，留着就成了正文里凭空多出来的空小节。
      if (skippingPlaceholder && !fixedHeading) continue
      skippingPlaceholder = false
      const kept = original(block)
      if (kept) finalBlocks.push(kept)
      continue
    }

    if (skippingPlaceholder) {
      // 占位区里只保留表格（模版固定表单常常就长在章节里），其余是示例内容，丢弃
      if (block.type === 'table') {
        const kept = original(block)
        if (kept) finalBlocks.push(kept)
      }
      continue
    }

    const kept = original(block)
    if (kept) finalBlocks.push(kept)
  }

  // ---- 目录域回填：见本函数末尾 ----
  // 目录在正文之前，装配到它的时候 tocEntries 还是空的；而「没匹配到槽位的章节」
  // 是在下面的兜底分支里才补出来的，补的时候同样会往 tocEntries 里登记。
  // 所以回填必须等到兜底章节也落位之后再执行，否则目录会永远少最后几章。
  // 两类都在这里兜底：
  // 1) 章节已映射到槽位，但模版里那个槽位标题没被解析出来（锚点缺失）——最隐蔽的一类，
  //    不兜底就会静默丢掉整章；
  // 2) 网关显式标记为未匹配的区间。
  // ---- 模板没有对应槽位的内容：文末单独成节，红字高亮，绝不丢 ----
  // 两类都在这里兜底：
  // 1) 章节已映射到槽位，但模版里那个槽位标题没被解析出来（锚点缺失）——最隐蔽的一类，
  //    不兜底就会静默丢掉整章；
  // 2) 网关显式标记为未匹配的区间。
  // 最后一节挂起的内容在这里落地（后面没有标题来触发了）
  flushPendingSectionContent()
  const orphanChapters = templateFirst && templateFirstUsable ? [] : spec.chapters.filter((chapter) => !emittedChapters.has(chapter))
  const unmatchedRanges = templateFirst && templateFirstUsable ? [] : tailRanges
  // 没有任何锚点时（appendedOnly）front 区也要先落位，否则会被当成文末内容
  if (!(templateFirst && templateFirstUsable)) flushFront()
  if (orphanChapters.length > 0 || unmatchedRanges.length > 0) {
    // appendedOnly = 整篇正文都没落位（模版骨架一个锚点都没认出来）。
    // 这时不能把每章都刷成红色警告——用户会看到一整篇红字，比排版错乱还吓人。
    // 改成一条中性过渡标题，正文照模版正文格式正常排；诊断信息走 warnings 到面板。
    if (appendedOnly) {
      finalBlocks.push(generated(
        [textRun(card, '论文正文（模版未识别章节位置，已按章节顺序追加于文末）')],
        { ...bodyFormat(card), align: 'center', indentFirstLine: 0 },
      ))
    } else {
      finalBlocks.push(generated([{
        ...textRun(card, '未归类内容（模板无对应槽位，请人工归位）'),
        color: 'C00000',
        highlight: 'yellow',
      }], bodyFormat(card)))
    }
    for (const chapter of orphanChapters) {
      if (!appendedOnly) {
        finalBlocks.push(generated([{
          ...textRun(card, `〔未落位：${chapter.headingText}〕`),
          color: 'C00000',
          highlight: 'yellow',
        }], bodyFormat(card)))
      }
      pushChapterHeading(chapter, null)
      pushSourceBlocks(contentBlocksOfChapter(spec, chapter.blockStart, chapter.blockEnd))
    }
    for (const range of unmatchedRanges) {
      if (range.label && !appendedOnly) {
        finalBlocks.push(generated([{
          ...textRun(card, `〔未匹配：${range.label}〕`),
          color: 'C00000',
          highlight: 'yellow',
        }], bodyFormat(card)))
      }
      const indexes: number[] = []
      for (let index = range.start; index < range.end; index += 1) indexes.push(index)
      pushSourceBlocks(indexes)
    }
  }

  // ---- 目录域回填：所有章节（含文末兜底补出来的）都登记完才填 ----
  // 目录在正文之前，装配到它时还不知道后面有哪些章；兜底章节又是最后才补的。
  // 放在这里（saveDocx 之前）才能保证目录条目 = 正文里实际出现的章节。
  if (tocSlotIndex !== null) {
    const maxLevel = Number.parseInt(card.toc.levels.replace(/[^0-9]/g, '').slice(-1), 10)
    const tocXml = tocEntries.length > 0
      ? generateTocFieldXml(
        tocEntries,
        Number.isFinite(maxLevel) && maxLevel >= 1 && maxLevel <= 9 ? maxLevel : undefined,
      ).join('')
      : ''
    if (tocXml) {
      finalBlocks[tocSlotIndex] = { kind: 'xml', xml: tocXml }
    } else {
      finalBlocks.splice(tocSlotIndex, 1)
    }
  }

  // 不传 section/header/footer：模版的页面设置与页眉页脚原样保留
  return { bytes: await saveDocx(base, finalBlocks, {}), diagnostics }
}
