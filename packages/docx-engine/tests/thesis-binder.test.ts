import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { buildBlankDocx } from '../src/blank'
import { parseDocx } from '../src/parse'
import { saveDocx, type SaveBlock } from '../src/patch'
import { buildThesisDocx, describeThesisBlocks, type ThesisFormatCardInput, type ThesisBindingSpecInput } from '../src/thesis-binder'

const font = (eastAsia: string, size: number, bold = false) => ({
  eastAsia, ascii: 'Times New Roman', size, bold,
})

const CARD: ThesisFormatCardInput = {
  page: {
    widthMm: 210, heightMm: 297, orientation: 'portrait',
    marginTopMm: 25, marginBottomMm: 25, marginLeftMm: 20, marginRightMm: 20, gutterMm: 5,
  },
  base: {
    font: font('宋体', 12),
    paragraph: { align: 'justify', firstLineIndentChars: 2, lineSpacingMultiple: 1.25 },
  },
  header: { enabled: true, template: '{title}', font: font('宋体', 10.5), underline: true },
  footer: { enabled: true, font: font('黑体', 10.5, true), frontMatterFormat: 'roman' },
  headings: [
    { level: 1, font: font('黑体', 18, true), align: 'center', numbering: 'numeric-tree', paragraph: { spaceBeforePt: 12, spaceAfterPt: 12 } },
    { level: 2, font: font('黑体', 15, true), align: 'left', numbering: 'numeric-tree' },
    { level: 3, font: font('黑体', 14, true), align: 'left', numbering: 'numeric-tree' },
  ],
  abstract: {
    titleFont: font('黑体', 18, true),
    bodyFont: font('宋体', 12),
    bodyParagraph: { align: 'justify', firstLineIndentChars: 2, lineSpacingMultiple: 1.25 },
    keywordPrefix: '关键词：',
    keywordPrefixBold: true,
    keywordSeparator: '，',
  },
  toc: { enabled: true, levels: '1-3', titleFont: font('黑体', 18, true) },
  references: { listFont: font('宋体', 12), listParagraph: { align: 'justify', lineSpacingMultiple: 1.25 } },
  cover: {
    schoolName: '嘉应学院',
    degreeLine: '本科毕业论文（设计）',
    yearLine: '（2026届）',
    fields: [
      { label: '题    目', key: 'title' },
      { label: '姓    名', key: 'name' },
    ],
    footerLine: '嘉 应 学 院 教 务 部 制',
    schoolFont: font('宋体', 26, true),
    fieldLabelFont: font('宋体', 15),
  },
  fixedTexts: { commitment: ['本人郑重承诺：', '一、所呈交论文系本人独立完成。'] },
}

async function sourceDocx(): Promise<Uint8Array> {
  const blank = await buildBlankDocx()
  const parsed = await parseDocx(blank)
  const p = (text: string, styleId?: string): SaveBlock => ({
    kind: 'generated',
    block: { type: styleId ? 'heading' : 'paragraph', ...(styleId ? { styleId, level: 1 } : {}), runs: [{ text }] },
  })
  const blocks: SaveBlock[] = [
    p('前言', 'Heading1'),
    p('选题背景的正文段落。'),
    p('相关研究现状。'),
    p('硬件设计', 'Heading1'),
    p('硬件系统框图说明。'),
    p('总结与展望', 'Heading1'),
    p('全文总结。'),
  ]
  return saveDocx(parsed, blocks)
}

async function unzipXml(bytes: Uint8Array): Promise<{ document: string; styles: string; parts: string[] }> {
  const zip = await JSZip.loadAsync(bytes)
  const document = await zip.file('word/document.xml')!.async('string')
  const styles = await zip.file('word/styles.xml')?.async('string') ?? ''
  const parts = Object.keys(zip.files)
  return { document, styles, parts }
}

describe('thesis-binder', () => {
  it('describeThesisBlocks 返回带层级的块大纲', async () => {
    const outline = await describeThesisBlocks(await sourceDocx())
    expect(outline.length).toBe(7)
    expect(outline[0]).toMatchObject({ kind: 'heading', level: 1, text: '前言' })
    expect(outline[1]!.kind).toBe('paragraph')
  })

  it('buildThesisDocx 装订出三节结构与全套部件', async () => {
    const source = await sourceDocx()
    const outline = await describeThesisBlocks(source)
    const spec: ThesisBindingSpecInput = {
      title: '基于ARM的嵌入式系统设计',
      coverFields: { title: '基于ARM的嵌入式系统设计', name: '' },
      abstractZh: '本设计完成了嵌入式系统的软硬件设计。',
      keywordsZh: ['嵌入式', 'ARM'],
      abstractEn: 'This design completes an embedded system.',
      keywordsEn: ['Embedded', 'ARM'],
      chapters: [
        { level: 1, headingText: '前言', blockStart: 0, blockEnd: 3 },
        { level: 2, headingText: '硬件系统框图', blockStart: 3, blockEnd: 4 },
        { level: 1, headingText: '总结与展望', blockStart: 5, blockEnd: 7 },
      ],
      references: ['[1] 作者. 书名[M]. 出版社, 2024.'],
      appendixBlocks: null,
      acknowledgementParagraphs: ['感谢导师指导。'],
      achievements: [],
    }
    const bytes = await buildThesisDocx(source, CARD, spec)
    const { document, styles, parts } = await unzipXml(bytes)

    // 部件齐全
    expect(document).toContain('嘉应学院')
    expect(document).toContain('本科毕业论文（设计）')
    expect(document).toContain('本人郑重承诺：')
    expect(document).toContain('摘要')
    expect(document).toContain('Abstract')
    expect(document).toContain('目录')
    expect(document).toContain('致谢')
    expect(document).toContain('感谢导师指导。')
    expect(document).toContain('参考文献')
    expect(document).toContain('[1] 作者. 书名[M]. 出版社, 2024.')

    // 章节编号 numeric-tree
    expect(document).toContain('1 前言')
    expect(document).toContain('1.1 硬件系统框图')
    expect(document).toContain('2 总结与展望')

    // 目录域
    expect(document).toContain('TOC \\o "1-3"')

    // 三节：封面（无页眉页脚引用）、前置（罗马页码）、正文（阿拉伯从 1）
    const sects = document.match(/<w:sectPr(?:\s[^>]*)?>[\s\S]*?<\/w:sectPr>/g) ?? []
    expect(sects.length).toBe(3)
    expect(sects[0]).not.toMatch(/headerReference|footerReference/)
    expect(sects[1]).toContain('w:fmt="upperRoman"')
    expect(sects[2]).toContain('w:start="1"')
    expect(sects[2]).not.toContain('upperRoman')

    // 页边距（2.5cm=1417twips，2cm=1134twips，装订线 5mm≈283twips）
    expect(sects[0]).toContain('w:top="1417"')
    expect(sects[0]).toContain('w:bottom="1417"')
    expect(sects[0]).toContain('w:left="1134"')
    expect(sects[0]).toContain('w:right="1134"')
    expect(sects[0]).toMatch(/w:gutter="28[0-9]"/)

    // 页眉页脚部件存在且内容正确
    const headerPart = parts.find((p) => /^word\/header\d+\.xml$/.test(p))
    const footerPart = parts.find((p) => /^word\/footer\d+\.xml$/.test(p))
    expect(headerPart).toBeTruthy()
    expect(footerPart).toBeTruthy()
    const headerXml = await (await JSZip.loadAsync(bytes)).file(headerPart!)!.async('string')
    expect(headerXml).toContain('基于ARM的嵌入式系统设计')

    // 标式：Heading1 黑体小二加粗
    expect(styles).toMatch(/w:styleId="Heading1"[^]*?黑体/)
    expect(styles).toMatch(/w:styleId="Heading1"[^]*?w:val="36"/)

    // 产物可被引擎重新解析
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks.length).toBeGreaterThan(10)
  })

  it('空摘要跳过部件、无附录不渲染附录标题', async () => {
    const source = await sourceDocx()
    const spec: ThesisBindingSpecInput = {
      title: '题目',
      coverFields: {},
      abstractZh: '',
      keywordsZh: [],
      abstractEn: '',
      keywordsEn: [],
      chapters: [{ level: 1, headingText: '前言', blockStart: 0, blockEnd: 7 }],
      references: [],
      appendixBlocks: null,
      acknowledgementParagraphs: [],
      achievements: [],
    }
    const bytes = await buildThesisDocx(source, CARD, spec)
    const { document } = await unzipXml(bytes)
    expect(document).not.toContain('关键词：')
    expect(document).not.toContain('Keywords: ')
    expect(document).not.toContain('<w:t>附录')
  })

  it('parts 分页规则生效（部件换页+每章换页）且固定行距写入 exact', async () => {
    const source = await sourceDocx()
    const spec: ThesisBindingSpecInput = {
      title: '基于ARM的嵌入式系统设计',
      coverFields: { title: '基于ARM的嵌入式系统设计' },
      abstractZh: '摘要正文。',
      keywordsZh: ['嵌入式'],
      abstractEn: 'Abstract body.',
      keywordsEn: ['Embedded'],
      chapters: [
        { level: 1, headingText: '前言', blockStart: 0, blockEnd: 3 },
        { level: 1, headingText: '总结与展望', blockStart: 5, blockEnd: 7 },
      ],
      references: ['[1] 作者. 书名[M]. 出版社, 2024.'],
      appendixBlocks: null,
      acknowledgementParagraphs: ['感谢导师指导。'],
      achievements: [],
    }
    const card: ThesisFormatCardInput = {
      ...CARD,
      base: {
        font: CARD.base.font,
        paragraph: { ...CARD.base.paragraph, lineSpacingExactPt: 22 },
      },
      parts: [
        { type: 'abstract_zh', title: '摘  要', titleFont: font('黑体', 16, true), breakBefore: 'page', optional: false },
        { type: 'abstract_en', title: 'Abstract', titleFont: font('黑体', 16, true), breakBefore: 'page', optional: false },
        { type: 'toc', title: '目  录', titleFont: font('黑体', 16, true), breakBefore: 'page', optional: false },
        { type: 'chapter', title: '', titleFont: font('黑体', 16, true), breakBefore: 'page', optional: false },
        { type: 'references', title: '参考文献', titleFont: font('黑体', 16, true), breakBefore: 'page', optional: false },
        { type: 'acknowledgement', title: '谢辞', titleFont: font('黑体', 16, true), breakBefore: 'page', optional: false },
      ],
    }
    const bytes = await buildThesisDocx(source, card, spec)
    const { document } = await unzipXml(bytes)

    // 自定义部件标题与标题字体生效
    expect(document).toContain('摘  要')
    expect(document).toContain('目  录')
    expect(document).toContain('谢辞')

    // 部件换页：承诺书（无规则不换）→ 摘要/Abstract(题目行)/目录/参考文献/致谢 → 5 处
    // 章级换页：两个一级章标题 → 2 处（首章在节首，Word 规范下不再额外分页，但属性仍写出）
    const breaks = document.match(/<w:pageBreakBefore\/>/g) ?? []
    expect(breaks.length).toBe(7)

    // 固定行距 22 磅 = 440 twips，lineRule=exact
    expect(document).toContain('w:lineRule="exact"')
    expect(document).toMatch(/w:line="440"/)
    expect(document).not.toMatch(/w:lineRule="auto"[^/]*\/><\/w:spacing>/)
  })
})

describe('thesis-binder · 未归类内容（v2 映射兜底）', () => {
  it('未匹配块渲染为正文末尾的独立小节，且红色高亮标记', async () => {
    const source = await sourceDocx()
    const spec: ThesisBindingSpecInput = {
      title: '开题与需求、系统设计报告',
      coverFields: { title: '开题与需求、系统设计报告' },
      abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
      chapters: [{ level: 1, headingText: '前言', blockStart: 0, blockEnd: 1 }],
      references: [],
      appendixBlocks: null,
      acknowledgementParagraphs: [],
      achievements: [],
      unmatchedBlocks: [{ start: 1, end: 3, label: '位于首个章节之前的说明段' }],
    }
    const bytes = await buildThesisDocx(source, CARD, spec)
    const { document } = await unzipXml(bytes)

    // 小节标题 + 说明都在，且带高亮
    expect(document).toContain('未归类内容（模板无对应槽位，请人工归位）')
    expect(document).toContain('〔未匹配：位于首个章节之前的说明段〕')
    expect(document).toContain('w:highlight w:val="yellow"')
    expect(document).toContain('w:color w:val="C00000"')
    // 未匹配块的内容没有丢
    expect(document).toContain('选题背景的正文段落。')
    expect(document).toContain('相关研究现状。')
    // 标记小节排在参考文献/致谢之后（正文末尾）
    expect(document.indexOf('未归类内容')).toBeGreaterThan(document.indexOf('致谢'))
  })

  it('没有未匹配块时不生成该小节', async () => {
    const source = await sourceDocx()
    const spec: ThesisBindingSpecInput = {
      title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
      chapters: [{ level: 1, headingText: '前言', blockStart: 0, blockEnd: 3 }],
      references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
    }
    const bytes = await buildThesisDocx(source, CARD, spec)
    const { document } = await unzipXml(bytes)
    expect(document).not.toContain('未归类内容')
  })
})
