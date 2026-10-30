import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { buildBlankDocx } from '../src/blank'
import { parseDocx } from '../src/parse'
import { saveDocx, type SaveBlock } from '../src/patch'
import {
  buildThesisDocxOnTemplate,
  type ThesisTemplateBinderInput,
} from '../src/thesis-template-binder'

const font = (eastAsia: string, size: number, bold = false) => ({
  eastAsia, ascii: 'Times New Roman', size, bold,
})

const CARD: ThesisTemplateBinderInput['card'] = {
  page: {
    widthMm: 210, heightMm: 297, orientation: 'portrait',
    marginTopMm: 25, marginBottomMm: 25, marginLeftMm: 20, marginRightMm: 20, gutterMm: 5,
  },
  base: {
    font: font('宋体', 12),
    paragraph: { align: 'justify', firstLineIndentChars: 2, lineSpacingMultiple: 1.25 },
  },
  header: { enabled: false, template: '', font: font('宋体', 10.5), underline: false },
  footer: { enabled: false, font: font('宋体', 10.5), frontMatterFormat: 'roman' },
  headings: [
    { level: 1, font: font('黑体', 18, true), align: 'center', numbering: 'none' },
    { level: 2, font: font('黑体', 15, true), align: 'left', numbering: 'none' },
  ],
  abstract: {
    titleFont: font('黑体', 18, true),
    bodyFont: font('宋体', 12),
    bodyParagraph: { align: 'justify', firstLineIndentChars: 2 },
    keywordPrefix: '关键词：',
    keywordPrefixBold: true,
    keywordSeparator: '，',
  },
  toc: { enabled: false, levels: '1-3', titleFont: font('黑体', 18, true) },
  references: { listFont: font('宋体', 12), listParagraph: { align: 'justify' } },
  cover: {
    schoolName: 'XX大学',
    degreeLine: '',
    yearLine: '',
    fields: [],
    footerLine: '',
    schoolFont: font('宋体', 26, true),
    fieldLabelFont: font('宋体', 15),
  },
  fixedTexts: { commitment: [] },
  profile: {
    slots: [
      { id: 's1', title: '一、项目任务' },
      { id: 's2', title: '二、项目需求分析' },
    ],
  },
}

function para(text: string, styleId?: string): SaveBlock {
  return {
    kind: 'generated',
    block: {
      type: styleId ? 'heading' : 'paragraph',
      ...(styleId ? { styleId, level: 1 } : {}),
      runs: [{ text }],
    },
  }
}


/** 多列表格：rows[0] 是表头。拆列（验证结果 → 测试结果）测试用。 */
function gridTable(rows: string[][]): SaveBlock {
  const cells = (row: string[]): string =>
    row.map((text) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>`
      + `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`).join('')
  const grid = `<w:tblGrid>${rows[0]!.map(() => '<w:gridCol w:w="2000"/>').join('')}</w:tblGrid>`
  return {
    kind: 'xml',
    xml: '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
      + grid
      + rows.map((row) => `<w:tr>${cells(row)}</w:tr>`).join('')
      + '</w:tbl>',
  }
}

function table(cellText: string): SaveBlock {
  return {
    kind: 'xml',
    xml:
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
      + '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
      + `<w:p><w:r><w:t>${cellText}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  }
}

/** 论文模版：封面 + 核心信息区表格 + 摘要 + 目录 + 教师评价表 + 两个章节槽位 + 参考文献。 */
async function templateDocx(): Promise<Uint8Array> {
  const blank = await buildBlankDocx()
  const parsed = await parseDocx(blank)
  return saveDocx(parsed, [
    para('XX大学'),
    table('核心信息区 题目｜姓名｜学号'),
    para('摘要', 'Heading1'),
    para('（此处填写摘要）'),
    para('目录', 'Heading1'),
    para('一、项目任务（目录项）'),
    table('教师评价表 评分｜评语'),
    para('一、项目任务', 'Heading1'),
    para('（此处写项目任务）'),
    para('二、项目需求分析', 'Heading1'),
    para('（此处写需求分析）'),
    para('参考文献', 'Heading1'),
    para('（此处列参考文献）'),
  ])
}

async function sourceDocx(): Promise<Uint8Array> {
  const blank = await buildBlankDocx()
  const parsed = await parseDocx(blank)
  return saveDocx(parsed, [
    para('业务场景分析', 'Heading1'),
    para('源正文一：业务背景。'),
    table('源表格 用户角色｜权限'),
    para('需求梳理', 'Heading1'),
    para('源正文二：需求清单。'),
  ])
}

async function unzipDocument(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return zip.file('word/document.xml')!.async('string')
}

async function build(): Promise<string> {
  const template = await templateDocx()
  const source = await sourceDocx()
  const { bytes } = await buildThesisDocxOnTemplate({
    templateBytes: template,
    sourceBytes: source,
    card: CARD,
    spec: {
      title: '我的论文',
      coverFields: { title: '我的论文' },
      abstractZh: '这是摘要正文。',
      keywordsZh: ['关键词一', '关键词二'],
      abstractEn: '',
      keywordsEn: [],
      chapters: [
        { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
        { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 's2' },
      ],
      references: ['[1] 参考文献条目。'],
      appendixBlocks: null,
      acknowledgementParagraphs: [],
      achievements: [],
    },
  })
  return unzipDocument(bytes)
}

describe('buildThesisDocxOnTemplate（模板骨架装订）', () => {
  it('封面与核心信息区原样保留，且顺序不变', async () => {
    const document = await build()
    expect(document).toContain('XX大学')
    expect(document).toContain('核心信息区')
    expect(document.indexOf('XX大学')).toBeLessThan(document.indexOf('核心信息区'))
    expect(document.indexOf('核心信息区')).toBeLessThan(document.indexOf('摘要'))
  })

  it('模版固定表单（教师评价表）留在原位——不再被搬到目录/正文之后', async () => {
    const document = await build()
    expect(document).toContain('教师评价表')
    // 关键回归：评价表在「目录」之后、第一章之前，而不是被追加到正文尾部
    expect(document.indexOf('教师评价表')).toBeGreaterThan(document.indexOf('目录'))
    expect(document.indexOf('教师评价表')).toBeLessThan(document.indexOf('业务场景分析'))
    expect(document.indexOf('教师评价表')).toBeLessThan(document.indexOf('参考文献'))
  })

  it('槽位标题换成源章节标题，模版的手工编号前缀保留', async () => {
    const document = await build()
    expect(document).toContain('一、业务场景分析')
    expect(document).toContain('二、需求梳理')
    // 模版占位文案被替换掉，不残留
    expect(document).not.toContain('（此处写项目任务）')
    expect(document).not.toContain('（此处写需求分析）')
  })

  it('源正文与源表格进入对应槽位', async () => {
    const document = await build()
    expect(document).toContain('源正文一：业务背景。')
    expect(document).toContain('源表格')
    expect(document).toContain('源正文二：需求清单。')
    // 源内容落在对应章节标题之后
    expect(document.indexOf('业务场景分析')).toBeLessThan(document.indexOf('源正文一：业务背景。'))
    expect(document.indexOf('需求梳理')).toBeLessThan(document.indexOf('源正文二：需求清单。'))
  })

  it('摘要与参考文献落到模版自己的位置，不另起章节', async () => {
    const document = await build()
    expect(document).toContain('这是摘要正文。')
    expect(document).toContain('[1] 参考文献条目。')
    expect(document.indexOf('这是摘要正文。')).toBeLessThan(document.indexOf('目录'))
    expect(document.indexOf('[1] 参考文献条目。')).toBeGreaterThan(document.indexOf('需求梳理'))
  })

  it('模版里没有对应槽位的内容不会被吞掉（未归类小节）', async () => {
    const template = await templateDocx()
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: CARD,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
        unmatchedBlocks: [{ start: 3, end: 5, label: '需求梳理' }],
      },
    })
    const document = await unzipDocument(bytes)
    expect(document).toContain('未归类内容')
    expect(document).toContain('源正文二：需求清单。')
  })

  it('模版目录项即使套了 Heading 样式，也不会被当成章节锚点', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('目录', 'Heading1'),
      para('一、项目任务........1', 'Heading1'),
      para('二、项目需求分析......3', 'Heading1'),
      para('一、项目任务', 'Heading1'),
      para('（此处写项目任务）'),
      para('二、项目需求分析', 'Heading1'),
      para('（此处写需求分析）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: CARD,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 's2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 目录项原样保留
    expect(document).toContain('一、项目任务........1')
    // 正文落在真正的章节标题处，而不是被插进目录区
    expect(document.indexOf('源正文一：业务背景。')).toBeGreaterThan(document.indexOf('二、项目需求分析......3'))
    expect(document.indexOf('源正文二：需求清单。')).toBeGreaterThan(document.indexOf('源正文一：业务背景。'))
    expect(document).not.toContain('（此处写项目任务）')
    expect(document).not.toContain('（此处写需求分析）')
  })

  it('模版已有手工目录标题时不再插第二个目录', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('目录', 'Heading1'),
      para('一、项目任务........1', 'Heading1'),
      para('一、项目任务', 'Heading1'),
      para('（此处写项目任务）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true } },
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const out = await parseDocx(bytes)
    const tocHeadings = out.blocks.filter((block) => (
      block.type === 'heading' && (block.runs ?? []).map((run) => run.text).join('').trim() === '目录'
    ))
    expect(tocHeadings).toHaveLength(1)
  })

  it('模版用普通段落排标题时，前置表单表格留在目录前、正文落到章节处', async () => {
    // 复刻真实中文论文模版：章节标题不套 Heading 样式，是「普通段落 + 手工编号」；
    // 核心信息区下方紧跟项目经理评价表、指导老师评分表，然后才是目录和正文。
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      table('核心信息区 题目｜姓名｜学号'),
      para('项目经理评价'),
      table('项目经理评价表 成员｜任务'),
      para('指导教师评分及评语'),
      table('教师评分表 评语｜评分'),
      para('目录'),
      para('一、项目任务........1'),
      para('二、项目需求分析......3'),
      para('一、项目任务'),
      para('（此处写项目任务）'),
      para('二、项目需求分析'),
      para('（此处写需求分析）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: CARD,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 's2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)

    // 关键回归：两张前置表单表格都还在目录之前（不再被搬到正文首章/文末）
    expect(document).toContain('项目经理评价表')
    expect(document).toContain('教师评分表')
    expect(document.indexOf('项目经理评价表')).toBeLessThan(document.indexOf('目录'))
    expect(document.indexOf('教师评分表')).toBeLessThan(document.indexOf('目录'))
    // 顺序：核心信息区 → 项目经理评价表 → 教师评分表 → 目录
    expect(document.indexOf('核心信息区')).toBeLessThan(document.indexOf('项目经理评价表'))
    expect(document.indexOf('项目经理评价表')).toBeLessThan(document.indexOf('教师评分表'))

    // 正文落到普通段落标题之后，而不是插进前置表单里
    expect(document.indexOf('一、业务场景分析')).toBeGreaterThan(document.indexOf('目录'))
    expect(document.indexOf('源正文一：业务背景。')).toBeGreaterThan(document.indexOf('一、业务场景分析'))
    expect(document.indexOf('源正文二：需求清单。')).toBeGreaterThan(document.indexOf('二、需求梳理'))
    expect(document).not.toContain('（此处写项目任务）')
  })

  it('卡片槽位为空（旧卡片解析结果）时按编号标题顺序兜底，表格仍留在目录前', async () => {
    // 老卡片的 profile.slots 是旧代码算出来的空数组，存在卡片里不会自动重算。
    // 没有这条兜底，整篇会回落 v2，两张前置表单表格又会被搬走。
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      table('核心信息区'),
      para('项目经理评价'),
      table('项目经理评价表'),
      para('指导教师评分及评语'),
      table('教师评分表'),
      para('目录'),
      para('一、项目任务........1'),
      para('二、项目需求分析......3'),
      para('一、项目任务'),
      para('（此处写项目任务）'),
      para('二、项目需求分析'),
      para('（此处写需求分析）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, profile: { slots: [] } },
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3 },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5 },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 前置表单表格全部留在目录之前
    expect(document.indexOf('项目经理评价表')).toBeLessThan(document.indexOf('目录'))
    expect(document.indexOf('教师评分表')).toBeLessThan(document.indexOf('目录'))
    // 正文落进模板的两个章节位置
    expect(document.indexOf('一、业务场景分析')).toBeGreaterThan(document.indexOf('目录'))
    expect(document.indexOf('源正文一：业务背景。')).toBeGreaterThan(document.indexOf('一、业务场景分析'))
    expect(document.indexOf('源正文二：需求清单。')).toBeGreaterThan(document.indexOf('二、需求梳理'))
  })

  it('模版一个槽位锚点都没有时，整篇正文按顺序追加到文末且不刷红字', async () => {
    // 模版只解析出「摘要」一个锚点，两个章节的槽位标题都不存在
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('摘要', 'Heading1'),
      para('（此处填写摘要）'),
    ])
    const source = await sourceDocx()
    const result = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: CARD,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '这是摘要正文。', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 's2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(result.bytes)
    // 诊断要如实告诉上层「这次没认出章节骨架」
    expect(result.diagnostics.appendedOnly).toBe(true)
    expect(result.diagnostics.anchors).toBeGreaterThanOrEqual(1)
    // 摘要照常落位
    expect(document).toContain('这是摘要正文。')
    // 两个章节的内容与标题都在，且按顺序出现
    expect(document).toContain('源正文一：业务背景。')
    expect(document).toContain('源表格')
    expect(document).toContain('源正文二：需求清单。')
    expect(document).toContain('业务场景分析')
    expect(document).toContain('需求梳理')
    expect(document.indexOf('业务场景分析')).toBeLessThan(document.indexOf('需求梳理'))
    // 中性过渡标题 + 不刷红字（整篇红字比排版错乱更吓人）
    expect(document).toContain('模版未识别章节位置')
    expect(document).not.toContain('未归类内容')
    expect(document).not.toContain('C00000')
  })

  it('部分章节缺锚点时，缺的那章显性标红补到文末', async () => {
    // 模版有「业务场景分析」这一个真槽位，但没有「需求梳理」
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('摘要', 'Heading1'),
      para('（此处填写摘要）'),
      para('业务场景分析', 'Heading1'),
      para('（此处填写业务场景）'),
    ])
    const source = await sourceDocx()
    const result = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: {
        ...CARD,
        profile: { slots: [{ id: 's1', title: '业务场景分析', level: 1 }] },
      } as typeof CARD,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '这是摘要正文。', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 's2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(result.bytes)
    expect(result.diagnostics.appendedOnly).toBe(false)
    expect(document).toContain('源正文一：业务背景。')
    expect(document).toContain('源正文二：需求清单。')
    // 缺锚点的那一章要有显性标记，不能静默消失
    expect(document).toContain('未归类内容')
    expect(document).toContain('未落位')
    expect(document).toContain('C00000')
  })

  it('模版用无编号普通段落排标题（slots=0）时，正文仍落到对应章节位置', async () => {
    // 真实场景：中文模版章节标题是普通段落、连手工编号都没有，卡片剖面解析为空，
    // 网关只能 R4-append 下发 slotId='slot-append'。此时不能整篇堆到文末。
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('XX大学'),
      table('核心信息区 题目｜姓名｜学号'),
      para('摘要', 'Heading1'),
      para('（此处填写摘要）'),
      para('目录', 'Heading1'),
      para('项目任务…………1'),
      para('项目需求分析…………5'),
      table('教师评价表 评分｜评语'),
      para('项目任务'),
      para('（此处写项目任务）'),
      para('项目需求分析'),
      para('（此处写需求分析）'),
      para('参考文献', 'Heading1'),
      para('（此处列参考文献）'),
    ])
    const source = await sourceDocx()
    const result = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, profile: { slots: [] } } as typeof CARD,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '这是摘要正文。', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 'slot-append' },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 'slot-append' },
        ],
        references: ['[1] 参考文献条目。'], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(result.bytes)
    expect(result.diagnostics.slots).toBe(0)
    expect(result.diagnostics.appendedOnly).toBe(false)
    // 核心信息区表格仍在最前，教师评价表仍在目录之后、正文之前
    expect(document.indexOf('核心信息区')).toBeLessThan(document.indexOf('目录'))
    expect(document.indexOf('教师评价表')).toBeLessThan(document.indexOf('业务场景分析'))
    // 两章正文分别落在模版对应章节处，而不是堆到参考文献后面
    expect(document.indexOf('业务场景分析')).toBeLessThan(document.indexOf('需求梳理'))
    expect(document.indexOf('源正文二：需求清单。')).toBeLessThan(document.indexOf('参考文献'))
    expect(document).not.toContain('未归类内容')
  })

  it('封面节的分节符不被目录清理吞掉——封面保持自己的页边距/页眉/页码', async () => {
    // 真实模版的坑：封面节的分节符段落正好落在「目录旧条目区」的末尾
    // （目录标题之后、第一章之前），目录替换会把它当过期目录行删掉，
    // 封面随即并进正文节——核心信息区被挤错位、拆到第二页。
    const blank = await buildBlankDocx()
    const parsed = await parseDocx(blank)
    const coverSectionBreak =
      '<w:p><w:pPr><w:sectPr>'
      + '<w:footerReference r:id="rId4" w:type="default"/>'
      + '<w:pgSz w:w="11906" w:h="16838"/>'
      + '<w:pgMar w:top="1440" w:right="1247" w:bottom="1440" w:left="1247" w:header="851" w:footer="992" w:gutter="0"/>'
      + '<w:pgNumType w:fmt="upperRoman"/>'
      + '</w:sectPr></w:pPr></w:p>'
    const template = await saveDocx(parsed, [
      para('XX大学'),
      table('核心信息区 题目｜姓名｜学号'),
      para('目录', 'Heading1'),
      para('一、项目任务……1'), // 手工目录行（引导点+页码，isTocEntry 命中）
      { kind: 'xml', xml: coverSectionBreak }, // 目录条目区末尾的分节符
      para('一、项目任务', 'Heading1'),
      para('（此处写项目任务）'),
      para('二、项目需求分析', 'Heading1'),
      para('（此处写需求分析）'),
      para('参考文献', 'Heading1'),
      para('（此处列参考文献）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      // 目录替换开启才会走进「丢弃目录旧条目」的清理路径——分节符正是在那里被吞掉的
      card: { ...CARD, toc: { ...CARD.toc, enabled: true } },
      spec: {
        title: '我的论文', coverFields: { title: '我的论文' },
        abstractZh: '这是摘要正文。', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 's2' },
        ],
        references: ['[1] 参考文献条目。'], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 输出仍是两个节，封面节的页边距/页码方案原样保留
    const sects = document.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g) ?? []
    expect(sects.length).toBe(2)
    expect(sects[0]).toContain('upperRoman')
    expect(sects[0]).toContain('w:left="1247"')
    // 封面字段与核心信息区都在封面节内（首个分节符之前）
    const firstSectPos = document.indexOf('<w:sectPr')
    expect(document.indexOf('核心信息区')).toBeGreaterThan(-1)
    expect(document.indexOf('核心信息区')).toBeLessThan(firstSectPos)
    // 分节符在目录条目之后、第一章正文之前
    expect(document.indexOf(coverSectionBreak.slice(0, 60))).toBeGreaterThan(document.indexOf('目录'))
    expect(document.indexOf(coverSectionBreak.slice(0, 60))).toBeLessThan(document.indexOf('源正文一：业务背景。'))
  })

  it('模版有章节位但论文没内容：保留空标题，且写进目录（正文与目录一样多）', async () => {
    const blank = await buildBlankDocx()
    const parsed = await parseDocx(blank)
    const template = await saveDocx(parsed, [
      para('XX大学'),
      table('核心信息区 题目｜姓名｜学号'),
      para('目录', 'Heading1'),
      para('一、项目任务……1'),
      para('二、项目需求分析……2'),
      para('一、项目任务', 'Heading1'),
      para('（此处写项目任务）'),
      para('二、项目需求分析', 'Heading1'),
      para('（此处写需求分析）'),
      para('参考文献', 'Heading1'),
      para('（此处列参考文献）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true } },
      spec: {
        title: '我的论文', coverFields: { title: '我的论文' },
        abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        // 只填 s1；s2 这个模版章节位本次论文没有内容
        chapters: [{ level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' }],
        references: ['[1] 参考文献条目。'], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 保留模版的空章节标题（正文里真的有这个标题，不只是目录里）
    const bodyHeadings = document
      .split('</w:p>')
      // 排除目录条目（TOC 样式），只看真正的正文段落
      .filter((paragraph) => !paragraph.includes('w:val="TOC'))
      .map((paragraph) => [...paragraph.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((run) => run[1]).join(''))
    expect(bodyHeadings.some((heading) => heading.includes('项目需求分析'))).toBe(true)
    // 示例正文（占位说明）仍然要清掉
    expect(document).not.toContain('（此处写需求分析）')
    // 目录里也要有它：目录条目数 = 正文一级标题数 = 2
    const tocEntries = (document.match(/w:val="TOC1"/g) ?? []).length
    expect(tocEntries).toBe(2)
    const tocTexts = document
      .split('</w:p>')
      .filter((paragraph) => paragraph.includes('w:val="TOC1"'))
      .map((paragraph) => [...paragraph.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((run) => run[1]).join(''))
    expect(tocTexts.join('|')).toContain('项目需求分析')
  })

  it('文末兜底补出的章节也登记进目录（目录不少最后一章）', async () => {
    const blank = await buildBlankDocx()
    const parsed = await parseDocx(blank)
    const template = await saveDocx(parsed, [
      para('XX大学'),
      table('核心信息区 题目｜姓名｜学号'),
      para('目录', 'Heading1'),
      para('一、项目任务……1'),
      para('一、项目任务', 'Heading1'),
      para('（此处写项目任务）'),
      para('二、项目需求分析', 'Heading1'),
      para('（此处写需求分析）'),
      para('参考文献', 'Heading1'),
      para('（此处列参考文献）'),
    ])
    const source = await sourceDocx()
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true } },
      spec: {
        title: '我的论文', coverFields: { title: '我的论文' },
        abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 3, slotId: 's1' },
          // 模版里没有这个槽位 → 文末兜底补出；它的目录条目也必须在
          { level: 1, headingText: '需求梳理', blockStart: 3, blockEnd: 5, slotId: 'slot-missing' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    const orphanAt = document.indexOf('未归类内容')
    expect(orphanAt).toBeGreaterThan(-1)
    // 目录在正文之前：兜底章节的标题必须出现在目录区里
    expect(document.slice(0, orphanAt)).toContain('需求梳理')
  })

  it('模版优先模式：目录/标题/模版文字原样保留，内容按小节归类', async () => {
    const blank = await buildBlankDocx()
    const parsed = await parseDocx(blank)
    const template = await saveDocx(parsed, [
      para('XX大学'),
      para('目录', 'Heading1'),
      para('一、项目任务……1'),
      para('1.1 背景……1'),
      para('一、项目任务', 'Heading1'),
      para('（模版说明：本节写项目背景）'),
      para('1.1 背景', 'Heading2'),
      para('1.2 目标', 'Heading2'),
      para('二、项目需求分析', 'Heading1'),
      para('（模版说明：本节写需求）'),
    ])
    const srcBlank = await buildBlankDocx()
    const srcParsed = await parseDocx(srcBlank)
    const source = await saveDocx(srcParsed, [
      para('一、业务场景分析', 'Heading2'),
      para('1.1 政策与业务背景', 'Heading3'),
      para('源正文：政策背景内容。'),
      para('1.2 解决思路', 'Heading3'),
      para('源正文：解决思路内容。'),
      para('二、需求梳理', 'Heading2'),
      para('源正文：需求内容。'),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true }, profile: { ...CARD.profile!, mode: 'template-first' } },
      spec: {
        title: '论文', coverFields: {},
        abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '业务场景分析', blockStart: 0, blockEnd: 5, slotId: 's1' },
          { level: 1, headingText: '需求梳理', blockStart: 5, blockEnd: 7, slotId: 's2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 目录条目是模版自己的，没被重算
    expect(document).toContain('一、项目任务……1')
    expect(document).not.toContain('业务场景分析')
    // 模版自带说明文字保留
    expect(document).toContain('（模版说明：本节写项目背景）')
    expect(document).toContain('（模版说明：本节写需求）')
    // 内容按小节归类：政策背景落到 1.1 背景 之后、1.2 目标 之前
    const p11 = document.indexOf('1.1 背景')
    const p12 = document.indexOf('1.2 目标')
    const bg = document.indexOf('政策与业务背景')
    const bgBody = document.indexOf('源正文：政策背景内容。')
    const goal = document.indexOf('源正文：解决思路内容。')
    expect(p11).toBeGreaterThan(-1)
    expect(bg).toBeGreaterThan(p11)
    expect(bg).toBeLessThan(p12)
    expect(bgBody).toBeGreaterThan(bg)
    expect(goal).toBeGreaterThan(p12)
  })

  it('模版优先模式：标题用词不同也能按语义归类（里程碑→项目计划，不是→测试结果）', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('目录', 'Heading1'),
      para('一、项目任务……1'),
      para('一、项目任务', 'Heading1'),
      para('1.2 目标', 'Heading2'),
      para('1.4 项目计划', 'Heading2'),
      para('六、项目测试结果', 'Heading1'),
    ])
    const srcBlank = await buildBlankDocx()
    const source = await saveDocx(await parseDocx(srcBlank), [
      para('六、验证与里程碑（摘要）', 'Heading2'),
      para('源正文：里程碑表内容。'),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true }, profile: { ...CARD.profile!, mode: 'template-first' } },
      spec: {
        title: '论文', coverFields: {},
        abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '验证与里程碑', blockStart: 0, blockEnd: 2 }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    const plan = document.indexOf('1.4 项目计划')
    const milestone = document.indexOf('源正文：里程碑表内容。')
    const testChapter = document.indexOf('六、项目测试结果')
    expect(plan).toBeGreaterThan(-1)
    expect(milestone).toBeGreaterThan(plan)
    // 落在「六、项目测试结果」之前（否则说明被「验证」二字带跑了）
    expect(milestone).toBeLessThan(testChapter)
  })

  it('拆分发牌：表格里的「验证结果」列搬进测试结果小节，原表删掉该列', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('目录', 'Heading1'),
      para('一、项目任务……1'),
      para('1.4 项目计划', 'Heading2'),
      para('六、项目测试结果', 'Heading1'),
    ])
    const srcBlank = await buildBlankDocx()
    const source = await saveDocx(await parseDocx(srcBlank), [
      para('六、验证与里程碑（摘要）', 'Heading2'),
      gridTable([
        ['里程碑', '交付物', '验证结果'],
        ['M1 需求与原型', '本报告', '链路打通'],
        ['M2 安全与多角色', '登录三入口', '安全测试 11 项通过'],
      ]),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true }, profile: { ...CARD.profile!, mode: 'template-first' } },
      spec: {
        title: '论文', coverFields: {},
        abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '验证与里程碑', blockStart: 0, blockEnd: 3 }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 六、项目测试结果 拿到两行「里程碑：验证结果」
    expect(document).toContain('M1 需求与原型：链路打通')
    expect(document).toContain('M2 安全与多角色：安全测试 11 项通过')
    // 原表只剩两列：验证结果 表头被删掉
    expect(document).not.toContain('验证结果</w:t>')
    expect(document).toContain('里程碑</w:t>')
    expect(document).toContain('交付物</w:t>')
  })

  it('空小节留一行浅灰提示，不静默空白', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('一、项目任务', 'Heading1'),
      para('1.3 项目组成员分工及职责概览', 'Heading2'),
      para('1.4 项目计划', 'Heading2'),
      para('七、项目总结与展望', 'Heading1'),
    ])
    const srcBlank = await buildBlankDocx()
    const source = await saveDocx(await parseDocx(srcBlank), [
      para('六、验证与里程碑（摘要）', 'Heading2'),
      para('源正文：里程碑说明。'),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: source,
      card: { ...CARD, toc: { ...CARD.toc, enabled: true }, profile: { ...CARD.profile!, mode: 'template-first' } },
      spec: {
        title: '论文', coverFields: {},
        abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '验证与里程碑', blockStart: 0, blockEnd: 2 }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const document = await unzipDocument(bytes)
    // 论文里没有对应内容的两节各留一行提示；1.4 拿到了「里程碑」内容，不该有提示
    expect(document).toContain('本节待补充：1.3 项目组成员分工及职责概览')
    expect(document).toContain('本节待补充：七、项目总结与展望')
    expect(document).not.toContain('本节待补充：1.4 项目计划')
  })

  it('AI 补写：只有开关给了内容才插入，且带「AI 生成」标记', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('一、项目任务', 'Heading1'),
      para('七、项目总结与展望', 'Heading1'),
    ])
    const srcBlank = await buildBlankDocx()
    const source = await saveDocx(await parseDocx(srcBlank), [
      para('一、项目任务', 'Heading2'),
      para('源正文：任务说明。'),
    ])
    const spec = {
      title: '论文', coverFields: {},
      abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
      chapters: [{ level: 1, headingText: '项目任务', blockStart: 0, blockEnd: 2 }],
      references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
    }
    const card = { ...CARD, toc: { ...CARD.toc, enabled: true }, profile: { ...CARD.profile!, mode: 'template-first' } }
    // 关着：只有灰字提示
    const off = await buildThesisDocxOnTemplate({ templateBytes: template, sourceBytes: source, card, spec })
    const offDoc = await unzipDocument(off.bytes)
    expect(offDoc).toContain('本节待补充：七、项目总结与展望')
    expect(offDoc).not.toContain('AI 生成')
    // 开着：插入补写正文 + 标记，不再出提示
    const on = await buildThesisDocxOnTemplate({
      templateBytes: template, sourceBytes: source, card,
      spec: { ...spec, generatedSections: [{ title: '七、项目总结与展望', text: '本项目完成了任务说明所述工作，并形成了可交付成果。' }] },
    })
    const onDoc = await unzipDocument(on.bytes)
    expect(onDoc).toContain('AI补写，请认真核对')
    expect(onDoc).toContain('本项目完成了任务说明所述工作')
    expect(onDoc).not.toContain('本节待补充：七、项目总结与展望')
  })
})

/**
 * 「手打编号」模版兜底：学校模版常常整篇不用 Word 标题样式，
 * 章节标题就是普通段落（「第一章 前言」「1.1 目标」）。
 * 排版链路必须按文字认出这些标题，否则骨架版式完全用不上。
 */
describe('手打编号模版：按文字识别章节', () => {
  const plainCard: ThesisTemplateBinderInput['card'] = {
    ...CARD,
    profile: {
      mode: 'template-first',
      slots: [
        { id: 'c1', title: '第一章 前言' },
        { id: 'c2', title: '第二章 设计内容' },
      ],
      subSlots: [{ id: 'c1.1', parentId: 'c1', title: '1.1 课程设计的目标' }],
    } as ThesisTemplateBinderInput['card']['profile'],
  }

  async function plainTemplate(): Promise<Uint8Array> {
    const blank = await buildBlankDocx()
    return saveDocx(await parseDocx(blank), [
      para('XX大学'),
      para('目    录'),
      para('第一章 前言 1'),
      para('第二章 设计内容 3'),
      // 正文里的「标题」全是普通段落，没有 Heading 样式
      para('第一章 前言'),
      para('（此处写前言）'),
      para('1.1 课程设计的目标'),
      para('（此处写目标）'),
      para('第二章 设计内容'),
      para('（此处写设计内容）'),
      para('参考文献'),
      para('（此处列参考文献）'),
    ])
  }

  async function plainSource(): Promise<Uint8Array> {
    const blank = await buildBlankDocx()
    return saveDocx(await parseDocx(blank), [
      para('第一章 前言'),
      para('源正文：课题背景与意义。'),
      para('1.1 课程设计的目标'),
      para('源小节：掌握单片机最小系统。'),
      para('第二章 设计内容'),
      para('源正文：硬件与软件设计。'),
    ])
  }

  it('普通段落写的章节标题也能成为归类锚点', async () => {
    const { bytes, diagnostics } = await buildThesisDocxOnTemplate({
      templateBytes: await plainTemplate(),
      sourceBytes: await plainSource(),
      card: plainCard,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '第一章 前言', blockStart: 0, blockEnd: 2, slotId: 'c1' },
          { level: 1, headingText: '第二章 设计内容', blockStart: 4, blockEnd: 6, slotId: 'c2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    expect(diagnostics.templateFirstDowngraded).toBeFalsy()
    const document = await unzipDocument(bytes)
    expect(document).toContain('源正文：课题背景与意义。')
    expect(document).toContain('源小节：掌握单片机最小系统。')
    expect(document).toContain('源正文：硬件与软件设计。')
    // 内容落在各自章节之下，而不是全堆到文末
    // （lastIndexOf：目录里也有同名条目，这里要断言的是正文里的那个标题）
    expect(document.lastIndexOf('第一章 前言')).toBeLessThan(document.indexOf('源正文：课题背景与意义。'))
    expect(document.indexOf('源正文：课题背景与意义。')).toBeLessThan(document.lastIndexOf('第二章 设计内容'))
  })

  it('「只有目录、正文没有标题」的模版：降级重建并标记，而不是把内容全塞进一节', async () => {
    const blank = await buildBlankDocx()
    const tocOnlyTemplate = await saveDocx(await parseDocx(blank), [
      para('XX大学'),
      para('目    录'),
      para('第一章 前言 1'),
      para('第二章 设计内容 3'),
      para('（此处是格式说明）'),
    ])
    const { diagnostics } = await buildThesisDocxOnTemplate({
      templateBytes: tocOnlyTemplate,
      sourceBytes: await plainSource(),
      card: plainCard,
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [
          { level: 1, headingText: '第一章 前言', blockStart: 0, blockEnd: 2, slotId: 'c1' },
          { level: 1, headingText: '第二章 设计内容', blockStart: 4, blockEnd: 5, slotId: 'c2' },
        ],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    expect(diagnostics.templateFirstDowngraded).toBe(true)
  })
})

describe('章节标题下的模版固定块（实验日期/评分表）不能被正文挤走', () => {
  it('正文插在固定块之后、下一节标题之前', async () => {
    const blank = await buildBlankDocx()
    const template = await saveDocx(await parseDocx(blank), [
      para('目录', 'Heading1'),
      para('实验一 闪烁灯……1'),
      para('实验一 闪烁灯', 'Heading1'),
      para('实验日期：    年   月   日第   至   节；指导老师（现场签章）：'),
      table('实验评分表 环节｜占分｜评分'),
      para('★实验预习与准备：按要求课前预习。'),
      para('实验目的', 'Heading2'),
      para('（此处写实验目的）'),
    ])
    const src = await saveDocx(await parseDocx(await buildBlankDocx()), [
      para('实验一 闪烁灯', 'Heading1'),
      para('源正文：流水灯控制原理与实现。'),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: src,
      card: {
        ...CARD,
        profile: {
          mode: 'template-first',
          slots: [{ id: 'e1', title: '实验一 闪烁灯' }],
          subSlots: [{ id: 'e1.1', parentId: 'e1', title: '实验目的' }],
        } as ThesisTemplateBinderInput['card']['profile'],
      },
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '实验一 闪烁灯', blockStart: 0, blockEnd: 2, slotId: 'e1' }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const doc = await unzipDocument(bytes)
    const title = doc.indexOf('实验一 闪烁灯')
    const dateLine = doc.indexOf('实验日期')
    const sheet = doc.indexOf('实验评分表')
    const body = doc.indexOf('源正文：流水灯控制原理与实现。')
    // 固定块紧跟标题
    expect(dateLine).toBeGreaterThan(title)
    expect(sheet).toBeGreaterThan(dateLine)
    // 正文在固定块之后
    expect(body).toBeGreaterThan(sheet)
  })
})

describe('三级小标题（1.1.1）', () => {
  it('也能接内容、也能出「本节待补充」提示', async () => {
    const template = await saveDocx(await parseDocx(await buildBlankDocx()), [
      para('目录', 'Heading1'),
      para('1.1 任务概述……1'),
      para('1.1.1 任务概述……1'),
      para('1.1 任务概述', 'Heading2'),
      para('1.1.1 任务概述', 'Heading3'),
      para('1.1.2 任务需求分析', 'Heading3'),
    ])
    const src = await saveDocx(await parseDocx(await buildBlankDocx()), [
      para('1.1 任务概述', 'Heading2'),
      para('1.1.1 任务概述', 'Heading3'),
      para('源正文：任务背景与目标。'),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: src,
      card: { ...CARD, profile: { mode: 'template-first', slots: [{ id: 's1', title: '1.1 任务概述' }] } as ThesisTemplateBinderInput['card']['profile'] },
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [{ level: 1, headingText: '1.1 任务概述', blockStart: 0, blockEnd: 3, slotId: 's1' }],
        references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
      },
    })
    const doc = await unzipDocument(bytes)
    expect(doc).toContain('本节待补充：1.1.2 任务需求分析')
    expect(doc).toContain('源正文：任务背景与目标。')
  })
})

describe('AI 补写：同名小节按「章｜小节」各写各的', () => {
  it('两个实验的「实验步骤」拿到不同内容', async () => {
    const template = await saveDocx(await parseDocx(await buildBlankDocx()), [
      para('实验一 闪烁灯', 'Heading1'),
      para('实验步骤', 'Heading2'),
      para('实验二 流水灯', 'Heading1'),
      para('实验步骤', 'Heading2'),
    ])
    const { bytes } = await buildThesisDocxOnTemplate({
      templateBytes: template,
      sourceBytes: template,
      card: { ...CARD, profile: { mode: 'template-first', slots: [{ id: 's1', title: '实验一 闪烁灯' }] } as ThesisTemplateBinderInput['card']['profile'] },
      spec: {
        title: '论文', coverFields: {}, abstractZh: '', keywordsZh: [], abstractEn: '', keywordsEn: [],
        chapters: [], references: [], appendixBlocks: null, acknowledgementParagraphs: [], achievements: [],
        generatedSections: [
          { title: '实验一 闪烁灯｜实验步骤', text: '实验一的步骤：点亮第一个 LED 并测量电流。' },
          { title: '实验二 流水灯｜实验步骤', text: '实验二的步骤：用循环移位实现流水灯效果。' },
        ],
      },
    })
    const doc = await unzipDocument(bytes)
    expect(doc).toContain('实验一的步骤')
    expect(doc).toContain('实验二的步骤')
    // 两段内容各归各的章节，不能互相串
    expect(doc.indexOf('实验一的步骤')).toBeLessThan(doc.indexOf('实验二 流水灯'))
    expect(doc.indexOf('实验二的步骤')).toBeGreaterThan(doc.indexOf('实验二 流水灯'))
  })
})
