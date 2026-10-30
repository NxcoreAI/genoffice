import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse.js'

/**
 * 「手打编号」文档兜底：整篇不用 Word 标题样式时（学校模版里极常见），
 * 排版链路要能按文字认出章节；编辑器/预览仍保持原行为（不传选项就不改块类型）。
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function para(text: string, pPr = ''): string {
  return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}

async function buildDocx(body: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file('word/document.xml', `${XML_DECL}<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`)
  return zip.generateAsync({ type: 'uint8array' })
}

const headingsOf = (parsed: { blocks: Array<{ type: string; level?: number; runs?: Array<{ text?: string }> }> }) =>
  parsed.blocks
    .filter((block) => block.type === 'heading' && (block.level ?? 0) > 0)
    .map((block) => `${block.level}:${(block.runs ?? []).map((run) => run.text ?? '').join('').trim()}`)

describe('手打编号标题兜底', () => {
  it('默认不改变块类型（编辑器/预览行为不变）', async () => {
    const bytes = await buildDocx(
      [para('第一章 前言'), para('1.1 课程设计的目标'), para('一、实验目的'), para('参考文献')].join(''),
    )
    expect(headingsOf(await parseDocx(bytes))).toEqual([])
  })

  it('排版链路开启后按文字认章节并定级', async () => {
    const bytes = await buildDocx(
      [
        para('第一章 前言'),
        para('1.1 课程设计的目标'),
        para('1.2课程设计的简介'),
        para('第二章 设计内容'),
        para('一、实验目的'),
        para('（二）具体要求'),
        para('参考文献'),
        para('附录 A 源程序'),
      ].join(''),
    )
    const heads = headingsOf(await parseDocx(bytes, { patternHeadings: true }))
    expect(heads).toEqual([
      '1:第一章 前言',
      '2:1.1 课程设计的目标',
      '2:1.2课程设计的简介',
      '1:第二章 设计内容',
      '1:一、实验目的',
      '2:（二）具体要求',
      '1:参考文献',
      '1:附录 A 源程序',
    ])
  })

  it('不误判：小数开头/带缩进的正文不算标题', async () => {
    const bytes = await buildDocx(
      [
        para('第一章 前言'),
        para('1.5 倍行距的正文段落'),
        para('3.5万元预算的说明'),
        para('一、这一段有首行缩进', '<w:ind w:firstLine="420"/>'),
        para('第一章 前言……1'),
      ].join(''),
    )
    const heads = headingsOf(await parseDocx(bytes, { patternHeadings: true }))
    expect(heads).toEqual(['1:第一章 前言'])
  })

  it('表格里的「一、」「1.1」不是标题', async () => {
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para('一、评分标准')}</w:tc></w:tr></w:tbl>`
    const bytes = await buildDocx([para('第一章 前言'), table].join(''))
    const heads = headingsOf(await parseDocx(bytes, { patternHeadings: true }))
    expect(heads).toEqual(['1:第一章 前言'])
  })
})

describe('Heading 样式 + 自动编号（被解析成 listItem）', () => {
  it('排版链路把它当章节标题，且不会把下一节正文误当成「本节已有内容」', async () => {
    const body = [
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>实验一 闪烁灯</w:t></w:r></w:p>',
      para('六、 实验分析与总结'),
      para('这里是第二节的正文，不应该算作上一节的内容'),
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>实验二 流水灯</w:t></w:r></w:p>',
      para('实验二正文'),
    ].join('')
    // Heading1 在 styles 里没有 numPr 时仍是 heading；本用例断言默认行为不被破坏
    const bytes = await buildDocx(body)
    const heads = headingsOf(await parseDocx(bytes, { patternHeadings: true }))
    expect(heads).toContain('1:实验一 闪烁灯')
    expect(heads).toContain('1:六、 实验分析与总结')
  })
})
