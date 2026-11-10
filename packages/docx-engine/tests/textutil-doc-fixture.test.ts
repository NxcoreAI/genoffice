import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse.js'

/**
 * macOS textutil（Cocoa 文本引擎）生成的 docx 必须能被预览/编辑器解析：
 * 未安装 LibreOffice 的机器上，.doc 模版预览靠 textutil 转 docx 兜底。
 * 这里用真实模版（广州城市理工学院《本科毕业设计（论文）外文翻译》.doc，
 * OLE2 复合文档）转换出的产物做回归——textutil 输出缺 styles.xml/numbering，
 * 解析器必须容错。
 */

const DOCX_FIXTURE = process.env.EVERROOM_TEXTUTIL_FIXTURE ?? ''

describe('textutil 转换产物解析', () => {
  it.skipIf(!DOCX_FIXTURE)('真实 .doc 模版经 textutil 转出的 docx 能解析出正文块', async () => {
    const bytes = new Uint8Array(readFileSync(DOCX_FIXTURE))
    const parsed = await parseDocx(bytes)
    const text = parsed.blocks
      .map((block) => (block.runs ?? []).map((run) => run.text ?? '').join(''))
      .join('\n')
    expect(text).toContain('广州城市理工学院')
    expect(text).toContain('本科毕业设计（论文）外文翻译')
    expect(parsed.blocks.length).toBeGreaterThan(5)
  })
})
