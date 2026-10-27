import { useCallback, useEffect, useState } from 'react'

import type { ThesisPanelCard } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { IconCaret } from './icons'
import { toggleDropdown, type SetDropdown } from './ribbon-tabs'

/*
 * EverRoom embed only: the "AI" ribbon tab hosts the thesis format panel
 * (first and — for now — only tool). Format cards, template upload, sample
 * preview and one-click reformat all run through the host bridge; the host
 * resolves room and file from this view, so no ids are passed here.
 */

const SIZE_PT: Record<string, number> = {
  初号: 42, 小初: 36, 一号: 26, 小一: 24, 二号: 22, 小二: 18,
  三号: 16, 小三: 15, 四号: 14, 小四: 12, 五号: 10.5, 小五: 9,
  六号: 7.5, 七号: 5.5,
}

function str(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value : fallback
}

function num(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback
}

function sizePt(value: unknown, fallback: number): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'string') return SIZE_PT[value] ?? fallback
  return fallback
}

function rec(value: unknown): Record<string, unknown> {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {}
}

function list(value: unknown): unknown[] {
  return Array.isArray(value) ? value : []
}

/** 卡的渲染投影：只取样张需要的字段，全部兜默认。 */
interface ThesisCardView {
  widthMm: number
  heightMm: number
  marginTopMm: number
  marginBottomMm: number
  marginLeftMm: number
  marginRightMm: number
  eastAsia: string
  ascii: string
  baseSizePt: number
  lineSpacingMultiple: number
  firstLineIndentChars: number
  h1: { eastAsia: string; ascii: string; sizePt: number; bold: boolean; align: string }
  schoolName: string
  degreeLine: string
  yearLine: string
  footerLine: string
  schoolSizePt: number
  schoolBold: boolean
  coverFields: string[]
  headerEnabled: boolean
  headerTemplate: string
  headerUnderline: boolean
}

function readThesisCardView(raw: Record<string, unknown>): ThesisCardView {
  const page = rec(raw.page)
  const base = rec(raw.base)
  const baseFont = rec(base.font)
  const baseParagraph = rec(base.paragraph)
  const headings = list(raw.headings).map(rec)
  const h1 = headings[0] ? rec(headings[0].font) : {}
  const heading = headings[0] ?? {}
  const cover = rec(raw.cover)
  const schoolFont = rec(cover.schoolFont)
  const header = rec(raw.header)
  return {
    widthMm: num(page.widthMm, 210),
    heightMm: num(page.heightMm, 297),
    marginTopMm: num(page.marginTopMm, 25),
    marginBottomMm: num(page.marginBottomMm, 25),
    marginLeftMm: num(page.marginLeftMm, 30),
    marginRightMm: num(page.marginRightMm, 30),
    eastAsia: str(baseFont.eastAsia, '宋体'),
    ascii: str(baseFont.ascii, 'Times New Roman'),
    baseSizePt: sizePt(baseFont.size, 12),
    lineSpacingMultiple: num(baseParagraph.lineSpacingMultiple, 1.25),
    firstLineIndentChars: num(baseParagraph.firstLineIndentChars, 2),
    h1: {
      eastAsia: str(h1.eastAsia, '黑体'),
      ascii: str(h1.ascii, 'Times New Roman'),
      sizePt: sizePt(h1.size, 16),
      bold: h1.bold !== false,
      align: str(heading.align, 'center'),
    },
    schoolName: str(cover.schoolName, '××大学'),
    degreeLine: str(cover.degreeLine, '本科毕业论文（设计）'),
    yearLine: str(cover.yearLine, '二〇二六年六月'),
    footerLine: str(cover.footerLine, ''),
    schoolSizePt: sizePt(schoolFont.size, 26),
    schoolBold: schoolFont.bold !== false,
    coverFields: list(cover.fields).map((field) => str(rec(field).label)).filter(Boolean).slice(0, 6),
    headerEnabled: header.enabled === true,
    headerTemplate: str(header.template, '{title}'),
    headerUnderline: header.underline === true,
  }
}

function paperFont(eastAsia: string, ascii: string): string {
  const cnAlias = eastAsia.includes('黑') ? '"Heiti SC"'
    : eastAsia.includes('楷') ? '"Kaiti SC"'
      : '"Songti SC"'
  return `"${ascii}", "${eastAsia}", ${cnAlias}, "PingFang SC", serif`
}

/** 面板列表里的缩略纸样：比例真实、内容为示意。纸样是文档内容，不随主题。 */
function ThesisMiniPaper({ card }: { card: ThesisCardView }) {
  const padX = (card.marginLeftMm / card.widthMm) * 100
  const padY = (card.marginTopMm / card.heightMm) * 100
  return (
    <span
      className="ai-thesis-mini-paper"
      style={{ aspectRatio: `${card.widthMm} / ${card.heightMm}`, padding: `${padY * 0.6}% ${padX}%` }}
      aria-hidden="true"
    >
      <span className="ai-thesis-mini-school" style={{ fontWeight: card.schoolBold ? 700 : 400 }}>
        {card.schoolName}
      </span>
      <span className="ai-thesis-mini-rule" />
      <span className="ai-thesis-mini-h1" style={{ fontWeight: card.h1.bold ? 700 : 400 }} />
      <span className="ai-thesis-mini-line is-w1" />
      <span className="ai-thesis-mini-line is-w2" />
      <span className="ai-thesis-mini-line is-w3" />
      <span className="ai-thesis-mini-page">1</span>
    </span>
  )
}

const SAMPLE_TITLE = '基于深度学习的×××研究'
const SAMPLE_PARAGRAPHS = [
  '本课题围绕×××展开研究。首先分析国内外研究现状，然后给出总体设计方案，最后通过实验验证方案的有效性。',
  '实验结果表明，所提出的方法在准确率与运行效率两方面均优于基线方案，达到了预期的设计目标。',
]

/** 大预览：封面页 + 正文页，按 mm/pt 真实比例换算（paperWidthPx 为纸宽像素）。 */
function ThesisFullPaper({ card }: { card: ThesisCardView }) {
  const paperWidthPx = 460
  const pxPerMm = paperWidthPx / card.widthMm
  const ptToPx = (pt: number) => pt * 0.3528 * pxPerMm
  const pageStyle = {
    width: paperWidthPx,
    height: Math.round(card.heightMm * pxPerMm),
    paddingTop: Math.round(card.marginTopMm * pxPerMm),
    paddingBottom: Math.round(card.marginBottomMm * pxPerMm),
    paddingLeft: Math.round(card.marginLeftMm * pxPerMm),
    paddingRight: Math.round(card.marginRightMm * pxPerMm),
  } as const
  const bodyFont = {
    fontFamily: paperFont(card.eastAsia, card.ascii),
    fontSize: ptToPx(card.baseSizePt),
    lineHeight: card.lineSpacingMultiple,
  } as const
  return (
    <div className="ai-thesis-full-wrap">
      <section className="ai-thesis-full-paper" style={pageStyle}>
        <div
          className="ai-thesis-full-school"
          style={{
            fontFamily: paperFont(card.eastAsia, card.ascii),
            fontSize: ptToPx(card.schoolSizePt),
            fontWeight: card.schoolBold ? 700 : 400,
          }}
        >
          {card.schoolName}
        </div>
        <div className="ai-thesis-full-degree" style={{ ...bodyFont, fontSize: ptToPx(card.baseSizePt + 4) }}>
          {card.degreeLine}
        </div>
        <div className="ai-thesis-full-fields">
          {card.coverFields.map((label) => (
            <span key={label} className="ai-thesis-full-field" style={bodyFont}>
              <em>{label}</em>
              <u aria-hidden="true" />
            </span>
          ))}
        </div>
        <div className="ai-thesis-full-year" style={bodyFont}>{card.yearLine}</div>
        {card.footerLine ? <div className="ai-thesis-full-cover-footer" style={bodyFont}>{card.footerLine}</div> : null}
      </section>
      <section className="ai-thesis-full-paper" style={pageStyle}>
        {card.headerEnabled ? (
          <header
            className="ai-thesis-full-header"
            style={{ ...bodyFont, borderBottom: card.headerUnderline ? '1px solid currentColor' : undefined }}
          >
            {card.headerTemplate.replace('{title}', SAMPLE_TITLE)}
          </header>
        ) : null}
        <h3
          className="ai-thesis-full-h1"
          style={{
            fontFamily: paperFont(card.h1.eastAsia, card.h1.ascii),
            fontSize: ptToPx(card.h1.sizePt),
            fontWeight: card.h1.bold ? 700 : 400,
            textAlign: card.h1.align === 'left' ? 'left' : 'center',
          }}
        >
          第一章&nbsp;&nbsp;绪论
        </h3>
        {SAMPLE_PARAGRAPHS.map((paragraph, index) => (
          <p key={index} className="ai-thesis-full-para" style={{ ...bodyFont, textIndent: `${card.firstLineIndentChars}em` }}>
            {paragraph}
          </p>
        ))}
        <footer className="ai-thesis-full-footer" style={bodyFont}>1</footer>
      </section>
    </div>
  )
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export function AiTab({ hasDoc, dropdown, setDropdown, onNotice }: {
  hasDoc: boolean
  dropdown: string | null
  setDropdown: SetDropdown
  onNotice?: (message: string) => void
}) {
  const { t, dateLocale } = useI18n()
  const [cards, setCards] = useState<ThesisPanelCard[] | null>(null)
  const [noRoom, setNoRoom] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [uploading, setUploading] = useState(false)
  const [removingId, setRemovingId] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [previewCard, setPreviewCard] = useState<Record<string, unknown> | null>(null)
  /** 正在删除模版的卡片 id（文件+卡，或只删文件）。 */
  const [deletingId, setDeletingId] = useState<string | null>(null)
  /** v2 槽位映射报告：排版完成后在面板里给出守恒/未匹配/待复核摘要。 */
  const [report, setReport] = useState<{
    matched?: number
    total?: number
    unmatched?: number
    appended?: number
    needsReview?: number
  } | null>(null)
  const open = dropdown === 'thesis'

  const refresh = useCallback(() => {
    void window.desktop.thesisListCards().then((outcome) => {
      if (outcome.ok) {
        setCards(outcome.value)
        setNoRoom(false)
      } else {
        setCards([])
        setNoRoom(outcome.reason === 'no_room')
      }
    })
  }, [])

  useEffect(() => {
    if (!open) return
    setCards(null)
    setError(null)
    setReport(null)
    refresh()
    return window.desktop.onThesisCardsChanged(refresh)
  }, [open, refresh])

  /**
   * 上传论文模版：只收 Word（用户预期「可导入 word 格式」）。
   * 解析成功后**直接对当前打开的文档排版**——导入模版的下一步就是排版，
   * 不再让用户多点一次（长任务最怕「点了没反应」）。
   */
  const uploadTemplate = async () => {
    setUploading(true)
    setError(null)
    setReport(null)
    try {
      const outcome = await window.desktop.thesisPickTemplate()
      if (!outcome.ok) {
        if (outcome.message) setError(outcome.message)
        return
      }
      const meta = outcome.value as { canceled?: boolean; cardId?: string; factGaps?: string[] } | undefined
      if (!meta || meta.canceled) return
      refresh()
      if (meta.factGaps && meta.factGaps.length > 0) {
        setError(t('aiThesisPartial', { gaps: meta.factGaps.join('；') }))
      }
      if (meta.cardId && hasDoc) await reformatWith(meta.cardId)
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setUploading(false)
    }
  }

  /** 删除模版：keepCard=false 连格式卡一起删（删干净）；true 只删文件、卡留着。 */
  const deleteTemplate = async (cardId: string, name: string, keepCard: boolean) => {
    const question = keepCard ? t('aiThesisDeleteFileConfirm', { name }) : t('aiThesisDeleteConfirm', { name })
    if (!window.confirm(question)) return
    setDeletingId(cardId)
    setError(null)
    try {
      const outcome = await window.desktop.thesisDeleteTemplate(cardId, keepCard)
      if (!outcome.ok) throw new Error(outcome.message || t('aiThesisFailed'))
      const result = outcome.value as { fileDeleted?: boolean; hadFile?: boolean; cardKept?: boolean } | null
      if (result?.hadFile && !result.fileDeleted) setError(t('aiThesisDeleteFileFailed'))
      onNotice?.(keepCard ? t('aiThesisFileDeleted') : t('aiThesisDeleted'))
      refresh()
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setDeletingId(null)
    }
  }

  const removeCard = async (cardId: string) => {
    setRemovingId(cardId)
    try {
      const outcome = await window.desktop.thesisRemoveCard(cardId)
      if (!outcome.ok && outcome.message) setError(outcome.message)
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setRemovingId(null)
    }
  }

  const reformatWith = async (cardId: string) => {
    setBusy(true)
    setDropdown(() => null)
    try {
      const outcome = await window.desktop.thesisReformat(cardId)
      if (!outcome.ok) throw new Error(outcome.message || t('aiThesisFailed'))
      const value = outcome.value as {
        originalName: string
        report?: {
          conservation?: { matched?: number; total?: number; unmatched?: number }
          report?: { appendedChapters?: number }
          needsReview?: unknown[]
        }
      }
      onNotice?.(t('aiThesisDone', { name: value.originalName }))
      // v2 映射报告：留在面板里（不打断，不弹窗）
      const conservation = value.report?.conservation
      if (conservation) {
        setReport({
          matched: conservation.matched ?? 0,
          total: conservation.total ?? 0,
          unmatched: conservation.unmatched ?? 0,
          appended: value.report?.report?.appendedChapters ?? 0,
          needsReview: value.report?.needsReview?.length ?? 0,
        })
        setDropdown(() => 'thesis')
      }
    } catch (err) {
      setError(errorMessage(err))
      setDropdown(() => 'thesis')
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc || busy}
              data-tip={t('aiThesisBtnTip')}
              onClick={() => toggleDropdown(setDropdown, 'thesis')}
            >
              <span className="rb-big-icon">
                <span className="ai-feature-icon" aria-hidden="true">
                  {busy ? (
                    <span className="ai-thesis-spinner" />
                  ) : (
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                      <path d="M15 4V2M15 16v-2M8 9h2M20 9h2M17.8 11.8L19 13M17.8 6.2L19 5M3 21L12.2 11.8M12.2 6.2L11 5" />
                    </svg>
                  )}
                </span>
                <IconCaret />
              </span>
              <span>{busy ? t('aiThesisRunning') : t('aiThesisBtn')}</span>
            </button>
            {open && (
              <div data-rb-panel="" className="layout-menu ai-thesis-panel">
                <p className="ai-thesis-title">{t('aiThesisPanelTitle')}</p>
                {cards === null ? (
                  <div className="ai-thesis-loading" aria-live="polite">
                    <span className="ai-thesis-spinner" />
                  </div>
                ) : noRoom ? (
                  <p className="ai-thesis-empty-hint">{t('aiThesisNoRoom')}</p>
                ) : cards.length === 0 ? (
                  <div className="ai-thesis-empty">
                    <p>{t('aiThesisPanelEmpty')}</p>
                    <button className="ai-thesis-upload" disabled={uploading} onClick={() => void uploadTemplate()}>
                      {uploading ? <span className="ai-thesis-spinner" /> : t('aiThesisUpload')}
                    </button>
                  </div>
                ) : (
                  <>
                    <div className="ai-thesis-cards" role="list">
                      {cards.map((row) => (
                        <div className="ai-thesis-card" key={row.id} role="listitem">
                          <button
                            className="ai-thesis-card-main"
                            disabled={busy}
                            title={t('aiThesisUseTip')}
                            onClick={() => void reformatWith(row.id)}
                          >
                            <ThesisMiniPaper card={readThesisCardView(row.card)} />
                            <span className="ai-thesis-card-body">
                              <b>{row.sourceName}</b>
                              <time>
                                {new Date(row.updatedAt).toLocaleString(dateLocale, { dateStyle: 'short', timeStyle: 'short' })}
                              </time>
                            </span>
                          </button>
                          <button
                            className="ai-thesis-card-act"
                            title={t('aiThesisPreview')}
                            onClick={() => setPreviewCard(row.card)}
                          >
                            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                              <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z" />
                              <circle cx="12" cy="12" r="3" />
                            </svg>
                          </button>
                          {row.sourceFileEntryId ? (
                            <button
                              className="ai-thesis-card-act"
                              title={t('aiThesisDeleteFileTip')}
                              disabled={deletingId === row.id}
                              onClick={() => void deleteTemplate(row.id, row.sourceName, true)}
                            >
                              {deletingId === row.id ? (
                                <span className="ai-thesis-spinner" />
                              ) : (
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                                  <path d="M12 3v12M7 10l5 5 5-5M5 21h14" />
                                </svg>
                              )}
                            </button>
                          ) : null}
                          <button
                            className="ai-thesis-card-act is-danger"
                            title={t('aiThesisRemove')}
                            disabled={deletingId === row.id}
                            onClick={() => void deleteTemplate(row.id, row.sourceName, false)}
                          >
                            {deletingId === row.id ? (
                              <span className="ai-thesis-spinner" />
                            ) : (
                              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
                                <path d="M3 6h18M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2M19 6l-1 14a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1L5 6" />
                              </svg>
                            )}
                          </button>
                        </div>
                      ))}
                    </div>
                    <button className="ai-thesis-upload" disabled={uploading} onClick={() => void uploadTemplate()}>
                      {uploading ? <span className="ai-thesis-spinner" /> : t('aiThesisUpload')}
                    </button>
                  </>
                )}
                {report && report.total ? (
                  <div className="ai-thesis-report" role="status">
                    <p>{t('aiThesisReport', { matched: report.matched ?? 0, total: report.total })}</p>
                    {report.unmatched ? <p className="is-warn">{t('aiThesisReportUnmatched', { count: report.unmatched })}</p> : null}
                    {report.needsReview ? <p className="is-warn">{t('aiThesisReportReview', { count: report.needsReview })}</p> : null}
                  </div>
                ) : null}
                {error ? <p className="ai-thesis-error">{error}</p> : null}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('aiThesisGroupLabel')}</div>
      </div>

      {previewCard ? (
        <div
          className="modal-backdrop"
          onMouseDown={(event) => event.target === event.currentTarget && setPreviewCard(null)}
        >
          <div className="modal ai-thesis-preview-modal">
            <h2>{t('aiThesisPreview')}</h2>
            <div className="ai-thesis-preview-body">
              <ThesisFullPaper card={readThesisCardView(previewCard)} />
            </div>
            <div className="modal-actions">
              <button className="btn" onClick={() => setPreviewCard(null)}>{t('appClose')}</button>
            </div>
          </div>
        </div>
      ) : null}
    </>
  )
}
