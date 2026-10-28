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
  /** 待确认删除的卡片 id：面板内联二次确认（不用 window.confirm——它会抢焦点，
   *  确认框关闭时下拉面板被判定为「点到外部」而收起）。 */
  const [pendingDeleteId, setPendingDeleteId] = useState<string | null>(null)
  /** 重名模版待确认：非 null = 已存在同名模版，等用户选「替换原文件 / 取消上传」。 */
  const [pendingDuplicate, setPendingDuplicate] = useState<{ name: string } | null>(null)
  /** 排版作业：{ phase, progress, elapsed } → 面板顶部进度条；null = 空闲。 */
  const [job, setJob] = useState<{ phase: string; progress: number; elapsed: number } | null>(null)
  /** 目录融合模式：strict=只保留能对上模板目录的内容；blend=全部保留，未匹配的单独追加。 */
  const [mode, setMode] = useState<'strict' | 'blend'>('blend')
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
    setPendingDeleteId(null)
    refresh()
    return window.desktop.onThesisCardsChanged(refresh)
  }, [open, refresh])

  /**
   * 上传论文模版：只收 Word（用户预期「可导入 word 格式」）。
   * 解析成功后**直接对当前打开的文档排版**——导入模版的下一步就是排版，
   * 不再让用户多点一次（长任务最怕「点了没反应」）。
   */
  /**
   * 上传模版。
   *
   * 重名校验：同名模版已在列表里时**不直接覆盖**（解析要几分钟，覆盖了没法撤回），
   * 先返回 duplicate，让用户选「替换原文件 / 取消上传」。
   * overwrite=true = 用户已确认替换，复用主进程暂存的那份文件直接解析。
   */
  const uploadTemplate = async (overwrite = false) => {
    setUploading(true)
    setError(null)
    setReport(null)
    setPendingDuplicate(null)
    try {
      const outcome = await window.desktop.thesisPickTemplate(overwrite)
      if (!outcome.ok) {
        if (outcome.message) setError(outcome.message)
        return
      }
      const meta = outcome.value as {
        canceled?: boolean
        cardId?: string
        sourceName?: string
        factGaps?: string[]
        duplicate?: boolean
      } | undefined
      if (!meta || meta.canceled) return
      // 同名冲突：宿主还没解析，弹「替换 / 取消」
      if (meta.duplicate) {
        setPendingDuplicate({ name: meta.sourceName ?? '' })
        return
      }
      refresh()
      if (meta.factGaps && meta.factGaps.length > 0) {
        setError(t('aiThesisPartial', { gaps: meta.factGaps.join('；') }))
      }
      if (meta.cardId && hasDoc) await reformatWith(meta.cardId, { mode })
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setUploading(false)
    }
  }

  /**
   * 删除模版：确认后调后端删（模版文件 + 格式卡一起删，删干净不留残留），
   * 成功即从列表里移除该行。
   */
  const deleteTemplate = async (cardId: string) => {
    setPendingDeleteId(null)
    setDeletingId(cardId)
    setError(null)
    try {
      const outcome = await window.desktop.thesisDeleteTemplate(cardId, false)
      if (!outcome.ok) throw new Error(outcome.message || t('aiThesisFailed'))
      const result = outcome.value as { fileDeleted?: boolean; hadFile?: boolean } | null
      // 先把行从列表移除（用户要的即时反馈），再拉一次服务端确认
      setCards((prev) => (prev ? prev.filter((entry) => entry.id !== cardId) : prev))
      if (result?.hadFile && !result.fileDeleted) setError(t('aiThesisDeleteFileFailed'))
      onNotice?.(t('aiThesisDeleted'))
      refresh()
    } catch (err) {
      setError(errorMessage(err))
    } finally {
      setDeletingId(null)
    }
  }


  /**
   * 按模版排版当前文档（任务化）：
   * 提交作业立即返回 → 每 900ms 轮询阶段进度（parsing/mapping/rendering），
   * 面板顶部走真实进度条；10 分钟未完成主动放弃并给出明确提示。
   */
  const reformatWith = async (cardId: string, options?: {
    mode?: 'strict' | 'blend'
    keepLogo?: boolean
    keepHeaderFooter?: boolean
    headingMapping?: { chapter?: number; section?: number; subsection?: number }
  }) => {
    const startedAt = Date.now()
    const TIMEOUT_MS = 10 * 60 * 1000 // 10 分钟硬上限
    setBusy(true)
    setJob({ phase: 'queued', progress: 5, elapsed: 0 })
    try {
      const submit = await window.desktop.thesisReformatJob(cardId, options)
      if (!submit.ok) throw new Error(submit.message || t('aiThesisFailed'))
      const { jobId } = submit.value as { jobId: string }

      // 轮询循环：完成 / 失败 / 超时三种出口
      for (;;) {
        await new Promise((resolve) => setTimeout(resolve, 900))
        const elapsed = Date.now() - startedAt
        if (elapsed > TIMEOUT_MS) {
          throw new Error(t('aiThesisTimeout'))
        }
        const poll = await window.desktop.thesisReformatJobStatus(jobId)
        if (!poll.ok) throw new Error(poll.message || t('aiThesisFailed'))
        const snapshot = poll.value as {
          status: 'running' | 'completed' | 'failed'
          phase: string
          progress: number
          result?: unknown
          error?: string
        }
        setJob({ phase: snapshot.phase, progress: snapshot.progress, elapsed })
        if (snapshot.status === 'failed') {
          throw new Error(snapshot.error || t('aiThesisFailed'))
        }
        if (snapshot.status === 'completed') {
          const value = (snapshot.result ?? {}) as {
            originalName: string
            report?: {
              conservation?: { matched?: number; total?: number; unmatched?: number }
              report?: { appendedChapters?: number }
              needsReview?: unknown[]
            }
          }
          onNotice?.(t('aiThesisDone', { name: value.originalName }))
          const conservation = value.report?.conservation
          if (conservation) {
            setReport({
              matched: conservation.matched ?? 0,
              total: conservation.total ?? 0,
              unmatched: conservation.unmatched ?? 0,
              appended: value.report?.report?.appendedChapters ?? 0,
              needsReview: value.report?.needsReview?.length ?? 0,
            })
          }
          return
        }
      }
    } catch (err) {
      const message = errorMessage(err)
      setError(message === t('aiThesisTimeout') ? t('aiThesisTimeout') : message)
      setDropdown(() => 'thesis')
    } finally {
      setBusy(false)
      setJob(null)
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
                {job ? (
                  <div className="ai-thesis-progress" role="status">
                    <div className="ai-thesis-progress-bar">
                      <span style={{ width: `${Math.max(4, Math.min(100, job.progress))}%` }} />
                    </div>
                    <em>
                      {job.phase === 'parsing'
                        ? t('aiThesisPhaseParsing')
                        : job.phase === 'mapping'
                          ? t('aiThesisPhaseMapping')
                          : job.phase === 'rendering'
                            ? t('aiThesisPhaseRendering')
                            : t('aiThesisPhaseQueued')}
                      {' · '}
                      {Math.floor(job.elapsed / 1000)}s
                    </em>
                  </div>
                ) : null}
                {pendingDuplicate ? (
                  <div className="ai-thesis-dup" role="alertdialog" aria-label={t('aiThesisDuplicateTitle')}>
                    <p>
                      <b>{t('aiThesisDuplicateTitle')}</b>
                      <em>{pendingDuplicate.name}</em>
                    </p>
                    <div className="ai-thesis-dup-acts">
                      <button type="button" className="ai-thesis-dup-cancel" onClick={() => setPendingDuplicate(null)}>
                        {t('aiThesisDuplicateCancel')}
                      </button>
                      <button
                        type="button"
                        className="ai-thesis-dup-replace"
                        disabled={uploading}
                        onClick={() => void uploadTemplate(true)}
                      >
                        {uploading ? <span className="ai-thesis-spinner" /> : t('aiThesisDuplicateReplace')}
                      </button>
                    </div>
                  </div>
                ) : null}
                {cards === null ? (
                  <div className="ai-thesis-loading" aria-live="polite">
                    <span className="ai-thesis-spinner" />
                  </div>
                ) : noRoom ? (
                  <p className="ai-thesis-empty-hint">{t('aiThesisNoRoom')}</p>
                ) : cards.length === 0 ? (
                  <div className="ai-thesis-empty">
                    <p>{t('aiThesisPanelEmpty')}</p>
                    <label className="ai-thesis-mode">
                      <input type="radio" name="ai-thesis-mode" checked={mode === 'blend'} onChange={() => setMode('blend')} />
                      <span>{t('aiThesisModeBlend')}</span>
                    </label>
                    <label className="ai-thesis-mode">
                      <input type="radio" name="ai-thesis-mode" checked={mode === 'strict'} onChange={() => setMode('strict')} />
                      <span>{t('aiThesisModeStrict')}</span>
                    </label>
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
                            onClick={() => void reformatWith(row.id, { mode })}
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
                          {/* 删除模版：与预览按钮同尺寸（24×24、图标 14px），浅灰常态、
                              hover 转红；stopPropagation 防止触发父级的「用此模版排版」。
                              点击进入行内二次确认，面板保持打开。 */}
                          <button
                            type="button"
                            className="ai-thesis-card-act is-danger ai-thesis-card-del"
                            title={t('aiThesisRemove')}
                            aria-label={t('aiThesisRemove')}
                            disabled={deletingId === row.id}
                            onClick={(event) => {
                              event.stopPropagation();
                              event.preventDefault();
                              setPendingDeleteId((current) => (current === row.id ? null : row.id));
                            }}
                          >
                            {deletingId === row.id ? (
                              <span className="ai-thesis-spinner" />
                            ) : (
                              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round">
                                <path d="M3 6h18M8 6V4a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v2M19 6l-1 14a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1L5 6" />
                                <path d="M10 11v6M14 11v6" />
                              </svg>
                            )}
                          </button>

                        </div>
                      ))}
                    </div>
                    <div className="ai-thesis-mode-row">
                      <label className="ai-thesis-mode">
                        <input type="radio" name="ai-thesis-mode" checked={mode === 'blend'} onChange={() => setMode('blend')} />
                        <span>{t('aiThesisModeBlend')}</span>
                      </label>
                      <label className="ai-thesis-mode">
                        <input type="radio" name="ai-thesis-mode" checked={mode === 'strict'} onChange={() => setMode('strict')} />
                        <span>{t('aiThesisModeStrict')}</span>
                      </label>
                    </div>
                    <button className="ai-thesis-upload" disabled={uploading} onClick={() => void uploadTemplate()}>
                      {uploading ? <span className="ai-thesis-spinner" /> : t('aiThesisUpload')}
                    </button>
                  </>
                )}
                {pendingDeleteId ? (
                  <div
                    className="ai-thesis-modal-mask"
                    role="presentation"
                    onClick={() => setPendingDeleteId(null)}
                  >
                    <div
                      className="ai-thesis-modal"
                      role="alertdialog"
                      aria-modal="true"
                      aria-label={t('aiThesisDeleteConfirmTitle')}
                      onClick={(event) => event.stopPropagation()}
                    >
                      <p className="ai-thesis-modal-title">{t('aiThesisDeleteConfirmTitle')}</p>
                      <p className="ai-thesis-modal-text">{t('aiThesisDeleteConfirmNote')}</p>
                      <div className="ai-thesis-modal-acts">
                        <button
                          type="button"
                          className="ai-thesis-modal-cancel"
                          onClick={() => setPendingDeleteId(null)}
                        >
                          {t('aiThesisDeleteCancel')}
                        </button>
                        <button
                          type="button"
                          className="ai-thesis-modal-ok"
                          autoFocus
                          disabled={deletingId !== null}
                          onClick={() => void deleteTemplate(pendingDeleteId)}
                        >
                          {deletingId ? <span className="ai-thesis-spinner" /> : t('aiThesisDeleteConfirmOk')}
                        </button>
                      </div>
                    </div>
                  </div>
                ) : null}
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
