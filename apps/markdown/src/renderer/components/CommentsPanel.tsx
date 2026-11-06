import { memo, useEffect, useMemo, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import type { MdComment } from '../../shared/ipc'

function formatDate(iso?: string): string {
  if (!iso) return ''
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return ''
  return d.toLocaleString(undefined, {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  })
}

/**
 * Markdown comments pane (docs CommentsPanel parity): thread list + jump to the
 * anchored excerpt, a composer for New Comment, per-thread reply / resolve /
 * delete. The anchor text comes from the sidecar excerpt itself — the
 * decoration layer highlights the live range.
 */
export const CommentsPanel = memo(function CommentsPanel({
  comments,
  composing,
  onSubmitNew,
  onReply,
  onResolve,
  onCancelNew,
  onDelete,
  onJump,
  onClose,
}: {
  comments: MdComment[]
  /** show the new-comment composer (anchor captured by the caller) */
  composing: boolean
  onSubmitNew: (text: string) => void
  onReply: (parentId: string, text: string) => void
  /** resolve/reopen the whole thread */
  onResolve: (id: string, done: boolean) => void
  onCancelNew: () => void
  onDelete: (id: string) => void
  /** scroll the editor to this thread's anchor; false = dangling (dimmed) */
  onJump: (id: string) => boolean
  onClose: () => void
}) {
  const { t } = useI18n()
  const [activeId, setActiveId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const draftRef = useRef<HTMLTextAreaElement>(null)
  const [replyTo, setReplyTo] = useState<string | null>(null)
  const [replyDraft, setReplyDraft] = useState('')
  // Word parity: resolved comments are collapsed/hidden by default
  const [showResolved, setShowResolved] = useState(false)
  /** threads whose anchor no longer resolves (block deleted) */
  const [dangling, setDangling] = useState<Set<string>>(new Set())

  const topLevel = useMemo(() => comments.filter((c) => !c.parentId), [comments])
  const openThreads = topLevel.filter((c) => !c.done)
  const resolvedThreads = topLevel.filter((c) => c.done)
  const repliesOf = (id: string) => comments.filter((c) => c.parentId === id)

  const submitReply = () => {
    const text = replyDraft.trim()
    if (!text || !replyTo) return
    onReply(replyTo, text)
    setReplyDraft('')
    setReplyTo(null)
  }

  useEffect(() => {
    if (composing) draftRef.current?.focus()
  }, [composing])

  const submit = () => {
    const text = draft.trim()
    if (!text) return
    onSubmitNew(text)
    setDraft('')
  }

  const renderThread = (c: MdComment) => {
    const replies = repliesOf(c.id)
    const isDangling = dangling.has(c.id)
    return (
      <div key={c.id} className={`comment-thread ${c.done ? 'resolved' : ''}`}>
        <button
          className={`comment-card ${activeId === c.id ? 'active' : ''} ${isDangling ? 'dangling' : ''}`}
          onClick={() => {
            setActiveId(c.id)
            const ok = onJump(c.id)
            setDangling((prev) => {
              if (ok === prev.has(c.id)) return prev
              const next = new Set(prev)
              if (ok) next.delete(c.id)
              else next.add(c.id)
              return next
            })
          }}
        >
          <div className="comment-card-head">
            <span className="comment-avatar">{(c.initials || c.author || '?').slice(0, 2)}</span>
            <span className="comment-author">{c.author || t('appUnknownAuthor')}</span>
            <span className="comment-date">{formatDate(c.date)}</span>
            {c.done && <span className="comment-resolved-badge">{t('appResolved')}</span>}
            <span
              className="comment-card-del"
              data-tip={
                replies.length > 0 ? t('appDeleteCommentWithReplies') : t('appDeleteComment')
              }
              role="button"
              aria-label={t('appDeleteComment')}
              tabIndex={0}
              onClick={(e) => {
                e.stopPropagation()
                onDelete(c.id)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.stopPropagation()
                  onDelete(c.id)
                }
              }}
            >
              ✕
            </span>
          </div>
          {c.excerpt && <div className="comment-anchor">“{c.excerpt.slice(0, 60)}”</div>}
          <div className="comment-text">{c.text}</div>
        </button>
        {replies.map((r) => (
          <div key={r.id} className="comment-reply">
            <div className="comment-card-head">
              <span className="comment-avatar">{(r.initials || r.author || '?').slice(0, 2)}</span>
              <span className="comment-author">{r.author || t('appUnknownAuthor')}</span>
              <span className="comment-date">{formatDate(r.date)}</span>
              <span
                className="comment-card-del"
                data-tip={t('appDeleteReply')}
                aria-label={t('appDeleteReply')}
                role="button"
                tabIndex={0}
                onClick={() => onDelete(r.id)}
              >
                ✕
              </span>
            </div>
            <div className="comment-text">{r.text}</div>
          </div>
        ))}
        <div className="comment-thread-actions">
          {replyTo === c.id ? (
            <div className="comment-reply-compose">
              <textarea
                autoFocus
                placeholder={t('appReplyPlaceholder')}
                value={replyDraft}
                onChange={(e) => setReplyDraft(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submitReply()
                  if (e.key === 'Escape') setReplyTo(null)
                }}
              />
              <div className="comment-compose-actions">
                <button onClick={() => setReplyTo(null)}>{t('appCancel')}</button>
                <button className="primary" disabled={!replyDraft.trim()} onClick={submitReply}>
                  {t('appReply')}
                </button>
              </div>
            </div>
          ) : (
            <>
              <button onClick={() => setReplyTo(c.id)}>{t('appReply')}</button>
              <button onClick={() => onResolve(c.id, !c.done)}>
                {c.done ? t('appReopen') : t('appResolve')}
              </button>
            </>
          )}
        </div>
      </div>
    )
  }

  return (
    <aside className="comments-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">{t('appCommentsTitle', { n: comments.length })}</span>
        <button
          className="comments-pane-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      {composing && (
        <div className="comment-compose">
          <textarea
            ref={draftRef}
            placeholder={t('appCommentPlaceholder')}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
              if (e.key === 'Escape') onCancelNew()
            }}
          />
          <div className="comment-compose-actions">
            <button onClick={onCancelNew}>{t('appCancel')}</button>
            <button className="primary" disabled={!draft.trim()} onClick={submit}>
              {t('appCommentBtn')}
            </button>
          </div>
        </div>
      )}
      <div className="comments-pane-list">
        {openThreads.map(renderThread)}
        {resolvedThreads.length > 0 && (
          <div className="comments-resolved-group">
            <button className="comments-resolved-toggle" onClick={() => setShowResolved((v) => !v)}>
              <span className={`comments-resolved-caret ${showResolved ? 'open' : ''}`}>▸</span>
              {t('appResolvedComments', { n: resolvedThreads.length })}
            </button>
            {showResolved && resolvedThreads.map(renderThread)}
          </div>
        )}
        {comments.length === 0 && !composing && (
          <div className="comments-empty">{t('appNoComments')}</div>
        )}
      </div>
    </aside>
  )
})
