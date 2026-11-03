/**
 * EverRoom embed: the host's annotation queue for this document, shown beside
 * the page. Each row pairs the anchored original text with the user's note so
 * both are visible without re-reading the conversation; the footer sends the
 * whole batch (or drops one row) back to the host.
 *
 * Data comes from the host (`window.desktop.onAskQueue`); this panel never owns
 * the queue, so the conversation side and this panel always show the same items.
 */
import type { DocsAskQueueItem } from '../../shared/ipc'

interface Props {
  items: DocsAskQueueItem[]
  onFocus: (item: DocsAskQueueItem) => void
  onRemove: (id: string) => void
  onSendAll: () => void
}

const KIND_LABEL: Record<string, string> = {
  text: '正文',
  image: '图片',
  chart: '图表',
  table: '表格',
}

export function AskQueuePanel({ items, onFocus, onRemove, onSendAll }: Props): React.JSX.Element | null {
  if (items.length === 0) return null
  return (
    <aside className="ask-queue-panel" aria-label="待发送批注">
      <header className="ask-queue-head">
        <span className="ask-queue-title">待发送批注 {items.length} 条</span>
        <button type="button" className="ask-queue-send" onClick={onSendAll}>
          全部发送
        </button>
      </header>
      <ul className="ask-queue-list">
        {items.map((item, index) => (
          <li key={item.id} className="ask-queue-item">
            <div className="ask-queue-item-head">
              <span className="ask-queue-index">{index + 1}</span>
              <span className="ask-queue-kind">{KIND_LABEL[item.kind] ?? item.kind}</span>
              <button
                type="button"
                className="ask-queue-remove"
                aria-label="移除这条批注"
                onClick={() => onRemove(item.id)}
              >
                ✕
              </button>
            </div>
            {item.excerpt ? (
              <button
                type="button"
                className="ask-queue-excerpt"
                title="定位到原文"
                onClick={() => onFocus(item)}
              >
                「{item.excerpt}」
              </button>
            ) : null}
            <div className="ask-queue-note">{item.instruction}</div>
          </li>
        ))}
      </ul>
    </aside>
  )
}
