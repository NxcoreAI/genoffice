import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'

import './side-dock.css'

/**
 * EverRoom host side dock (Agent / Materials / History), embedded into each
 * GenOffice editor when the host opens the view with the `dock=1` query
 * parameter. Self-contained by design: no app dependencies, inline icons,
 * mock data included — the host wires the real Agent bridge in later stages.
 */

export type SideDockKind = 'docs' | 'sheets' | 'slides' | 'pdf'

/* ───────────── mock 数据（按编辑器类型） ───────────── */

interface MaterialItem {
  id: string
  kind: 'mail' | 'meeting' | 'wiki' | 'fact' | 'file'
  title: string
  excerpt: string
  meta: string
  addedLabel: string
}

interface HistoryEntry {
  id: string
  actor: 'user' | 'agent' | 'system'
  title: string
  detail?: string
  restorable?: boolean
  diffable?: boolean
}

const MATERIAL_ICONS: Record<MaterialItem['kind'], string> = {
  mail: '✉️',
  meeting: '🗓️',
  wiki: '📘',
  fact: '💡',
  file: '📄',
}

const MOCK_MATERIALS: Record<SideDockKind, MaterialItem[]> = {
  docs: [
    { id: 'm1', kind: 'mail', title: '设计周报：V1 视觉定稿', excerpt: 'V1 视觉已定稿，动效统一 240ms，列表类内容做 40ms 错峰。', meta: '邮件 · 林薇 · 2 天前', addedLabel: '2 天前添加' },
    { id: 'm2', kind: 'meeting', title: '周会：连接器排期', excerpt: '阶段一范围确认：身份合并 + 命名对齐，联调窗口 10/28–11/05。', meta: '会议 · 52 分钟 · 3 天前', addedLabel: '3 天前添加' },
    { id: 'm3', kind: 'wiki', title: 'Wiki · 双链路收敛', excerpt: 'Gmail 双链路并入统一格式层，映射表三处同值待收口。', meta: '知识库 · Wiki', addedLabel: '上周添加' },
  ],
  sheets: [
    { id: 'm1', kind: 'mail', title: '供应商报价（Q4）', excerpt: '三家供应商报价见正文，请确认后回复，交期以报价单为准。', meta: '邮件 · 采购部 · 5 小时前', addedLabel: '今天添加' },
    { id: 'm2', kind: 'fact', title: '事实 · 预算上限', excerpt: 'Q4 采购预算上限 50 万，超限需 VP 审批。', meta: '自动提炼 · 9/20', addedLabel: '上周添加' },
  ],
  slides: [
    { id: 'm1', kind: 'wiki', title: 'Wiki · V1 视觉规范', excerpt: '字阶 L1–L4、240ms 动效、颜色守门清单。', meta: '设计系统 · Wiki', addedLabel: '上周添加' },
    { id: 'm2', kind: 'meeting', title: '发布评审 0926', excerpt: '11 月第一周灰度、第二周全量，回滚预案齐备。', meta: '会议 · 28 分钟', addedLabel: '上周添加' },
  ],
  pdf: [
    { id: 'm1', kind: 'wiki', title: 'Wiki · 竞品跟踪', excerpt: 'Notion AI 定价与能力清单（每月更新）。', meta: '设计系统 · Wiki', addedLabel: '上周添加' },
  ],
}

const MOCK_HISTORY: Record<SideDockKind, HistoryEntry[]> = {
  docs: [
    { id: 'h1', actor: 'user', title: '王小雨 编辑', detail: '修改 3 处 · 已保存', restorable: true },
    { id: 'h2', actor: 'agent', title: 'Agent 修改', detail: '插入引用 ×2 · 已审阅通过', diffable: true },
    { id: 'h3', actor: 'system', title: '导入', detail: '来源：邮件附件' },
  ],
  sheets: [
    { id: 'h1', actor: 'system', title: '自动成表', detail: '来源：邮件《供应商报价（Q4）》' },
    { id: 'h2', actor: 'agent', title: 'Agent 修改', detail: '补全合计行 · 已审阅通过', diffable: true },
  ],
  slides: [
    { id: 'h1', actor: 'user', title: '王小雨 编辑', detail: '调整时间线页 · 已保存', restorable: true },
    { id: 'h2', actor: 'agent', title: 'AI 生成初稿', detail: '依据 3 份材料生成 3 页', restorable: true, diffable: true },
  ],
  pdf: [
    { id: 'h1', actor: 'user', title: '王小雨 标注', detail: '标注结论段落', restorable: true },
    { id: 'h2', actor: 'system', title: '导入', detail: '来源：本地文件' },
  ],
}

const STARTER_PROMPTS: Record<SideDockKind, string[]> = {
  docs: ['把这份文档的结论改得更精炼', '根据本周邮件更新「下周计划」', '检查全文的引用是否还有效'],
  sheets: ['哪个区域的数据异常？', '按总价排序并标红最高项', '把这张表写成一段采购建议'],
  slides: ['把时间线页改成表格', '给每页要点补一句风险提示', '整体换成更收敛的配色'],
  pdf: ['这份 PDF 的核心结论是什么？', '把第 2 页的观点改写进我的周报', '提炼三句给管理层看的摘要'],
}

const KIND_LABEL: Record<SideDockKind, string> = {
  docs: '文档',
  sheets: '表格',
  slides: '演示',
  pdf: 'PDF',
}

/* ───────────── 图标（内联 SVG，避免跨 app 依赖） ───────────── */

const icon = (path: ReactNode, size = 17): ReactNode => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    {path}
  </svg>
)

const AGENT_ICON = icon(<><rect x="4" y="8" width="16" height="12" rx="3" /><path d="M12 8V4" /><circle cx="9" cy="14" r="1" /><circle cx="15" cy="14" r="1" /></>)
const MATERIALS_ICON = icon(<><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" /></>)
const HISTORY_ICON = icon(<><path d="M3 12a9 9 0 1 0 3-6.7" /><path d="M3 4v4h4" /><path d="M12 8v4l3 2" /></>)
const PLUS_ICON = icon(<><path d="M12 5v14" /><path d="M5 12h14" /></>, 14)
const X_ICON = icon(<><path d="M18 6 6 18" /><path d="m6 6 12 12" /></>, 12)
const CHEVRON_ICON = icon(<><path d="m6 9 6 6 6-6" /></>, 13)
const SEND_ICON = icon(<><path d="M20 13a8 8 0 1 1-8-8" /><path d="M12 13l4-4" /><path d="M16 9h-4" /><path d="M20 3v5h-5" /></>, 13)
const RESTORE_ICON = icon(<><path d="M3 12a9 9 0 1 0 3-6.7" /><path d="M3 4v4h4" /></>, 12)
const DIFF_ICON = icon(<><path d="M12 3v14" /><path d="M8 7h8" /><path d="M8 17h8" /><path d="M6 21h12" /></>, 12)
const USER_ICON = icon(<><circle cx="12" cy="8" r="4" /><path d="M4 21c0-4 3.6-6 8-6s8 2 8 6" /></>, 11)

/* ───────────── 组件 ───────────── */

type DockTab = 'agent' | 'materials' | 'history'

const TABS: Array<{ key: DockTab; label: string; icon: ReactNode }> = [
  { key: 'agent', label: 'Agent', icon: AGENT_ICON },
  { key: 'materials', label: '材料', icon: MATERIALS_ICON },
  { key: 'history', label: '历史记录', icon: HISTORY_ICON },
]

let idSeq = 0
const nextId = (): string => `evsd-${++idSeq}`

export function SideDock({ kind, fileName }: { kind: SideDockKind; fileName?: string }) {
  const [tab, setTab] = useState<DockTab>('agent')
  const [collapsed, setCollapsed] = useState(false)
  const [width, setWidth] = useState(360)
  const [toast, setToast] = useState<string | null>(null)
  const toastTimer = useRef<number | null>(null)

  useEffect(() => {
    if (!toast) return
    toastTimer.current = window.setTimeout(() => setToast(null), 2200)
    return () => {
      if (toastTimer.current !== null) window.clearTimeout(toastTimer.current)
    }
  }, [toast])

  const say = (message: string): void => setToast(message)

  const startResize = (e: React.PointerEvent): void => {
    e.preventDefault()
    const startX = e.clientX
    const startWidth = width
    const onMove = (ev: PointerEvent): void => {
      setWidth(Math.min(560, Math.max(280, startWidth + (startX - ev.clientX))))
    }
    const onUp = (): void => {
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', onUp)
    }
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', onUp)
  }

  const activate = (key: DockTab): void => {
    if (!collapsed && tab === key) setCollapsed(true)
    else {
      setTab(key)
      setCollapsed(false)
    }
  }

  return (
    <div className={`evsd-root${collapsed ? ' collapsed' : ''}`} style={collapsed ? undefined : { width }}>
      {!collapsed && (
        <>
          <div className="evsd-resize" onPointerDown={startResize} role="separator" aria-orientation="vertical" />
          <div className="evsd-body">
            {tab === 'agent' && <AgentView kind={kind} fileName={fileName} say={say} />}
            {tab === 'materials' && <MaterialsView kind={kind} say={say} />}
            {tab === 'history' && <HistoryView kind={kind} say={say} />}
          </div>
        </>
      )}
      <div className="evsd-strip">
        {TABS.map((t) => (
          <button
            key={t.key}
            type="button"
            className={`evsd-strip-btn${!collapsed && tab === t.key ? ' active' : ''}`}
            title={t.label}
            onClick={() => activate(t.key)}
          >
            {t.icon}
          </button>
        ))}
      </div>
      {toast && <div className="evsd-toast">{toast}</div>}
    </div>
  )
}

/* ── Agent ── */

interface ApprovalCard {
  title: string
  detail: string
  status: 'pending' | 'accepted' | 'rejected'
}

interface ChatItem {
  id: string
  role: 'user' | 'agent'
  text?: string
  tools?: Array<{ label: string; done: boolean }>
  approval?: ApprovalCard
}

interface MockSession {
  id: string
  title: string
  items: ChatItem[]
}

const TIERS = [
  { key: 'fast', label: '极速' },
  { key: 'primary', label: '标准' },
  { key: 'deep', label: '深度' },
] as const

const PERMISSIONS = [
  { key: 'ask', label: '写入前确认' },
  { key: 'auto', label: '自动接受编辑' },
] as const

function AgentView({ kind, fileName, say }: { kind: SideDockKind; fileName?: string; say: (msg: string) => void }) {
  const baseTitle = fileName ? fileName.replace(/\.[^.]+$/, '') : `当前${KIND_LABEL[kind]}`
  const [sessions, setSessions] = useState<MockSession[]>([{ id: 's1', title: baseTitle, items: [] }])
  const [activeId, setActiveId] = useState('s1')
  const [menuOpen, setMenuOpen] = useState(false)
  const [draft, setDraft] = useState('')
  const [sending, setSending] = useState(false)
  const [tier, setTier] = useState<(typeof TIERS)[number]['key']>('primary')
  const [permission, setPermission] = useState<(typeof PERMISSIONS)[number]['key']>('ask')
  const listRef = useRef<HTMLDivElement>(null)

  const active = sessions.find((s) => s.id === activeId) ?? sessions[0]

  useEffect(() => {
    const el = listRef.current
    if (el) el.scrollTop = el.scrollHeight
  }, [active.items.length, sending])

  const patchActive = (fn: (s: MockSession) => MockSession): void => {
    setSessions((prev) => prev.map((s) => (s.id === activeId ? fn(s) : s)))
  }

  const send = (): void => {
    const text = draft.trim()
    if (!text || sending) return
    setDraft('')
    patchActive((s) => ({ ...s, items: [...s.items, { id: nextId(), role: 'user', text }] }))
    setSending(true)
    window.setTimeout(() => {
      patchActive((s) => ({
        ...s,
        items: [
          ...s.items,
          {
            id: nextId(),
            role: 'agent',
            text: '已根据「材料」标签里的内容处理完毕，改动进入审阅：',
            approval: {
              title: `修改${KIND_LABEL[kind]}：${fileName ?? baseTitle}`,
              detail: `${text.slice(0, 24)}${text.length > 24 ? '…' : ''} · 等待确认`,
              status: 'pending',
            },
          },
        ],
      }))
      setSending(false)
    }, 900)
  }

  const resolveApproval = (itemId: string, status: 'accepted' | 'rejected'): void => {
    patchActive((s) => ({
      ...s,
      items: s.items.map((it) => (it.id === itemId && it.approval ? { ...it, approval: { ...it.approval, status } } : it)),
    }))
    say(status === 'accepted' ? '已接受 Agent 的修改（演示）' : '已拒绝 Agent 的修改（演示）')
  }

  const newSession = (): void => {
    const id = `s${Date.now()}`
    setSessions((prev) => [...prev, { id, title: '新会话', items: [] }])
    setActiveId(id)
    setMenuOpen(false)
  }

  const deleteSession = (id: string): void => {
    const remaining = sessions.filter((s) => s.id !== id)
    const nextList = remaining.length > 0 ? remaining : [{ id: `s${Date.now()}`, title: '新会话', items: [] }]
    setSessions(nextList)
    if (id === activeId) setActiveId(nextList[0].id)
    say('已删除会话')
  }

  return (
    <div className="evsd-agent">
      <div className="evsd-agent-toolbar">
        <button type="button" className="evsd-switcher" onClick={() => setMenuOpen((v) => !v)}>
          <span className="evsd-switcher-dot" />
          <span className="evsd-switcher-title">{active.title}</span>
          {CHEVRON_ICON}
        </button>
        <button type="button" className="evsd-icon-btn" title="新会话" onClick={newSession}>
          {PLUS_ICON}
        </button>
        {menuOpen && (
          <div className="evsd-menu">
            {sessions.map((s) => (
              <div key={s.id} className={`evsd-menu-item${s.id === activeId ? ' active' : ''}`}>
                <button
                  type="button"
                  onClick={() => {
                    setActiveId(s.id)
                    setMenuOpen(false)
                  }}
                >
                  {s.title}
                </button>
                <button type="button" className="evsd-menu-del" title="删除会话" onClick={() => deleteSession(s.id)}>
                  {X_ICON}
                </button>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="evsd-chat" ref={listRef}>
        {active.items.length === 0 && (
          <div className="evsd-empty">
            <div className="evsd-empty-title">和 Agent 一起处理这份{KIND_LABEL[kind]}</div>
            <div className="evsd-empty-hint">它能看到「材料」里的内容，写入前会先请你审阅</div>
            <div className="evsd-prompts">
              {STARTER_PROMPTS[kind].map((p) => (
                <button key={p} type="button" onClick={() => setDraft(p)}>
                  {p}
                </button>
              ))}
            </div>
          </div>
        )}
        {active.items.map((item) =>
          item.role === 'user' ? (
            <div key={item.id} className="evsd-msg user">{item.text}</div>
          ) : (
            <div key={item.id} className="evsd-msg agent">
              <div className="evsd-msg-body">
                {item.text && <p>{item.text}</p>}
                {item.tools && (
                  <ul className="evsd-tools">
                    {item.tools.map((t) => (
                      <li key={t.label} className={t.done ? 'done' : ''}>
                        <span className="evsd-tool-dot" />
                        {t.label}
                      </li>
                    ))}
                  </ul>
                )}
                {item.approval && (
                  <div className={`evsd-approval${item.approval.status !== 'pending' ? ` ${item.approval.status}` : ''}`}>
                    <div className="evsd-approval-title">{item.approval.title}</div>
                    <div className="evsd-approval-detail">{item.approval.detail}</div>
                    {item.approval.status === 'pending' ? (
                      <div className="evsd-approval-actions">
                        <button type="button" className="accept" onClick={() => resolveApproval(item.id, 'accepted')}>接受</button>
                        <button type="button" className="reject" onClick={() => resolveApproval(item.id, 'rejected')}>拒绝</button>
                      </div>
                    ) : (
                      <div className="evsd-approval-state">{item.approval.status === 'accepted' ? '✓ 已接受' : '✕ 已拒绝'}</div>
                    )}
                  </div>
                )}
              </div>
            </div>
          ),
        )}
        {sending && (
          <div className="evsd-msg agent">
            <div className="evsd-msg-body">
              <span className="evsd-typing"><span /><span /><span /></span>
            </div>
          </div>
        )}
      </div>

      <div className="evsd-composer">
        <textarea
          placeholder={`让 Agent 处理这份${KIND_LABEL[kind]}…（Enter 发送）`}
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              send()
            }
          }}
        />
        <div className="evsd-composer-bar">
          <div className="evsd-selects">
            <label className="evsd-select">
              <select value={tier} onChange={(e) => setTier(e.target.value as typeof tier)}>
                {TIERS.map((t) => <option key={t.key} value={t.key}>{`模型 · ${t.label}`}</option>)}
              </select>
            </label>
            <label className="evsd-select">
              <select value={permission} onChange={(e) => setPermission(e.target.value as typeof permission)}>
                {PERMISSIONS.map((p) => <option key={p.key} value={p.key}>{p.label}</option>)}
              </select>
            </label>
          </div>
          <button type="button" className="evsd-send" disabled={!draft.trim() || sending} onClick={send} title="发送">
            {SEND_ICON}
          </button>
        </div>
      </div>
    </div>
  )
}

/* ── 材料 ── */

function formatSize(bytes: number): string {
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  if (bytes >= 1024) return `${Math.round(bytes / 1024)} KB`
  return `${bytes} B`
}

function MaterialsView({ kind, say }: { kind: SideDockKind; say: (msg: string) => void }) {
  const [items, setItems] = useState<MaterialItem[]>(() => MOCK_MATERIALS[kind])
  const fileRef = useRef<HTMLInputElement>(null)

  const onFiles = (files: FileList | null): void => {
    if (!files || files.length === 0) return
    const added: MaterialItem[] = Array.from(files).map((f, i) => ({
      id: `local-${Date.now()}-${i}`,
      kind: 'file',
      title: f.name,
      excerpt: `本地文件 · ${formatSize(f.size)}`,
      meta: '本地文件',
      addedLabel: '刚刚添加',
    }))
    setItems((prev) => [...prev, ...added])
    say(`已添加 ${added.length} 份材料，Agent 可直接引用`)
  }

  return (
    <div className="evsd-materials">
      <div className="evsd-materials-list">
        <div className="evsd-hint">Agent 处理这份{KIND_LABEL[kind]}时可以引用的材料</div>
        {items.length === 0 && <div className="evsd-blank">还没有材料，从底部添加一份试试</div>}
        {items.map((m) => (
          <div key={m.id} className="evsd-material">
            <span className="evsd-material-icon">{MATERIAL_ICONS[m.kind]}</span>
            <div className="evsd-material-body">
              <div className="evsd-material-title">{m.title}</div>
              <div className="evsd-material-excerpt">「{m.excerpt}」</div>
              <div className="evsd-material-meta">{m.meta} · {m.addedLabel}</div>
            </div>
            <button type="button" className="evsd-material-del" title="移除材料" onClick={() => {
              setItems((prev) => prev.filter((x) => x.id !== m.id))
              say(`已移除材料「${m.title}」`)
            }}>
              {X_ICON}
            </button>
          </div>
        ))}
      </div>
      <div className="evsd-materials-foot">
        <input ref={fileRef} type="file" multiple style={{ display: 'none' }} onChange={(e) => { onFiles(e.target.files); e.target.value = '' }} />
        <button type="button" className="evsd-materials-add" onClick={() => fileRef.current?.click()}>
          {PLUS_ICON} 添加材料（本地文件）
        </button>
      </div>
    </div>
  )
}

/* ── 历史记录 ── */

type HistFilter = 'all' | 'user' | 'agent'

function HistoryView({ kind, say }: { kind: SideDockKind; say: (msg: string) => void }) {
  const [filter, setFilter] = useState<HistFilter>('all')
  const entries = MOCK_HISTORY[kind].filter((e) => filter === 'all' || e.actor === filter)

  const restore = (entry: HistoryEntry): void => {
    const ok = window.confirm(`恢复到此版本？\n\n${entry.title}${entry.detail ? `\n${entry.detail}` : ''}\n\n恢复后会生成一个新版本，当前内容不会丢失。`)
    if (ok) say(`已恢复到「${entry.title}」（演示）`)
  }

  return (
    <div className="evsd-hist">
      <div className="evsd-hist-filters">
        {[{ key: 'all' as const, label: '全部' }, { key: 'user' as const, label: '我的修改' }, { key: 'agent' as const, label: 'Agent' }].map((f) => (
          <button key={f.key} type="button" className={filter === f.key ? 'active' : ''} onClick={() => setFilter(f.key)}>
            {f.label}
          </button>
        ))}
      </div>
      <div className="evsd-hist-list">
        {entries.length === 0 && <div className="evsd-blank">该筛选下没有记录</div>}
        {entries.map((e) => (
          <div key={e.id} className="evsd-hist-entry">
            <span className={`evsd-hist-dot ${e.actor}`} />
            <div className="evsd-hist-body">
              <div className="evsd-hist-head">
                {e.actor === 'agent' ? (
                  <span className="evsd-actor agent">Agent</span>
                ) : e.actor === 'user' ? (
                  <span className="evsd-actor user">{USER_ICON} 王小雨</span>
                ) : (
                  <span className="evsd-actor system">系统</span>
                )}
                <span className="evsd-hist-title">{e.title}</span>
              </div>
              {e.detail && <div className="evsd-hist-detail">{e.detail}</div>}
              {(e.restorable || e.diffable) && (
                <div className="evsd-hist-actions">
                  {e.diffable && (
                    <button type="button" onClick={() => say(`「${e.title}」的改动预览将在下一阶段接入`)}>
                      {DIFF_ICON} 查看改动
                    </button>
                  )}
                  {e.restorable && (
                    <button type="button" className="restore" onClick={() => restore(e)}>
                      {RESTORE_ICON} 恢复此版本
                    </button>
                  )}
                </div>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}
