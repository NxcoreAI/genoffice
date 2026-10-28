import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'

import './side-dock.css'

/**
 * EverRoom host side dock (Agent / Materials / History), embedded into each
 * GenOffice editor when the host opens the view with the `dock=1` query
 * parameter. Self-contained by design: no app dependencies, inline icons,
 * mock data included — the host wires the real Agent bridge in later stages.
 *
 * `docksrc=<url>`: the host may supply the URL of its own Agent chat page
 * (EverRoom ported chat); the Agent tab then renders it in a filling iframe
 * instead of the built-in mock chat, keeping the dock geometry unchanged.
 */

export type SideDockKind = 'docs' | 'sheets' | 'slides' | 'pdf'

const DOCK_AGENT_SRC: string | null = (() => {
  try {
    const raw = new URLSearchParams(window.location.search).get('docksrc')
    return raw && raw.length > 0 ? raw : null
  } catch {
    return null
  }
})()

/* ───────────── mock 数据（按编辑器类型） ───────────── */

/** 材料库 = 用户挂给当前文档、供 Agent 引用的本地文件。 */
type MaterialKind = 'docx' | 'xlsx' | 'pptx' | 'pdf' | 'md' | 'image' | 'txt' | 'other'

interface MaterialItem {
  id: string
  kind: MaterialKind
  /** 文件名（含扩展名）。 */
  title: string
  /** 展示用大小，如「86 KB」。 */
  sizeLabel: string
  addedLabel: string
}

const MATERIAL_LABEL: Record<MaterialKind, string> = {
  docx: '文档',
  xlsx: '表格',
  pptx: '演示',
  pdf: 'PDF',
  md: '笔记',
  image: '图片',
  txt: '文本',
  other: '文件',
}

/** 本地添加的文件只有名字可用，按扩展名推导类型。 */
const kindFromName = (name: string): MaterialKind => {
  const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase()
  if (ext === 'docx' || ext === 'doc') return 'docx'
  if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') return 'xlsx'
  if (ext === 'pptx' || ext === 'ppt') return 'pptx'
  if (ext === 'pdf') return 'pdf'
  if (ext === 'md' || ext === 'markdown') return 'md'
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'heic'].includes(ext)) return 'image'
  if (ext === 'txt' || ext === 'text' || ext === 'log') return 'txt'
  return 'other'
}

const MOCK_MATERIALS: Record<SideDockKind, MaterialItem[]> = {
  docs: [
    { id: 'm1', kind: 'docx', title: '上周周报-1021.docx', sizeLabel: '86 KB', addedLabel: '今天添加' },
    { id: 'm2', kind: 'xlsx', title: '连接器排期总表.xlsx', sizeLabel: '24 KB', addedLabel: '今天添加' },
    { id: 'm3', kind: 'image', title: '视觉定稿-终版.png', sizeLabel: '2.3 MB', addedLabel: '昨天' },
    { id: 'm4', kind: 'pdf', title: '动效规范-V1.pdf', sizeLabel: '1.2 MB', addedLabel: '3 天前' },
  ],
  sheets: [
    { id: 'm1', kind: 'pdf', title: '供应商报价Q4.pdf', sizeLabel: '1.2 MB', addedLabel: '今天添加' },
    { id: 'm2', kind: 'xlsx', title: '报价明细-三家对比.csv', sizeLabel: '18 KB', addedLabel: '今天添加' },
    { id: 'm3', kind: 'txt', title: '预算上限说明.txt', sizeLabel: '3 KB', addedLabel: '上周添加' },
  ],
  slides: [
    { id: 'm1', kind: 'pdf', title: 'V1视觉规范.pdf', sizeLabel: '2.8 MB', addedLabel: '上周添加' },
    { id: 'm2', kind: 'docx', title: '发布评审0926纪要.docx', sizeLabel: '64 KB', addedLabel: '上周添加' },
    { id: 'm3', kind: 'pptx', title: '时间线页-旧版.pptx', sizeLabel: '1.6 MB', addedLabel: '上周添加' },
  ],
  pdf: [
    { id: 'm1', kind: 'md', title: '竞品跟踪-1024.md', sizeLabel: '12 KB', addedLabel: '上周添加' },
    { id: 'm2', kind: 'image', title: 'Notion-AI-定价页.png', sizeLabel: '640 KB', addedLabel: '上周添加' },
  ],
}

interface HistoryEntry {
  id: string
  actor: 'user' | 'agent' | 'system'
  title: string
  detail?: string
  restorable?: boolean
  diffable?: boolean
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

/* 材料库文件类型 squircle 图标（自包含，配色与宿主文件岛的类型图标一致）。 */
const MATERIAL_THEMES: Record<MaterialKind, { from: string; to: string; letters: string }> = {
  docx: { from: '#4aa3ff', to: '#0a6fe8', letters: 'W' },
  xlsx: { from: '#4cc478', to: '#1f9450', letters: 'S' },
  pptx: { from: '#ffb340', to: '#ef7d0e', letters: 'P' },
  pdf: { from: '#ff6f61', to: '#e0362a', letters: 'PDF' },
  md: { from: '#b9bbc2', to: '#83868f', letters: 'MD' },
  image: { from: '#b78aff', to: '#7d3fe0', letters: 'IMG' },
  txt: { from: '#4ec3d4', to: '#0f92a3', letters: 'TXT' },
  other: { from: '#b9bbc2', to: '#83868f', letters: 'FILE' },
}

function FileGlyph({ kind }: { kind: MaterialKind }) {
  const th = MATERIAL_THEMES[kind]
  const letters = th.letters
  const fs = letters.length > 1 ? 12.5 : 20
  return (
    <svg width="28" height="28" viewBox="0 0 48 48" aria-hidden="true" style={{ flex: 'none' }}>
      <defs>
        <linearGradient id={`evsd-ftg-${kind}`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor={th.from} />
          <stop offset="1" stopColor={th.to} />
        </linearGradient>
        <linearGradient id={`evsd-ftg-${kind}-g`} x1="0" y1="0" x2="0" y2="1">
          <stop offset="0" stopColor="#fff" stopOpacity="0.32" />
          <stop offset="1" stopColor="#fff" stopOpacity="0" />
        </linearGradient>
      </defs>
      <rect x="1.5" y="1.5" width="45" height="45" rx="11.5" fill={`url(#evsd-ftg-${kind})`} />
      <path d="M1.5 13 A11.5 11.5 0 0 1 13 1.5 H35 A11.5 11.5 0 0 1 46.5 13 V21 H1.5 Z" fill={`url(#evsd-ftg-${kind}-g)`} />
      <text
        x="24" y="25.5" textAnchor="middle" dominantBaseline="central"
        fill="#fff" fontWeight="800" fontSize={fs}
        letterSpacing={letters.length > 1 ? 0.8 : 0}
        fontFamily="-apple-system, 'PingFang SC', 'Microsoft YaHei', sans-serif"
      >
        {letters}
      </text>
    </svg>
  )
}

/* ───────────── 组件 ───────────── */

type DockTab = 'agent' | 'materials' | 'history'

const TABS: Array<{ key: DockTab; label: string; icon: ReactNode }> = [
  { key: 'agent', label: 'Agent', icon: AGENT_ICON },
  { key: 'materials', label: '材料库', icon: MATERIALS_ICON },
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
            {tab === 'agent' &&
              (DOCK_AGENT_SRC ? (
                <AgentFrame kind={kind} fileName={fileName} />
              ) : (
                <AgentView kind={kind} fileName={fileName} say={say} />
              ))}
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

/* ── Agent（宿主聊天页装载） ── */

function AgentFrame({ kind, fileName }: { kind: SideDockKind; fileName?: string }) {
  const src = useMemo(() => {
    const url = new URL(DOCK_AGENT_SRC!, window.location.href)
    url.searchParams.set('kind', kind)
    if (fileName) url.searchParams.set('file', fileName)
    return url.toString()
  }, [kind, fileName])
  return (
    <div className="evsd-agent-frame">
      <iframe src={src} title="Agent" />
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
            <div className="evsd-empty-hint">它能看到「材料库」里的文件，写入前会先请你审阅</div>
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

/* ── 材料库（本地文件架） ── */

function formatSize(bytes: number): string {
  if (bytes >= 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024 / 1024).toFixed(1)} GB`
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`
  if (bytes >= 1024) return `${Math.round(bytes / 1024)} KB`
  return `${bytes} B`
}

const SIZE_RE = /^([\d.]+)\s*(B|KB|MB|GB)$/
const SIZE_UNITS: Record<string, number> = { B: 1, KB: 1024, MB: 1024 ** 2, GB: 1024 ** 3 }

/** mock 只有展示字符串，统计总量时反解析 sizeLabel。 */
function parseSize(label: string): number {
  const m = SIZE_RE.exec(label.trim())
  return m ? parseFloat(m[1]) * SIZE_UNITS[m[2]] : 0
}

function MaterialsView({ kind, say }: { kind: SideDockKind; say: (msg: string) => void }) {
  const [items, setItems] = useState<MaterialItem[]>(() => MOCK_MATERIALS[kind])
  const [dragOver, setDragOver] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

  const addFiles = (files: Array<{ name: string; size: number }>): void => {
    if (files.length === 0) return
    const added: MaterialItem[] = files.map((f, i) => ({
      id: `local-${Date.now()}-${i}`,
      kind: kindFromName(f.name),
      title: f.name,
      sizeLabel: formatSize(f.size),
      addedLabel: '刚刚添加',
    }))
    setItems((prev) => [...added, ...prev])
    say(`已添加 ${added.length} 个文件，Agent 可直接引用`)
  }

  const total = formatSize(items.reduce((sum, m) => sum + parseSize(m.sizeLabel), 0))

  return (
    <div
      className="evsd-materials"
      onDragOver={(e) => {
        if (!Array.from(e.dataTransfer.types).includes('Files')) return
        e.preventDefault()
        setDragOver(true)
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node)) setDragOver(false)
      }}
      onDrop={(e) => {
        e.preventDefault()
        setDragOver(false)
        addFiles(Array.from(e.dataTransfer.files).map((f) => ({ name: f.name, size: f.size })))
      }}
    >
      <div className="evsd-materials-list">
        {items.length > 0 && (
          <div className="evsd-hint">{items.length} 个文件 · {total} · 供 Agent 引用</div>
        )}
        {items.length === 0 && (
          <div className="evsd-materials-empty">
            <div className="evsd-materials-empty-title">拖入文件，或点下方按钮添加</div>
            <div className="evsd-materials-empty-sub">添加后 Agent 处理这份{KIND_LABEL[kind]}时可以直接引用</div>
          </div>
        )}
        {items.map((m) => (
          <div
            key={m.id}
            className="evsd-material"
            role="button"
            tabIndex={0}
            title={m.title}
            onClick={() => say(`「${m.title}」的预览将在下一阶段接入`)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                say(`「${m.title}」的预览将在下一阶段接入`)
              }
            }}
          >
            <FileGlyph kind={m.kind} />
            <div className="evsd-material-body">
              <div className="evsd-material-title">{m.title}</div>
              <div className="evsd-material-meta">{MATERIAL_LABEL[m.kind]} · {m.sizeLabel} · {m.addedLabel}</div>
            </div>
            <button
              type="button"
              className="evsd-material-del"
              title="移除文件"
              onClick={(e) => {
                e.stopPropagation()
                setItems((prev) => prev.filter((x) => x.id !== m.id))
                say(`已移除「${m.title}」`)
              }}
            >
              {X_ICON}
            </button>
          </div>
        ))}
      </div>
      {dragOver && <div className="evsd-materials-dropzone">松开添加到材料库</div>}
      <div className="evsd-materials-foot">
        <input
          ref={fileRef}
          type="file"
          multiple
          style={{ display: 'none' }}
          onChange={(e) => {
            addFiles(Array.from(e.target.files ?? []).map((f) => ({ name: f.name, size: f.size })))
            e.target.value = ''
          }}
        />
        <button type="button" className="evsd-materials-add" onClick={() => fileRef.current?.click()}>
          {PLUS_ICON} 添加文件
        </button>
        <div className="evsd-materials-foot-hint">也可把文件直接拖进面板</div>
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
