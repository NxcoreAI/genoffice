import SvgPath from 'svgpath'

export const MAX_PATH_D_CHARS = 2000
const MAX_SEGMENTS = 400
/** 归一目标边长，与烘焙图标同空间（24×9000）：输入坐标任意，输出坐标有界 */
const TARGET = 216000

interface Pt {
  x: number
  y: number
}

interface PathCmd {
  kind: 'M' | 'L' | 'C' | 'Q' | 'Z'
  pts: Pt[]
}

/**
 * SVG path d → OOXML `<a:pathLst>`。相对命令/H/V/S/T/A 由 svgpath 归一为
 * 绝对 M/L/C/Q/Z；包围盒平移到原点后按最长边缩放到 TARGET 空间（拉伸语义
 * 由 PowerPoint 完成：path 空间被非等比映射到元素框）。非法或超限返回 null。
 */
export function svgDToPathLst(d: string, fillRule: 'nonzero' | 'evenodd'): string | null {
  if (typeof d !== 'string' || !d.trim() || d.length > MAX_PATH_D_CHARS) return null
  const sp = new SvgPath(d).abs().unshort().unarc()

  const cmds: PathCmd[] = []
  let cx = 0
  let cy = 0
  try {
    sp.iterate((seg) => {
      const kind = String(seg[0])
      const at = (i: number): Pt => ({ x: Number(seg[i]), y: Number(seg[i + 1]) })
      if (kind === 'M' || kind === 'L' || kind === 'C' || kind === 'Q') {
        const pts: Pt[] = []
        for (let i = 1; i + 1 < seg.length; i += 2) pts.push(at(i))
        if (pts.some((p) => !Number.isFinite(p.x) || !Number.isFinite(p.y))) return
        cmds.push({ kind, pts })
        const last = pts[pts.length - 1]!
        cx = last.x
        cy = last.y
      } else if (kind === 'H') {
        const x = Number(seg[1])
        if (Number.isFinite(x)) {
          cmds.push({ kind: 'L', pts: [{ x, y: cy }] })
          cx = x
        }
      } else if (kind === 'V') {
        const y = Number(seg[1])
        if (Number.isFinite(y)) {
          cmds.push({ kind: 'L', pts: [{ x: cx, y }] })
          cy = y
        }
      } else if (kind === 'Z' || kind === 'z') {
        cmds.push({ kind: 'Z', pts: [] })
      }
    })
  } catch {
    return null
  }
  if (cmds.length === 0 || cmds.length > MAX_SEGMENTS || !cmds.some((c) => c.kind === 'M')) {
    return null
  }

  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const c of cmds) {
    for (const p of c.pts) {
      if (p.x < minX) minX = p.x
      if (p.x > maxX) maxX = p.x
      if (p.y < minY) minY = p.y
      if (p.y > maxY) maxY = p.y
    }
  }
  const bw = Math.max(maxX - minX, 1e-6)
  const bh = Math.max(maxY - minY, 1e-6)
  const s = TARGET / Math.max(bw, bh)
  const px = (x: number) => Math.round((x - minX) * s)
  const py = (y: number) => Math.round((y - minY) * s)

  let xml = ''
  for (const c of cmds) {
    const ptXml = (p: Pt) => `<a:pt x="${px(p.x)}" y="${py(p.y)}"/>`
    if (c.kind === 'M') xml += `<a:moveTo>${ptXml(c.pts[0]!)}</a:moveTo>`
    else if (c.kind === 'L') xml += `<a:lnTo>${ptXml(c.pts[0]!)}</a:lnTo>`
    else if (c.kind === 'C')
      xml += `<a:cubicBezTo>${ptXml(c.pts[0]!)}${ptXml(c.pts[1]!)}${ptXml(c.pts[2]!)}</a:cubicBezTo>`
    else if (c.kind === 'Q') xml += `<a:quadBezTo>${ptXml(c.pts[0]!)}${ptXml(c.pts[1]!)}</a:quadBezTo>`
    else xml += '<a:close/>'
  }
  const w = Math.max(1, Math.round(bw * s))
  const h = Math.max(1, Math.round(bh * s))
  const fill = fillRule === 'evenodd' ? 'evenOdd' : 'norm'
  return `<a:pathLst><a:path w="${w}" h="${h}" fill="${fill}">${xml}</a:path></a:pathLst>`
}
