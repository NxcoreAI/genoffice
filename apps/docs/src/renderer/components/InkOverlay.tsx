import { useCallback, useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'
import {
  HIGHLIGHTER_ALPHA,
  eraseStrokeAt,
  nextInkId,
  strokePathD,
  type InkAnnotation,
  type InkPoint,
  type InkStroke,
  type InkTool,
} from '../editor/ink'

/**
 * Transparent drawing layer covering the page (pen/highlighter/eraser). Lives inside
 * .page-wrap, which is inside the zoomed .doc-zoom container — pointer
 * coordinates are divided by the zoom factor so stroke geometry is stored at
 * 100% zoom, matching the engine's EMU offsets on save.
 *
 * Finished strokes are anchored to the nearest body paragraph (data-idx) and
 * stored relative to its top-left corner, so they follow the paragraph when
 * content above reflows.
 */

interface InkOverlayProps {
  tool: InkTool
  /** hex without '#' */
  color: string
  width: number
  zoom: number
  annotations: InkAnnotation[]
  onAdd: (annotation: InkAnnotation) => void
  /** Replace the whole annotation list: the eraser splits strokes into fragments. */
  onReplace: (annotations: InkAnnotation[]) => void
}

/** overlay-relative top-left of every anchored paragraph, at 100% zoom */
type AnchorPositions = Map<number, InkPoint>

const ERASER_RADIUS = 6

function measureAnchors(overlay: HTMLElement, zoomFactor: number): AnchorPositions {
  const positions: AnchorPositions = new Map()
  const base = overlay.getBoundingClientRect()
  const blocks = overlay.parentElement?.querySelectorAll<HTMLElement>('.ProseMirror > [data-idx]')
  for (const el of blocks ? Array.from(blocks) : []) {
    const idx = parseInt(el.dataset.idx ?? '', 10)
    if (!Number.isFinite(idx)) continue
    const rect = el.getBoundingClientRect()
    positions.set(idx, {
      x: (rect.left - base.left) / zoomFactor,
      y: (rect.top - base.top) / zoomFactor,
    })
  }
  return positions
}

/** anchor for a stroke: the paragraph whose top is closest above its first point */
function pickAnchor(positions: AnchorPositions, start: InkPoint): number | null {
  let best: number | null = null
  let bestTop = -Infinity
  for (const [idx, pos] of positions) {
    if (pos.y <= start.y && pos.y > bestTop) {
      bestTop = pos.y
      best = idx
    }
  }
  if (best !== null) return best
  // stroke above every paragraph: fall back to the topmost one
  let top = Infinity
  for (const [idx, pos] of positions) {
    if (pos.y < top) {
      top = pos.y
      best = idx
    }
  }
  return best
}

export function InkOverlay({ tool, color, width, zoom, annotations, onAdd, onReplace }: InkOverlayProps) {
  const overlayRef = useRef<HTMLDivElement>(null)
  const [anchors, setAnchors] = useState<AnchorPositions>(new Map())
  const [livePoints, setLivePoints] = useState<InkPoint[] | null>(null)
  const drawingRef = useRef<InkPoint[] | null>(null)
  /** eraser drag working copy: mutated locally so one gesture stays consistent */
  const erasingRef = useRef<InkAnnotation[] | null>(null)

  const zoomFactor = zoom / 100
  const drawing = tool === 'pen' || tool === 'highlighter'

  const remeasure = useCallback(() => {
    const overlay = overlayRef.current
    if (overlay) setAnchors(measureAnchors(overlay, zoomFactor))
  }, [zoomFactor])

  // paragraph positions shift on every edit/zoom/resize; observe the page and
  // re-measure so anchored strokes stay glued to their paragraphs
  useEffect(() => {
    remeasure()
    const page = overlayRef.current?.parentElement
    if (!page) return
    const observer = new ResizeObserver(remeasure)
    observer.observe(page)
    const mutations = new MutationObserver(remeasure)
    mutations.observe(page, { childList: true, subtree: true, characterData: true })
    return () => {
      observer.disconnect()
      mutations.disconnect()
    }
  }, [remeasure, annotations])

  const localPoint = (e: ReactPointerEvent): InkPoint => {
    const rect = overlayRef.current!.getBoundingClientRect()
    return { x: (e.clientX - rect.left) / zoomFactor, y: (e.clientY - rect.top) / zoomFactor }
  }

  const eraseAt = useCallback(
    (point: InkPoint) => {
      const working = erasingRef.current
      if (!working) return
      let changed = false
      const next: InkAnnotation[] = []
      for (const ann of working) {
        const pos = anchors.get(ann.anchorIndex)
        if (!pos) {
          next.push(ann)
          continue
        }
        const rel = { x: point.x - pos.x, y: point.y - pos.y }
        if (ann.stroke) {
          // 局部擦除：只擦掉笔尖覆盖的那一段，穿过笔迹时把它拆成两段，
          // 而不是整条笔迹一起消失。
          const fragments = eraseStrokeAt(ann.stroke, rel, ERASER_RADIUS)
          if (!fragments) {
            next.push(ann)
            continue
          }
          changed = true
          fragments.forEach((stroke, index) => {
            next.push({ ...ann, id: index === 0 ? ann.id : nextInkId(), stroke })
          })
          continue
        }
        if (ann.image) {
          const { offsetXPx: x, offsetYPx: y, widthPx: w, heightPx: h } = ann.image
          if (rel.x >= x && rel.x <= x + w && rel.y >= y && rel.y <= y + h) {
            // 位图形态的旧标注没有矢量点可拆：整体删除（保持原行为）。
            changed = true
            continue
          }
        }
        next.push(ann)
      }
      if (!changed) return
      erasingRef.current = next
      onReplace(next)
    },
    [anchors, onReplace],
  )

  const onPointerDown = (e: ReactPointerEvent) => {
    if (tool === 'select' || e.button !== 0) return
    e.preventDefault()
    ;(e.target as HTMLElement).setPointerCapture(e.pointerId)
    const point = localPoint(e)
    if (tool === 'eraser') {
      erasingRef.current = annotations
      drawingRef.current = [] // non-null marks "eraser drag active"
      eraseAt(point)
      return
    }
    drawingRef.current = [point]
    setLivePoints([point])
  }

  const onPointerMove = (e: ReactPointerEvent) => {
    if (!drawingRef.current) return
    const point = localPoint(e)
    if (tool === 'eraser') {
      eraseAt(point)
      return
    }
    const points = drawingRef.current
    const last = points[points.length - 1]
    if (Math.hypot(point.x - last.x, point.y - last.y) < 0.8) return
    points.push(point)
    setLivePoints([...points])
  }

  const onPointerUp = () => {
    const points = drawingRef.current
    drawingRef.current = null
    erasingRef.current = null
    if (tool !== 'pen' && tool !== 'highlighter') return
    setLivePoints(null)
    if (!points || points.length === 0) return
    const anchorIndex = pickAnchor(anchors, points[0])
    if (anchorIndex === null) return
    const pos = anchors.get(anchorIndex)!
    const stroke: InkStroke = {
      tool,
      color,
      width,
      points: points.map((p) => ({ x: p.x - pos.x, y: p.y - pos.y })),
    }
    onAdd({ id: nextInkId(), anchorIndex, stroke })
  }

  const strokeStyle = (s: InkStroke) => ({
    stroke: `#${s.color}`,
    strokeWidth: s.width,
    strokeOpacity: s.tool === 'highlighter' ? HIGHLIGHTER_ALPHA : 1,
    fill: 'none' as const,
    strokeLinecap: 'round' as const,
    strokeLinejoin: 'round' as const,
  })

  return (
    <div
      ref={overlayRef}
      className={`ink-overlay ${tool !== 'select' ? `ink-active ink-tool-${tool}` : ''}`}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerCancel={onPointerUp}
    >
      <svg className="ink-svg" aria-hidden>
        {annotations.map((ann) => {
          const pos = anchors.get(ann.anchorIndex)
          if (!pos) return null
          if (ann.stroke) {
            return (
              <path
                key={ann.id}
                d={strokePathD(ann.stroke.points)}
                transform={`translate(${pos.x} ${pos.y})`}
                {...strokeStyle(ann.stroke)}
              />
            )
          }
          if (ann.image?.dataUrl) {
            return (
              <image
                key={ann.id}
                href={ann.image.dataUrl}
                x={pos.x + ann.image.offsetXPx}
                y={pos.y + ann.image.offsetYPx}
                width={ann.image.widthPx}
                height={ann.image.heightPx}
              />
            )
          }
          return null
        })}
        {livePoints && drawing && (
          <path d={strokePathD(livePoints)} {...strokeStyle({ tool, color, width, points: livePoints })} />
        )}
      </svg>
    </div>
  )
}
