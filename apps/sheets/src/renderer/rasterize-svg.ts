/**
 * SVG → PNG 光栅化（Chromium 自带解码，无额外依赖）。
 *
 * Excel 的 SVG 图片不是"一个 media 文件"就够：OOXML 里要同时写一张位图回退
 * （`<a:blip r:embed=PNG>`）和矢量本体（`<asvg:svgBlip>`）。老版本 Excel、部分
 * 预览器、PDF 导出只认位图回退，缺了它图片就是空白。所以插入 SVG 时先在渲染进程
 * 把它画成 PNG，作为回退一起写进文件。
 */

export interface RasterizedSvg {
  readonly pngDataUrl: string
  readonly width: number
  readonly height: number
}

/// SVG 只给了 viewBox 时 Chromium 会把自然尺寸当成 300×150，这里按 viewBox 还原比例。
function viewBoxSize(svgText: string): { width: number; height: number } | null {
  const box = /viewBox\s*=\s*["']([^"']+)["']/i.exec(svgText)
  if (!box?.[1]) return null
  const parts = box[1].trim().split(/[\s,]+/).map(Number)
  if (parts.length !== 4) return null
  const width = parts[2] ?? 0
  const height = parts[3] ?? 0
  if (!(width > 0) || !(height > 0)) return null
  return { width, height }
}

function base64Payload(dataUrl: string): string | null {
  const comma = dataUrl.indexOf(',')
  if (comma < 0) return null
  return dataUrl.slice(comma + 1)
}

export async function rasterizeSvgDataUrl(
  svgDataUrl: string,
  options?: { readonly maxEdge?: number },
): Promise<RasterizedSvg> {
  const maxEdge = options?.maxEdge ?? 2048
  const image = new Image()
  const loaded = new Promise<void>((resolve, reject) => {
    image.onload = () => resolve()
    image.onerror = () => reject(new Error('无法解码该 SVG 图片。'))
  })
  image.src = svgDataUrl
  await loaded

  let naturalWidth = image.naturalWidth || image.width || 0
  let naturalHeight = image.naturalHeight || image.height || 0
  if (!(naturalWidth > 0) || !(naturalHeight > 0)) {
    // 只有 viewBox 的 SVG：从源文本里还原宽高比，按 1024 宽出图。
    const payload = base64Payload(svgDataUrl)
    let viewBox: { width: number; height: number } | null = null
    if (payload) {
      try {
        viewBox = viewBoxSize(atob(payload))
      } catch {
        viewBox = null
      }
    }
    naturalWidth = viewBox?.width ?? 300
    naturalHeight = viewBox?.height ?? 150
  }

  const scale = Math.min(1, maxEdge / Math.max(naturalWidth, naturalHeight))
  const width = Math.max(1, Math.round(naturalWidth * scale))
  const height = Math.max(1, Math.round(naturalHeight * scale))
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('画布不可用，无法把 SVG 转成位图回退。')
  context.drawImage(image, 0, 0, width, height)
  return { pngDataUrl: canvas.toDataURL('image/png'), width, height }
}
