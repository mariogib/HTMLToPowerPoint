import type JSZip from 'jszip'
import {
  loadRelationships,
  loadSlideColorScheme,
  loadXml,
  resolveColorElement,
  type ThemeColorMap,
} from './pptxBackground'
import { DEFAULT_SLIDE_SIZE, emuToPx, type SlideSize } from './pptxSlideSize'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface SmartArtShape {
  leftPct: number
  topPct: number
  widthPct: number
  heightPct: number
  text: string
  fill: string
  textColor: string
  fontSizePx: number
  shape: 'ellipse' | 'roundRect' | 'rect' | 'other'
}

export interface SmartArtExport {
  id: string
  title: string
  nodes: string[]
  shapes: SmartArtShape[]
  imageDataUrl?: string
  leftPx: number
  topPx: number
  widthPx: number
  heightPx: number
}

const localName = (el: Element | null | undefined): string => {
  if (!el) return ''
  return el.localName || el.tagName.replace(/^.*:/, '')
}

const elementsByLocal = (root: ParentNode | null | undefined, name: string): Element[] => {
  if (!root) return []
  const scope = root instanceof Document ? root.documentElement : (root as Element)
  if (!scope?.getElementsByTagName) return []
  return Array.from(scope.getElementsByTagName('*')).filter((el) => localName(el) === name) as Element[]
}

const firstByLocal = (parent: ParentNode | null | undefined, name: string): Element | null =>
  elementsByLocal(parent, name)[0] || null

const getZipEntry = (zip: JSZip, path: string) => {
  const candidates = [path, path.replace(/^\//, ''), `/${path.replace(/^\//, '')}`]
  for (const candidate of candidates) {
    const entry = zip.file(candidate) || zip.files[candidate]
    if (entry && !entry.dir) return entry
  }
  const needle = path.replace(/^\//, '').toLowerCase()
  const match = Object.keys(zip.files).find((key) => key.replace(/^\//, '').toLowerCase() === needle)
  return match ? zip.files[match] : null
}

const readEmu = (el: Element | null, attr: string): number => {
  if (!el) return 0
  return parseInt(el.getAttribute(attr) || '0', 10) || 0
}

const readTransform = (
  container: Element | null,
  slideSize: SlideSize
): {
  leftPx: number
  topPx: number
  widthPx: number
  heightPx: number
  widthEmu: number
  heightEmu: number
} => {
  const xfrm =
    firstByLocal(container, 'xfrm') ||
    firstByLocal(firstByLocal(container, 'spPr'), 'xfrm') ||
    firstByLocal(firstByLocal(container, 'pic'), 'xfrm')

  const off = firstByLocal(xfrm, 'off')
  const ext = firstByLocal(xfrm, 'ext')
  const x = readEmu(off, 'x')
  const y = readEmu(off, 'y')
  const cx = readEmu(ext, 'cx') || slideSize.widthEmu * 0.5
  const cy = readEmu(ext, 'cy') || slideSize.heightEmu * 0.35

  return {
    leftPx: Math.max(0, emuToPx(x)),
    topPx: Math.max(0, emuToPx(y)),
    widthPx: Math.max(120, emuToPx(cx)),
    heightPx: Math.max(80, emuToPx(cy)),
    widthEmu: Math.max(1, cx),
    heightEmu: Math.max(1, cy),
  }
}

const getAttrNS = (el: Element, local: string): string | null =>
  el.getAttributeNS(R_NS, local) ||
  el.getAttribute(`r:${local}`) ||
  el.getAttribute(local)

const getEmbedId = (el: Element | null): string | null => {
  if (!el) return null
  return getAttrNS(el, 'embed') || getAttrNS(el, 'link')
}

const isImagePath = (path: string): boolean => {
  const lower = path.toLowerCase()
  return (
    lower.endsWith('.png') ||
    lower.endsWith('.jpg') ||
    lower.endsWith('.jpeg') ||
    lower.endsWith('.gif') ||
    lower.endsWith('.webp') ||
    lower.endsWith('.emf') ||
    lower.endsWith('.wmf')
  )
}

const mimeForImage = (path: string): string => {
  const ext = path.split('.').pop()?.toLowerCase()
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg'
  if (ext === 'gif') return 'image/gif'
  if (ext === 'webp') return 'image/webp'
  return 'image/png'
}

const toImageDataUrl = async (zip: JSZip, mediaPath: string): Promise<string | null> => {
  const entry = getZipEntry(zip, mediaPath)
  if (!entry) return null
  try {
    const base64 = await entry.async('base64')
    return `data:${mimeForImage(mediaPath)};base64,${base64}`
  } catch (error) {
    console.error('Failed reading SmartArt image:', mediaPath, error)
    return null
  }
}

const escapeHtml = (text: string): string =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

const isDiagramGraphicData = (graphicData: Element): boolean => {
  const uri = (graphicData.getAttribute('uri') || '').toLowerCase()
  if (uri.includes('diagram')) return true
  return elementsByLocal(graphicData, 'relIds').length > 0
}

const findSmartArtFrames = (slideDoc: Document): Element[] => {
  const frames: Element[] = []
  for (const frame of elementsByLocal(slideDoc, 'graphicFrame')) {
    const graphicData = firstByLocal(frame, 'graphicData')
    if (graphicData && isDiagramGraphicData(graphicData)) {
      frames.push(frame)
      continue
    }
    if (elementsByLocal(frame, 'relIds').length > 0) {
      frames.push(frame)
    }
  }
  return frames
}

const extractNodeTexts = (dataDoc: Document): string[] => {
  const texts: string[] = []
  for (const pt of elementsByLocal(dataDoc, 'pt')) {
    const type = (pt.getAttribute('type') || 'node').toLowerCase()
    if (type === 'doc' || type === 'pres' || type === 'partrans' || type === 'sibtrans') {
      continue
    }

    const textContainer = firstByLocal(pt, 't')
    if (!textContainer) continue

    const chunks = elementsByLocal(textContainer, 't')
      .map((t) => (t.textContent || '').trim())
      .filter(Boolean)

    const joined =
      chunks.length > 0
        ? chunks.join(' ')
        : (textContainer.textContent || '').replace(/\s+/g, ' ').trim()

    if (joined) texts.push(joined)
  }

  const seen = new Set<string>()
  return texts.filter((t) => {
    const key = t.toLowerCase()
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

const findFallbackImagePath = (
  frame: Element,
  rels: Map<string, { type: string; target: string }>
): string | null => {
  let current: Element | null = frame
  while (current) {
    if (localName(current) === 'AlternateContent') {
      for (const fallback of elementsByLocal(current, 'Fallback')) {
        for (const blip of elementsByLocal(fallback, 'blip')) {
          const embed = getEmbedId(blip)
          if (embed && rels.has(embed) && isImagePath(rels.get(embed)!.target)) {
            return rels.get(embed)!.target
          }
        }
      }
    }
    current = current.parentElement
  }
  return null
}

const resolveDiagramDataPath = (
  frame: Element,
  rels: Map<string, { type: string; target: string }>
): string | null => {
  const relIds = firstByLocal(frame, 'relIds')
  if (relIds) {
    const dm = getAttrNS(relIds, 'dm')
    if (dm && rels.has(dm)) return rels.get(dm)!.target
  }
  return null
}

/**
 * Drawing part is referenced from diagram data via extLst/dataModelExt@relId,
 * resolved against the data part's own .rels (not the slide rels).
 */
const resolveDiagramDrawingPath = async (
  zip: JSZip,
  dataPath: string,
  dataDoc: Document
): Promise<string | null> => {
  const dataRels = await loadRelationships(zip, dataPath)

  for (const ext of elementsByLocal(dataDoc, 'dataModelExt')) {
    const relId =
      ext.getAttribute('relId') ||
      getAttrNS(ext, 'relId') ||
      ext.getAttribute('r:relId')
    if (relId && dataRels.has(relId)) {
      return dataRels.get(relId)!.target
    }
  }

  for (const rel of dataRels.values()) {
    if (
      rel.type.toLowerCase().includes('diagramdrawing') ||
      rel.target.toLowerCase().includes('drawing')
    ) {
      return rel.target
    }
  }

  // Convention: data1.xml -> drawing1.xml
  const guess = dataPath.replace(/data(\d+)\.xml$/i, 'drawing$1.xml')
  if (guess !== dataPath && getZipEntry(zip, guess)) return guess

  return null
}

const luminanceOf = (hex: string): number => {
  const rgb = hex.replace('#', '')
  if (rgb.length < 6) return 0.5
  const r = parseInt(rgb.slice(0, 2), 16) || 0
  const g = parseInt(rgb.slice(2, 4), 16) || 0
  const b = parseInt(rgb.slice(4, 6), 16) || 0
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255
}

const contrastText = (fill: string): string => (luminanceOf(fill) > 0.55 ? '#000000' : '#FFFFFF')

const resolveFillColor = (fillParent: Element | null, scheme: ThemeColorMap, fallback: string): string => {
  if (!fillParent) return fallback
  return resolveColorElement(fillParent, scheme) || fallback
}

const resolveShapeFill = (sp: Element, scheme: ThemeColorMap): string => {
  const spPr = firstByLocal(sp, 'spPr')
  const solid = firstByLocal(spPr, 'solidFill')
  if (solid) return resolveFillColor(solid, scheme, '#5B9BD5')

  const grad = firstByLocal(spPr, 'gradFill')
  if (grad) {
    const stops = elementsByLocal(grad, 'gs')
    if (stops.length > 0) {
      return resolveFillColor(stops[0], scheme, '#5B9BD5')
    }
  }

  // Style list fill reference (common in diagram drawings)
  const styleFill = firstByLocal(firstByLocal(sp, 'style'), 'fillRef')
  if (styleFill) return resolveFillColor(styleFill, scheme, '#5B9BD5')

  return '#5B9BD5'
}

const resolveShapeTextColor = (sp: Element, scheme: ThemeColorMap, fill: string): string => {
  const txBody = firstByLocal(sp, 'txBody')
  if (txBody) {
    for (const rPr of elementsByLocal(txBody, 'rPr')) {
      const solid = firstByLocal(rPr, 'solidFill')
      const color = solid ? resolveColorElement(solid, scheme) : null
      if (color) return color
    }
    for (const defRPr of elementsByLocal(txBody, 'defRPr')) {
      const solid = firstByLocal(defRPr, 'solidFill')
      const color = solid ? resolveColorElement(solid, scheme) : null
      if (color) return color
    }
  }

  const fontRef = firstByLocal(firstByLocal(sp, 'style'), 'fontRef')
  if (fontRef) {
    const color = resolveColorElement(fontRef, scheme)
    if (color) return color
  }

  return contrastText(fill)
}

const readShapeText = (sp: Element): string => {
  const txBody = firstByLocal(sp, 'txBody')
  if (!txBody) return ''
  return elementsByLocal(txBody, 't')
    .map((t) => (t.textContent || '').trim())
    .filter(Boolean)
    .join(' ')
}

const readShapeFontSizePx = (sp: Element, shapeHeightPx: number): number => {
  const txBody = firstByLocal(sp, 'txBody')
  const candidates = [
    ...elementsByLocal(txBody, 'rPr'),
    ...elementsByLocal(txBody, 'defRPr'),
  ]
  for (const pr of candidates) {
    const sz = parseInt(pr.getAttribute('sz') || '', 10)
    if (sz > 0) {
      // OOXML sz is in hundredths of a point; 1pt ≈ 1.333px at 96dpi
      return Math.max(8, Math.min(48, (sz / 100) * (96 / 72)))
    }
  }
  // Scale with shape height so short SmartArt labels keep readable proportion
  return Math.max(9, Math.min(28, shapeHeightPx * 0.32))
}

const classifyShape = (sp: Element): SmartArtShape['shape'] => {
  const prst = firstByLocal(sp, 'prstGeom')?.getAttribute('prst') || ''
  if (prst === 'ellipse' || prst === 'circle') return 'ellipse'
  if (prst === 'roundRect') return 'roundRect'
  if (prst === 'rect') return 'rect'
  return 'other'
}

/**
 * Map drawing shapes using the group child coordinate space (chOff/chExt).
 * Only trust xfrm under grpSpPr — searching the whole spTree can pick a shape
 * xfrm and blow up the coordinate space (shapes clipped → SmartArt "missing").
 */
const readGroupChildSpace = (
  drawingDoc: Document
): { chOffX: number; chOffY: number; chExtCx: number; chExtCy: number } | null => {
  const spTree = firstByLocal(drawingDoc, 'spTree') || drawingDoc.documentElement
  const grpSpPr = firstByLocal(spTree, 'grpSpPr')
  if (!grpSpPr) return null

  // Prefer direct child xfrm of grpSpPr only
  let xfrm: Element | null = null
  for (const child of Array.from(grpSpPr.children) as Element[]) {
    if (localName(child) === 'xfrm') {
      xfrm = child
      break
    }
  }
  if (!xfrm) xfrm = firstByLocal(grpSpPr, 'xfrm')
  if (!xfrm) return null

  const chOff = firstByLocal(xfrm, 'chOff')
  const chExt = firstByLocal(xfrm, 'chExt')
  const chExtCx = readEmu(chExt, 'cx')
  const chExtCy = readEmu(chExt, 'cy')
  if (chExtCx <= 0 || chExtCy <= 0) return null

  return {
    chOffX: readEmu(chOff, 'x'),
    chOffY: readEmu(chOff, 'y'),
    chExtCx,
    chExtCy,
  }
}

const tightBoundsOfShapes = (
  shapes: Element[]
): { chOffX: number; chOffY: number; chExtCx: number; chExtCy: number } => {
  let minX = Number.POSITIVE_INFINITY
  let minY = Number.POSITIVE_INFINITY
  let maxX = 0
  let maxY = 0

  for (const sp of shapes) {
    const xfrm = firstByLocal(firstByLocal(sp, 'spPr'), 'xfrm') || firstByLocal(sp, 'xfrm')
    const off = firstByLocal(xfrm, 'off')
    const ext = firstByLocal(xfrm, 'ext')
    const x = readEmu(off, 'x')
    const y = readEmu(off, 'y')
    const cx = Math.max(readEmu(ext, 'cx'), 1)
    const cy = Math.max(readEmu(ext, 'cy'), 1)
    minX = Math.min(minX, x)
    minY = Math.min(minY, y)
    maxX = Math.max(maxX, x + cx)
    maxY = Math.max(maxY, y + cy)
  }

  if (!Number.isFinite(minX) || !Number.isFinite(minY)) {
    return { chOffX: 0, chOffY: 0, chExtCx: 1, chExtCy: 1 }
  }

  return {
    chOffX: minX,
    chOffY: minY,
    chExtCx: Math.max(maxX - minX, 1),
    chExtCy: Math.max(maxY - minY, 1),
  }
}

const mapShapesToPercents = (
  shapes: Element[],
  space: { chOffX: number; chOffY: number; chExtCx: number; chExtCy: number },
  scheme: ThemeColorMap,
  frameHeightPx: number
): SmartArtShape[] =>
  shapes.map((sp) => {
    const xfrm = firstByLocal(firstByLocal(sp, 'spPr'), 'xfrm') || firstByLocal(sp, 'xfrm')
    const off = firstByLocal(xfrm, 'off')
    const ext = firstByLocal(xfrm, 'ext')
    const x = readEmu(off, 'x')
    const y = readEmu(off, 'y')
    const cx = Math.max(readEmu(ext, 'cx'), 1)
    const cy = Math.max(readEmu(ext, 'cy'), 1)

    const leftPct = ((x - space.chOffX) / space.chExtCx) * 100
    const topPct = ((y - space.chOffY) / space.chExtCy) * 100
    const widthPct = (cx / space.chExtCx) * 100
    const heightPct = (cy / space.chExtCy) * 100
    const shapeHeightPx = Math.max(8, (heightPct / 100) * frameHeightPx)

    const fill = resolveShapeFill(sp, scheme)
    const textColor = resolveShapeTextColor(sp, scheme, fill)

    return {
      leftPct,
      topPct,
      widthPct,
      heightPct,
      text: readShapeText(sp),
      fill,
      textColor,
      fontSizePx: readShapeFontSizePx(sp, shapeHeightPx),
      shape: classifyShape(sp),
    }
  })

const shapesIntersectFrame = (shapes: SmartArtShape[]): boolean => {
  const visible = shapes.filter(
    (s) =>
      s.widthPct > 0.25 &&
      s.heightPct > 0.25 &&
      s.leftPct < 100 &&
      s.topPct < 100 &&
      s.leftPct + s.widthPct > 0 &&
      s.topPct + s.heightPct > 0
  )
  if (visible.length === 0) return false

  // Reject mappings that leave content as a tiny speck (bad chExt / nested groups)
  let minL = Infinity
  let minT = Infinity
  let maxR = -Infinity
  let maxB = -Infinity
  for (const s of visible) {
    minL = Math.min(minL, s.leftPct)
    minT = Math.min(minT, s.topPct)
    maxR = Math.max(maxR, s.leftPct + s.widthPct)
    maxB = Math.max(maxB, s.topPct + s.heightPct)
  }
  const coverW = Math.max(0, Math.min(100, maxR) - Math.max(0, minL))
  const coverH = Math.max(0, Math.min(100, maxB) - Math.max(0, minT))
  return coverW * coverH >= 8
}

const parseDrawingShapes = (
  drawingDoc: Document,
  scheme: ThemeColorMap,
  frameHeightPx: number,
  frameWidthEmu: number,
  frameHeightEmu: number
): SmartArtShape[] => {
  const spTree = firstByLocal(drawingDoc, 'spTree') || drawingDoc.documentElement
  // Prefer direct child shapes of spTree / nested grpSp, but skip empty placeholders
  const shapes = [...elementsByLocal(spTree, 'sp')].filter((sp) => {
    const xfrm = firstByLocal(firstByLocal(sp, 'spPr'), 'xfrm') || firstByLocal(sp, 'xfrm')
    return !!xfrm
  })
  if (shapes.length === 0) return []

  const groupSpace = readGroupChildSpace(drawingDoc)
  if (groupSpace) {
    const mapped = mapShapesToPercents(shapes, groupSpace, scheme, frameHeightPx)
    if (shapesIntersectFrame(mapped)) return mapped
    console.warn('SmartArt group chExt produced off-frame shapes; trying frame coordinate space')
  }

  // Many diagram drawings ship an empty grpSpPr (no chOff/chExt). Shape EMUs are
  // already in the graphicFrame's coordinate space — mapping against tight content
  // bounds stretches circles whenever the frame is taller/wider than the content.
  if (frameWidthEmu > 0 && frameHeightEmu > 0) {
    const frameSpace = {
      chOffX: 0,
      chOffY: 0,
      chExtCx: frameWidthEmu,
      chExtCy: frameHeightEmu,
    }
    const mapped = mapShapesToPercents(shapes, frameSpace, scheme, frameHeightPx)
    if (shapesIntersectFrame(mapped)) return mapped
    console.warn('SmartArt frame coordinate space missed shapes; falling back to tight bounds')
  }

  return mapShapesToPercents(shapes, tightBoundsOfShapes(shapes), scheme, frameHeightPx)
}

export const extractSlideSmartArt = async (
  zip: JSZip,
  slidePath: string,
  slideDoc: Document,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SmartArtExport[]> => {
  const frames = findSmartArtFrames(slideDoc)
  if (frames.length === 0) return []

  const slideRels = await loadRelationships(zip, slidePath)
  let scheme: ThemeColorMap = {}
  try {
    scheme = await loadSlideColorScheme(zip, slidePath)
  } catch (error) {
    console.warn('SmartArt theme colours unavailable; using defaults', error)
  }
  const results: SmartArtExport[] = []

  for (let i = 0; i < frames.length; i++) {
    const frame = frames[i]
    const transform = readTransform(frame, slideSize)
    const dataPath = resolveDiagramDataPath(frame, slideRels)

    let nodes: string[] = []
    let shapes: SmartArtShape[] = []

    if (dataPath) {
      const dataDoc = await loadXml(zip, dataPath)
      if (dataDoc) {
        nodes = extractNodeTexts(dataDoc)
        const drawingPath = await resolveDiagramDrawingPath(zip, dataPath, dataDoc)
        if (drawingPath) {
          const drawingDoc = await loadXml(zip, drawingPath)
          if (drawingDoc) {
            shapes = parseDrawingShapes(
              drawingDoc,
              scheme,
              transform.heightPx,
              transform.widthEmu,
              transform.heightEmu
            )
            console.log(`SmartArt drawing shapes from ${drawingPath}:`, shapes.length)
          }
        }
      }
    }

    const imagePath = findFallbackImagePath(frame, slideRels)
    const imageDataUrl = imagePath ? (await toImageDataUrl(zip, imagePath)) || undefined : undefined

    if (nodes.length === 0 && shapes.length === 0 && !imageDataUrl) {
      console.warn(`SmartArt on ${slidePath} had no extractable content`)
      continue
    }

    if (nodes.length === 0 && shapes.length > 0) {
      nodes = shapes.map((s) => s.text).filter(Boolean)
    }

    const title =
      firstByLocal(frame, 'cNvPr')?.getAttribute('name') || `SmartArt ${i + 1}`

    results.push({
      id: `smartart-${i + 1}`,
      title,
      nodes,
      shapes,
      imageDataUrl,
      leftPx: transform.leftPx,
      topPx: transform.topPx,
      widthPx: transform.widthPx,
      heightPx: transform.heightPx,
    })

    console.log(`Extracted SmartArt on ${slidePath}:`, {
      title,
      nodes: nodes.length,
      shapes: shapes.length,
      hasImage: !!imageDataUrl,
      box: `${Math.round(transform.widthPx)}x${Math.round(transform.heightPx)}`,
    })
  }

  return results
}

export const renderSmartArtHtml = (items: SmartArtExport[]): string => {
  if (items.length === 0) return ''

  return items
    .map((item) => {
      const style = `position:absolute;left:${item.leftPx.toFixed(1)}px;top:${item.topPx.toFixed(1)}px;width:${item.widthPx.toFixed(1)}px;height:${item.heightPx.toFixed(1)}px;z-index:4;`

      // Prefer Office's pre-rendered drawing shapes (true SmartArt look)
      const visibleShapes = item.shapes.filter(
        (shape) =>
          shape.widthPct > 0.25 &&
          shape.heightPct > 0.25 &&
          shape.leftPct < 100 &&
          shape.topPct < 100 &&
          shape.leftPct + shape.widthPct > 0 &&
          shape.topPct + shape.heightPct > 0
      )

      if (visibleShapes.length > 0) {
        const shapeHtml = visibleShapes
          .map((shape) => {
            const radius =
              shape.shape === 'ellipse'
                ? '50%'
                : shape.shape === 'roundRect'
                  ? '8px'
                  : '2px'
            const fontSize = Number.isFinite(shape.fontSizePx) ? shape.fontSizePx : 12
            return `
        <div class="smartart-shape smartart-shape-${shape.shape}" style="left:${shape.leftPct.toFixed(2)}%;top:${shape.topPct.toFixed(2)}%;width:${shape.widthPct.toFixed(2)}%;height:${shape.heightPct.toFixed(2)}%;background:${shape.fill};color:${shape.textColor};border-radius:${radius};font-size:${fontSize.toFixed(1)}px;">
          <div class="smartart-shape-text">${escapeHtml(shape.text || '')}</div>
        </div>`
          })
          .join('')

        return `
    <div class="smartart smartart-drawing" data-smartart="${item.id}" title="${escapeHtml(item.title)}" style="${style}">
      ${shapeHtml}
    </div>`
      }

      if (item.imageDataUrl) {
        return `
    <div class="smartart smartart-image-wrap" data-smartart="${item.id}" title="${escapeHtml(item.title)}" style="${style}">
      <img class="smartart-image" src="${item.imageDataUrl}" alt="${escapeHtml(item.title)}" />
    </div>`
      }

      const nodeHtml = item.nodes
        .map(
          (node, index) => `
        <div class="smartart-node" data-index="${index + 1}">
          <span class="smartart-node-index">${index + 1}</span>
          <span class="smartart-node-text">${escapeHtml(node)}</span>
        </div>`
        )
        .join('')

      return `
    <div class="smartart smartart-nodes-wrap" data-smartart="${item.id}" title="${escapeHtml(item.title)}" style="${style}">
      <div class="smartart-nodes" data-count="${item.nodes.length}">
        ${nodeHtml}
      </div>
    </div>`
    })
    .join('\n')
}

export const smartArtCss = `
      .smartart {
        box-sizing: border-box;
        overflow: hidden;
        position: absolute;
      }
      .smartart-drawing {
        background: transparent;
      }
      .smartart-shape {
        position: absolute;
        box-sizing: border-box;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 4px 6px;
        overflow: hidden;
        box-shadow: 0 1px 3px rgba(0,0,0,0.12);
      }
      .smartart-shape-text {
        font-size: inherit;
        font-weight: 600;
        line-height: 1.15;
        text-align: center;
        width: 100%;
        word-break: break-word;
      }
      .smartart-image {
        width: 100%;
        height: 100%;
        object-fit: contain;
        display: block;
      }
      .smartart-nodes {
        width: 100%;
        height: 100%;
        display: flex;
        flex-wrap: wrap;
        align-items: stretch;
        justify-content: center;
        gap: 12px;
        padding: 12px;
        box-sizing: border-box;
      }
      .smartart-node {
        flex: 1 1 140px;
        min-width: 120px;
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 12px 14px;
        border-radius: 8px;
        background: linear-gradient(135deg, #2f71b5 0%, #1f4e79 100%);
        color: #fff;
        box-shadow: 0 2px 6px rgba(0,0,0,0.18);
      }
      .smartart-node-index {
        flex: 0 0 auto;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: rgba(255,255,255,0.2);
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 13px;
      }
      .smartart-node-text {
        font-size: 15px;
        line-height: 1.35;
        font-weight: 600;
      }
`

export const collectSmartArtImageEmbedIds = (slideDoc: Document): Set<string> => {
  const ids = new Set<string>()
  for (const frame of findSmartArtFrames(slideDoc)) {
    let current: Element | null = frame
    while (current) {
      if (localName(current) === 'AlternateContent') {
        for (const blip of elementsByLocal(current, 'blip')) {
          const embed = getEmbedId(blip)
          if (embed) ids.add(embed)
        }
      }
      current = current.parentElement
    }
  }
  return ids
}
