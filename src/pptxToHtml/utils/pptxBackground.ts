import type JSZip from 'jszip'
import { DEFAULT_SLIDE_SIZE, type SlideSize } from './pptxSlideSize'

export type SlideBackground =
  | { type: 'solid'; color: string; css: string }
  | { type: 'gradient'; css: string }
  | { type: 'image'; css: string }
  | { type: 'pattern'; css: string }
  | { type: 'none'; css: string }

type ColorMap = Record<string, string>
export type ThemeColorMap = ColorMap

const REL_NS = {
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
}

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const DEFAULT_SCHEME: ColorMap = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
  bg1: 'FFFFFF',
  bg2: 'E7E6E6',
  tx1: '000000',
  tx2: '44546A',
}

const parser = () => new DOMParser()

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

const directChildrenByLocal = (parent: Element, name: string): Element[] =>
  Array.from(parent.children).filter((el) => localName(el) === name) as Element[]

const firstDirectChild = (parent: Element, name: string): Element | null =>
  directChildrenByLocal(parent, name)[0] || null

const firstByLocal = (parent: ParentNode | null | undefined, name: string): Element | null =>
  elementsByLocal(parent, name)[0] || null

const resolveZipPath = (fromPart: string, target: string): string => {
  const normalizedTarget = target.replace(/\\/g, '/')
  if (normalizedTarget.startsWith('/')) {
    return normalizedTarget.replace(/^\//, '')
  }
  const baseParts = fromPart.replace(/\\/g, '/').split('/').slice(0, -1)
  for (const part of normalizedTarget.split('/')) {
    if (part === '..') baseParts.pop()
    else if (part !== '.') baseParts.push(part)
  }
  return baseParts.join('/')
}

const relsPathFor = (partPath: string): string => {
  const parts = partPath.replace(/\\/g, '/').split('/')
  const file = parts.pop()
  return `${parts.join('/')}/_rels/${file}.rels`
}

const getZipEntry = (zip: JSZip, path: string) => {
  const candidates = [
    path,
    path.replace(/^\//, ''),
    `/${path.replace(/^\//, '')}`,
    path.replace(/\\/g, '/'),
  ]
  for (const candidate of candidates) {
    const entry = zip.file(candidate) || zip.files[candidate]
    if (entry && !entry.dir) return entry
  }

  // Case-insensitive fallback for oddly cased packages
  const needle = path.replace(/^\//, '').toLowerCase()
  const match = Object.keys(zip.files).find((key) => key.replace(/^\//, '').toLowerCase() === needle)
  return match ? zip.files[match] : null
}

export const loadXml = async (zip: JSZip, path: string): Promise<Document | null> => {
  const file = getZipEntry(zip, path)
  if (!file) return null
  const text = await file.async('text')
  const doc = parser().parseFromString(text, 'text/xml')
  if (doc.getElementsByTagName('parsererror').length > 0) {
    console.warn('Failed to parse XML part:', path)
    return null
  }
  return doc
}

export const loadRelationships = async (
  zip: JSZip,
  partPath: string
): Promise<Map<string, { type: string; target: string }>> => {
  const map = new Map<string, { type: string; target: string }>()
  const relsDoc = await loadXml(zip, relsPathFor(partPath))
  if (!relsDoc) return map

  for (const rel of elementsByLocal(relsDoc, 'Relationship')) {
    const id = rel.getAttribute('Id') || rel.getAttribute('id')
    const type = rel.getAttribute('Type') || rel.getAttribute('type') || ''
    const target = rel.getAttribute('Target') || rel.getAttribute('target') || ''
    if (id && target) {
      map.set(id, { type, target: resolveZipPath(partPath, target) })
    }
  }
  return map
}

const findRelByType = (
  rels: Map<string, { type: string; target: string }>,
  type: string
): string | null => {
  for (const value of rels.values()) {
    if (value.type === type || value.type.endsWith(type.split('/').pop() || '')) {
      return value.target
    }
  }
  return null
}

/** Resolve theme + master/slide clrMap for a slide. */
export const loadSlideColorScheme = async (
  zip: JSZip,
  slidePath: string
): Promise<ThemeColorMap> => {
  const slideRels = await loadRelationships(zip, slidePath)
  const layoutPath = findRelByType(slideRels, REL_NS.slideLayout)
  const layoutRels = layoutPath ? await loadRelationships(zip, layoutPath) : new Map()
  const masterPath = findRelByType(layoutRels, REL_NS.slideMaster)
  const masterDoc = masterPath ? await loadXml(zip, masterPath) : null
  const masterRels = masterPath ? await loadRelationships(zip, masterPath) : new Map()
  const themePath = findRelByType(masterRels, REL_NS.theme)
  const themeDoc = themePath ? await loadXml(zip, themePath) : null
  const slideDoc = await loadXml(zip, slidePath)

  return applyClrMapToScheme(
    applyClrMapToScheme(parseThemeColors(themeDoc), parseClrMap(masterDoc)),
    parseClrMap(slideDoc)
  )
}

export const parseClrMap = (doc: Document | null): ColorMap => {
  const map: ColorMap = {}
  if (!doc) return map

  const clrMap =
    firstByLocal(doc, 'clrMap') ||
    firstByLocal(firstByLocal(doc, 'clrMapOvr'), 'overrideClrMapping') ||
    firstByLocal(doc, 'overrideClrMapping')

  if (!clrMap) return map

  for (const attr of Array.from(clrMap.attributes)) {
    map[attr.name] = attr.value
  }
  return map
}

export const parseThemeColors = (themeDoc: Document | null): ColorMap => {
  const colors: ColorMap = { ...DEFAULT_SCHEME }
  if (!themeDoc) return colors

  const scheme = firstByLocal(themeDoc, 'clrScheme')
  if (!scheme) return colors

  for (const child of Array.from(scheme.children) as Element[]) {
    const name = localName(child)
    const srgb = firstDirectChild(child, 'srgbClr') || firstByLocal(child, 'srgbClr')
    const sysClr = firstDirectChild(child, 'sysClr') || firstByLocal(child, 'sysClr')
    const val = srgb?.getAttribute('val') || sysClr?.getAttribute('lastClr') || sysClr?.getAttribute('val')
    if (val) {
      colors[name] = val.toUpperCase()
    }
  }

  // Semantic aliases before clrMap remapping
  colors.bg1 = colors.lt1
  colors.bg2 = colors.lt2
  colors.tx1 = colors.dk1
  colors.tx2 = colors.dk2
  return colors
}

export const applyClrMapToScheme = (scheme: ColorMap, clrMap: ColorMap): ColorMap => {
  if (Object.keys(clrMap).length === 0) return scheme
  const mapped: ColorMap = { ...scheme }

  for (const [from, to] of Object.entries(clrMap)) {
    const target = scheme[to] || scheme[from]
    if (target) mapped[from] = target
  }

  return mapped
}

const clampByte = (n: number) => Math.max(0, Math.min(255, Math.round(n)))

const hexToRgb = (hex: string): [number, number, number] => {
  const clean = hex.replace('#', '').padStart(6, '0').slice(0, 6)
  return [
    parseInt(clean.slice(0, 2), 16) || 0,
    parseInt(clean.slice(2, 4), 16) || 0,
    parseInt(clean.slice(4, 6), 16) || 0,
  ]
}

const rgbToHex = (r: number, g: number, b: number) =>
  `#${[r, g, b].map((v) => clampByte(v).toString(16).padStart(2, '0')).join('')}`.toUpperCase()

const applyColorTransforms = (hex: string, colorEl: Element): string => {
  let [r, g, b] = hexToRgb(hex)

  for (const child of Array.from(colorEl.children) as Element[]) {
    const name = localName(child)
    const raw = parseInt(child.getAttribute('val') || '0', 10)

    if (name === 'lumMod') {
      const mod = raw / 100000
      r *= mod
      g *= mod
      b *= mod
    } else if (name === 'lumOff') {
      const off = raw / 100000
      r += 255 * off
      g += 255 * off
      b += 255 * off
    } else if (name === 'shade') {
      const shade = raw / 100000
      r *= shade
      g *= shade
      b *= shade
    } else if (name === 'tint') {
      const tint = raw / 100000
      r = r * tint + 255 * (1 - tint)
      g = g * tint + 255 * (1 - tint)
      b = b * tint + 255 * (1 - tint)
    }
  }

  return rgbToHex(r, g, b)
}

const readAlpha = (colorEl: Element | null): number | undefined => {
  if (!colorEl) return undefined
  const alpha = firstDirectChild(colorEl, 'alpha') || firstByLocal(colorEl, 'alpha')
  if (!alpha) return undefined
  const val = parseInt(alpha.getAttribute('val') || '', 10)
  if (Number.isNaN(val)) return undefined
  return Math.max(0, Math.min(1, val / 100000))
}

export const resolveColorElement = (colorParent: Element | null, scheme: ColorMap): string | null => {
  if (!colorParent) return null

  const srgb = firstDirectChild(colorParent, 'srgbClr') || firstByLocal(colorParent, 'srgbClr')
  if (srgb) {
    const val = srgb.getAttribute('val')
    if (!val) return null
    return applyColorTransforms(`#${val}`, srgb)
  }

  const schemeClr = firstDirectChild(colorParent, 'schemeClr') || firstByLocal(colorParent, 'schemeClr')
  if (schemeClr) {
    const name = schemeClr.getAttribute('val') || ''
    const mapped = scheme[name] || DEFAULT_SCHEME[name]
    if (!mapped) return null
    return applyColorTransforms(`#${mapped}`, schemeClr)
  }

  const sysClr = firstDirectChild(colorParent, 'sysClr') || firstByLocal(colorParent, 'sysClr')
  if (sysClr) {
    const val = sysClr.getAttribute('lastClr') || sysClr.getAttribute('val')
    if (!val) return null
    return applyColorTransforms(`#${val}`, sysClr)
  }

  const prstClr = firstDirectChild(colorParent, 'prstClr') || firstByLocal(colorParent, 'prstClr')
  if (prstClr) {
    const preset = (prstClr.getAttribute('val') || '').toLowerCase()
    const presets: Record<string, string> = {
      black: '#000000',
      white: '#FFFFFF',
      red: '#FF0000',
      green: '#00FF00',
      blue: '#0000FF',
      yellow: '#FFFF00',
      gray: '#808080',
      grey: '#808080',
    }
    if (presets[preset]) return applyColorTransforms(presets[preset], prstClr)
  }

  return null
}

const colorWithAlpha = (hex: string, alpha: number | undefined): string => {
  if (alpha === undefined || alpha >= 0.999) return hex
  const [r, g, b] = hexToRgb(hex)
  return `rgba(${r}, ${g}, ${b}, ${alpha})`
}

const getPrimaryColorNode = (parent: Element): Element | null =>
  firstDirectChild(parent, 'srgbClr') ||
  firstDirectChild(parent, 'schemeClr') ||
  firstDirectChild(parent, 'sysClr') ||
  firstByLocal(parent, 'srgbClr') ||
  firstByLocal(parent, 'schemeClr') ||
  firstByLocal(parent, 'sysClr')

const getSlideBackgroundNode = (doc: Document | null): Element | null => {
  if (!doc) return null
  const cSld = firstByLocal(doc, 'cSld')
  if (!cSld) return null
  return firstDirectChild(cSld, 'bg') || firstByLocal(cSld, 'bg')
}

const mimeFromPath = (path: string): string => {
  const ext = path.split('.').pop()?.toLowerCase()
  switch (ext) {
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg'
    case 'gif':
      return 'image/gif'
    case 'svg':
      return 'image/svg+xml'
    case 'webp':
      return 'image/webp'
    case 'png':
    default:
      return 'image/png'
  }
}

const parseSolidFill = (fillEl: Element, scheme: ColorMap): SlideBackground | null => {
  const color = resolveColorElement(fillEl, scheme)
  if (!color) return null
  const alpha = readAlpha(getPrimaryColorNode(fillEl))
  const cssColor = colorWithAlpha(color, alpha)
  return {
    type: 'solid',
    color: cssColor,
    css: `background-color: ${cssColor}; background-image: none;`,
  }
}

const parseGradientFill = (fillEl: Element, scheme: ColorMap): SlideBackground | null => {
  const stops = elementsByLocal(fillEl, 'gs')
    .map((gs) => {
      const pos = parseInt(gs.getAttribute('pos') || '0', 10) / 1000
      const color = resolveColorElement(gs, scheme)
      if (!color) return null
      const alpha = readAlpha(getPrimaryColorNode(gs))
      return { pos, color: colorWithAlpha(color, alpha) }
    })
    .filter((s): s is { pos: number; color: string } => !!s)
    .sort((a, b) => a.pos - b.pos)

  if (stops.length === 0) return null

  const lin = firstDirectChild(fillEl, 'lin') || firstByLocal(fillEl, 'lin')
  let angle = 90
  if (lin) {
    const ang = parseInt(lin.getAttribute('ang') || '0', 10)
    angle = ((ang / 60000) + 90) % 360
  }

  const path = firstDirectChild(fillEl, 'path') || firstByLocal(fillEl, 'path')
  const stopCss = stops.map((s) => `${s.color} ${s.pos}%`).join(', ')
  const gradient = path
    ? `radial-gradient(circle, ${stopCss})`
    : `linear-gradient(${angle}deg, ${stopCss})`

  return {
    type: 'gradient',
    css: `background-color: ${stops[0].color}; background-image: ${gradient}; background-repeat: no-repeat; background-size: cover;`,
  }
}

const parseBlipFill = async (
  fillEl: Element,
  zip: JSZip,
  rels: Map<string, { type: string; target: string }>,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SlideBackground | null> => {
  const blip = firstDirectChild(fillEl, 'blip') || firstByLocal(fillEl, 'blip')
  if (!blip) return null

  const embed =
    blip.getAttributeNS(R_NS, 'embed') ||
    blip.getAttribute('r:embed') ||
    blip.getAttribute('embed')
  if (!embed) return null

  const rel = rels.get(embed)
  if (!rel) return null

  const media = getZipEntry(zip, rel.target)
  if (!media) return null

  try {
    const base64 = await media.async('base64')
    const mime = mimeFromPath(rel.target)
    const dataUrl = `data:${mime};base64,${base64}`
    const stretch = firstDirectChild(fillEl, 'stretch') || firstByLocal(fillEl, 'stretch')
    const tile = firstDirectChild(fillEl, 'tile') || firstByLocal(fillEl, 'tile')

    let size = 'cover'
    let repeat = 'no-repeat'
    let position = 'center center'

    if (tile) {
      const algn = (tile.getAttribute('algn') || 'tl').toLowerCase()
      size = 'auto'
      repeat = 'repeat'
      position =
        algn.includes('ctr') || algn === 'ctr'
          ? 'center center'
          : algn.includes('t') && algn.includes('r')
            ? 'right top'
            : algn.includes('b') && algn.includes('r')
              ? 'right bottom'
              : algn.includes('b')
                ? 'left bottom'
                : algn.includes('r')
                  ? 'right top'
                  : 'left top'
    } else if (stretch) {
      const fillRect = firstDirectChild(stretch, 'fillRect') || firstByLocal(stretch, 'fillRect')
      const readPct = (attr: string): number => {
        const raw = parseInt(fillRect?.getAttribute(attr) || '0', 10)
        return Number.isFinite(raw) ? raw / 100000 : 0
      }
      const l = readPct('l')
      const r = readPct('r')
      const t = readPct('t')
      const b = readPct('b')
      const wFrac = 1 - l - r
      const hFrac = 1 - t - b

      // Default empty fillRect => stretch to shape (may distort).
      // Non-zero offsets define the destination rect for the image (OOXML),
      // often used to cover while preserving aspect (e.g. t/b = -65%).
      if (!fillRect || (l === 0 && r === 0 && t === 0 && b === 0)) {
        size = '100% 100%'
        position = 'center center'
      } else if (wFrac !== 0 && hFrac !== 0) {
        size = `${(Math.abs(wFrac) * 100).toFixed(4)}% ${(Math.abs(hFrac) * 100).toFixed(4)}%`
        position = `${(l * slideSize.widthPx).toFixed(2)}px ${(t * slideSize.heightPx).toFixed(2)}px`
      } else {
        size = 'cover'
        position = 'center center'
      }
      repeat = 'no-repeat'
    }

    return {
      type: 'image',
      css: `background-color: transparent; background-image: url('${dataUrl}'); background-size: ${size}; background-position: ${position}; background-repeat: ${repeat};`,
    }
  } catch (error) {
    console.error('Failed to load background image:', error)
    return null
  }
}

const svgPatternDataUrl = (width: number, height: number, body: string): string => {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${body}</svg>`
  // Use single quotes so this is safe inside HTML style="..." attributes.
  return `url('data:image/svg+xml,${encodeURIComponent(svg)}')`
}

const patternTile = (
  preset: string,
  fg: string,
  bg: string
): { image: string; size: string } => {
  const rect = (w: number, h: number, fill = bg) =>
    `<rect width="${w}" height="${h}" fill="${fill}"/>`

  switch (preset) {
    case 'pct5':
      return {
        image: svgPatternDataUrl(12, 12, `${rect(12, 12)}<circle cx="2" cy="2" r="1.2" fill="${fg}"/>`),
        size: '12px 12px',
      }
    case 'pct10':
      return {
        image: svgPatternDataUrl(10, 10, `${rect(10, 10)}<circle cx="2" cy="2" r="1.4" fill="${fg}"/>`),
        size: '10px 10px',
      }
    case 'pct20':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<circle cx="2" cy="2" r="1.6" fill="${fg}"/>`),
        size: '8px 8px',
      }
    case 'pct25':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<rect x="0" y="0" width="4" height="4" fill="${fg}"/><rect x="4" y="4" width="4" height="4" fill="${fg}"/>`),
        size: '8px 8px',
      }
    case 'pct30':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<circle cx="2" cy="2" r="2" fill="${fg}"/><circle cx="6" cy="6" r="1.4" fill="${fg}"/>`),
        size: '8px 8px',
      }
    case 'pct40':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<circle cx="2" cy="2" r="2.2" fill="${fg}"/><circle cx="6" cy="6" r="2" fill="${fg}"/>`),
        size: '8px 8px',
      }
    case 'pct50':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<rect width="4" height="4" fill="${fg}"/><rect x="4" y="4" width="4" height="4" fill="${fg}"/>`),
        size: '8px 8px',
      }
    case 'pct60':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8, fg)}<circle cx="2" cy="2" r="1.8" fill="${bg}"/><circle cx="6" cy="6" r="1.6" fill="${bg}"/>`),
        size: '8px 8px',
      }
    case 'pct70':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8, fg)}<circle cx="2" cy="2" r="1.5" fill="${bg}"/>`),
        size: '8px 8px',
      }
    case 'pct75':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8, fg)}<rect x="0" y="0" width="4" height="4" fill="${bg}"/><rect x="4" y="4" width="4" height="4" fill="${bg}"/>`),
        size: '8px 8px',
      }
    case 'pct80':
      return {
        image: svgPatternDataUrl(10, 10, `${rect(10, 10, fg)}<circle cx="2" cy="2" r="1.4" fill="${bg}"/>`),
        size: '10px 10px',
      }
    case 'pct90':
      return {
        image: svgPatternDataUrl(12, 12, `${rect(12, 12, fg)}<circle cx="2" cy="2" r="1.2" fill="${bg}"/>`),
        size: '12px 12px',
      }
    case 'horz':
    case 'ltHorz':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<rect y="0" width="8" height="2" fill="${fg}"/>`),
        size: preset === 'ltHorz' ? '8px 12px' : '8px 8px',
      }
    case 'dkHorz':
    case 'narHorz':
      return {
        image: svgPatternDataUrl(8, 6, `${rect(8, 6)}<rect y="0" width="8" height="${preset === 'dkHorz' ? 4 : 3}" fill="${fg}"/>`),
        size: '8px 6px',
      }
    case 'dashHorz':
      return {
        image: svgPatternDataUrl(12, 8, `${rect(12, 8)}<rect x="0" y="2" width="6" height="2" fill="${fg}"/>`),
        size: '12px 8px',
      }
    case 'vert':
    case 'ltVert':
      return {
        image: svgPatternDataUrl(8, 8, `${rect(8, 8)}<rect x="0" width="2" height="8" fill="${fg}"/>`),
        size: preset === 'ltVert' ? '12px 8px' : '8px 8px',
      }
    case 'dkVert':
    case 'narVert':
      return {
        image: svgPatternDataUrl(6, 8, `${rect(6, 8)}<rect x="0" width="${preset === 'dkVert' ? 4 : 3}" height="8" fill="${fg}"/>`),
        size: '6px 8px',
      }
    case 'dashVert':
      return {
        image: svgPatternDataUrl(8, 12, `${rect(8, 12)}<rect x="2" y="0" width="2" height="6" fill="${fg}"/>`),
        size: '8px 12px',
      }
    case 'cross':
    case 'smGrid':
    case 'lgGrid': {
      const s = preset === 'lgGrid' ? 16 : 8
      const stroke = preset === 'lgGrid' ? 1.5 : 1
      return {
        image: svgPatternDataUrl(
          s,
          s,
          `${rect(s, s)}<path d="M0 ${s / 2} H${s} M${s / 2} 0 V${s}" stroke="${fg}" stroke-width="${stroke}" fill="none"/>`
        ),
        size: `${s}px ${s}px`,
      }
    }
    case 'dotGrid':
      return {
        image: svgPatternDataUrl(
          10,
          10,
          `${rect(10, 10)}<circle cx="2" cy="2" r="1.1" fill="${fg}"/><circle cx="7" cy="2" r="1.1" fill="${fg}"/><circle cx="2" cy="7" r="1.1" fill="${fg}"/><circle cx="7" cy="7" r="1.1" fill="${fg}"/>`
        ),
        size: '10px 10px',
      }
    case 'dnDiag':
    case 'ltDnDiag':
    case 'dkDnDiag':
    case 'wdDnDiag':
    case 'dashDnDiag': {
      const size = preset === 'wdDnDiag' ? 16 : preset === 'ltDnDiag' ? 12 : 8
      const width = preset === 'dkDnDiag' || preset === 'wdDnDiag' ? 3 : 1.5
      const dash = preset === 'dashDnDiag' ? 'stroke-dasharray="3 3"' : ''
      return {
        image: svgPatternDataUrl(
          size,
          size,
          `${rect(size, size)}<path d="M-2 ${size * 0.25} L${size * 0.25} -2 M-2 ${size * 0.75} L${size * 0.75} -2 M${size * 0.25} ${size + 2} L${size + 2} ${size * 0.25} M${size * 0.75} ${size + 2} L${size + 2} ${size * 0.75}" stroke="${fg}" stroke-width="${width}" ${dash} fill="none"/>`
        ),
        size: `${size}px ${size}px`,
      }
    }
    case 'upDiag':
    case 'ltUpDiag':
    case 'dkUpDiag':
    case 'wdUpDiag':
    case 'dashUpDiag': {
      const size = preset === 'wdUpDiag' ? 16 : preset === 'ltUpDiag' ? 12 : 8
      const width = preset === 'dkUpDiag' || preset === 'wdUpDiag' ? 3 : 1.5
      const dash = preset === 'dashUpDiag' ? 'stroke-dasharray="3 3"' : ''
      return {
        image: svgPatternDataUrl(
          size,
          size,
          `${rect(size, size)}<path d="M-2 ${size * 0.75} L${size * 0.25} ${size + 2} M-2 ${size * 0.25} L${size * 0.75} ${size + 2} M${size * 0.25} -2 L${size + 2} ${size * 0.75} M${size * 0.75} -2 L${size + 2} ${size * 0.25}" stroke="${fg}" stroke-width="${width}" ${dash} fill="none"/>`
        ),
        size: `${size}px ${size}px`,
      }
    }
    case 'diagCross':
      return {
        image: svgPatternDataUrl(
          10,
          10,
          `${rect(10, 10)}<path d="M0 0 L10 10 M10 0 L0 10" stroke="${fg}" stroke-width="1.5" fill="none"/>`
        ),
        size: '10px 10px',
      }
    case 'smCheck':
    case 'lgCheck': {
      const s = preset === 'lgCheck' ? 16 : 8
      const half = s / 2
      return {
        image: svgPatternDataUrl(
          s,
          s,
          `${rect(s, s)}<rect width="${half}" height="${half}" fill="${fg}"/><rect x="${half}" y="${half}" width="${half}" height="${half}" fill="${fg}"/>`
        ),
        size: `${s}px ${s}px`,
      }
    }
    case 'horzBrick':
      return {
        image: svgPatternDataUrl(
          16,
          12,
          `${rect(16, 12)}<rect x="0" y="0" width="16" height="1" fill="${fg}"/><rect x="0" y="6" width="16" height="1" fill="${fg}"/><rect x="0" y="0" width="1" height="6" fill="${fg}"/><rect x="8" y="6" width="1" height="6" fill="${fg}"/>`
        ),
        size: '16px 12px',
      }
    case 'diagBrick':
      return {
        image: svgPatternDataUrl(
          12,
          12,
          `${rect(12, 12)}<path d="M0 6 H12 M6 0 V12" stroke="${fg}" stroke-width="1.2" fill="none"/><path d="M0 0 L12 12" stroke="${fg}" stroke-width="1" fill="none"/>`
        ),
        size: '12px 12px',
      }
    case 'solidDmnd':
    case 'openDmnd':
    case 'dotDmnd':
      return {
        image: svgPatternDataUrl(
          12,
          12,
          `${rect(12, 12)}<path d="M6 1 L11 6 L6 11 L1 6 Z" ${
            preset === 'solidDmnd'
              ? `fill="${fg}"`
              : `fill="none" stroke="${fg}" stroke-width="1.2"`
          }/>${
            preset === 'dotDmnd'
              ? `<circle cx="6" cy="6" r="1.2" fill="${fg}"/>`
              : ''
          }`
        ),
        size: '12px 12px',
      }
    case 'plaid':
      return {
        image: svgPatternDataUrl(
          12,
          12,
          `${rect(12, 12)}<rect x="0" y="0" width="12" height="3" fill="${fg}" fill-opacity="0.55"/><rect x="0" y="0" width="3" height="12" fill="${fg}" fill-opacity="0.55"/>`
        ),
        size: '12px 12px',
      }
    case 'weave':
    case 'trellis':
      return {
        image: svgPatternDataUrl(
          10,
          10,
          `${rect(10, 10)}<path d="M0 0 L10 10 M10 0 L0 10 M0 5 H10 M5 0 V10" stroke="${fg}" stroke-width="1.1" fill="none"/>`
        ),
        size: '10px 10px',
      }
    case 'divot':
      return {
        image: svgPatternDataUrl(
          12,
          12,
          `${rect(12, 12)}<path d="M3 3 l2 0 l-1 2 Z M8 8 l2 0 l-1 2 Z" fill="${fg}"/>`
        ),
        size: '12px 12px',
      }
    case 'shingle':
      return {
        image: svgPatternDataUrl(
          14,
          10,
          `${rect(14, 10)}<path d="M0 8 Q3 2 7 8 Q10 12 14 8" fill="none" stroke="${fg}" stroke-width="1.4"/>`
        ),
        size: '14px 10px',
      }
    case 'wave':
    case 'zigZag':
      return {
        image: svgPatternDataUrl(
          12,
          8,
          `${rect(12, 8)}<path d="M0 4 L3 1 L6 4 L9 1 L12 4" fill="none" stroke="${fg}" stroke-width="1.4"/>`
        ),
        size: '12px 8px',
      }
    case 'sphere':
      return {
        image: svgPatternDataUrl(
          12,
          12,
          `${rect(12, 12)}<circle cx="6" cy="6" r="3.2" fill="${fg}"/><circle cx="5" cy="5" r="1" fill="${bg}" fill-opacity="0.55"/>`
        ),
        size: '12px 12px',
      }
    case 'smConfetti':
    case 'lgConfetti': {
      const s = preset === 'lgConfetti' ? 14 : 10
      return {
        image: svgPatternDataUrl(
          s,
          s,
          `${rect(s, s)}<rect x="1" y="2" width="2" height="2" fill="${fg}"/><rect x="${s - 4}" y="${s - 5}" width="2" height="2" fill="${fg}"/><rect x="${s / 2}" y="1" width="2" height="2" fill="${fg}"/>`
        ),
        size: `${s}px ${s}px`,
      }
    }
    default:
      // Fallback dotted pattern (OOXML default when prst omitted is pct5)
      return {
        image: svgPatternDataUrl(10, 10, `${rect(10, 10)}<circle cx="2" cy="2" r="1.3" fill="${fg}"/>`),
        size: '10px 10px',
      }
  }
}

const parsePatternFill = (fillEl: Element, scheme: ColorMap): SlideBackground | null => {
  const preset = (fillEl.getAttribute('prst') || 'pct5').trim()

  const fgNode = firstDirectChild(fillEl, 'fgClr') || firstByLocal(fillEl, 'fgClr')
  const bgNode = firstDirectChild(fillEl, 'bgClr') || firstByLocal(fillEl, 'bgClr')

  const fgRaw = resolveColorElement(fgNode, scheme) || '#000000'
  const bgRaw = resolveColorElement(bgNode, scheme) || '#FFFFFF'
  const fg = colorWithAlpha(fgRaw, readAlpha(getPrimaryColorNode(fgNode || fillEl)))
  const bg = colorWithAlpha(bgRaw, readAlpha(getPrimaryColorNode(bgNode || fillEl)))

  const tile = patternTile(preset, fg, bg)
  console.log('Resolved pattern fill:', preset, { fg, bg })

  return {
    type: 'pattern',
    css: `background-color: ${bg}; background-image: ${tile.image}; background-size: ${tile.size}; background-repeat: repeat; background-position: 0 0;`,
  }
}

const parseFillElement = async (
  fillEl: Element | null,
  zip: JSZip,
  rels: Map<string, { type: string; target: string }>,
  scheme: ColorMap,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SlideBackground | null> => {
  if (!fillEl) return null
  const name = localName(fillEl)

  if (name === 'solidFill') return parseSolidFill(fillEl, scheme)
  if (name === 'gradFill') return parseGradientFill(fillEl, scheme)
  if (name === 'blipFill') return parseBlipFill(fillEl, zip, rels, slideSize)
  if (name === 'pattFill') return parsePatternFill(fillEl, scheme)
  if (name === 'noFill') {
    return { type: 'none', css: 'background-color: transparent; background-image: none;' }
  }
  return null
}

const parseBgPr = async (
  bgPr: Element,
  zip: JSZip,
  rels: Map<string, { type: string; target: string }>,
  scheme: ColorMap,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SlideBackground | null> => {
  for (const child of Array.from(bgPr.children) as Element[]) {
    const parsed = await parseFillElement(child, zip, rels, scheme, slideSize)
    if (parsed && parsed.type !== 'none') return parsed
    if (parsed?.type === 'none') return parsed
  }
  return null
}

const parseBgRef = async (
  bgRef: Element,
  zip: JSZip,
  rels: Map<string, { type: string; target: string }>,
  scheme: ColorMap,
  themeDoc: Document | null,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SlideBackground | null> => {
  if (!themeDoc) {
    const color = resolveColorElement(bgRef, scheme)
    if (color) {
      return {
        type: 'solid',
        color,
        css: `background-color: ${color}; background-image: none;`,
      }
    }
    return null
  }

  const idx = parseInt(bgRef.getAttribute('idx') || '0', 10)
  const fillStyles = firstByLocal(themeDoc, 'bgFillStyleLst')
  if (fillStyles) {
    const fills = Array.from(fillStyles.children) as Element[]
    const styleIndex = idx >= 1001 ? idx - 1001 : Math.max(0, idx - 1)
    const fill = fills[styleIndex]
    if (fill) {
      const fromStyle = await parseFillElement(fill, zip, rels, scheme, slideSize)
      if (fromStyle) {
        // Prefer explicit color on bgRef when style is a plain solid placeholder
        if (fromStyle.type === 'solid') {
          const override = resolveColorElement(bgRef, scheme)
          if (override) {
            return {
              type: 'solid',
              color: override,
              css: `background-color: ${override}; background-image: none;`,
            }
          }
        }
        return fromStyle
      }
    }
  }

  const color = resolveColorElement(bgRef, scheme)
  if (color) {
    return {
      type: 'solid',
      color,
      css: `background-color: ${color}; background-image: none;`,
    }
  }

  return null
}

const parseExplicitBackground = async (
  doc: Document | null,
  zip: JSZip,
  partPath: string | null,
  scheme: ColorMap,
  themeDoc: Document | null,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SlideBackground | null> => {
  const bg = getSlideBackgroundNode(doc)
  if (!bg || !partPath) return null

  const rels = await loadRelationships(zip, partPath)
  const bgPr = firstDirectChild(bg, 'bgPr') || firstByLocal(bg, 'bgPr')
  if (bgPr) return parseBgPr(bgPr, zip, rels, scheme, slideSize)

  const bgRef = firstDirectChild(bg, 'bgRef') || firstByLocal(bg, 'bgRef')
  if (bgRef) return parseBgRef(bgRef, zip, rels, scheme, themeDoc, slideSize)

  return null
}

const readEmu = (el: Element | null, attr: string): number => {
  if (!el) return 0
  return parseInt(el.getAttribute(attr) || '0', 10) || 0
}

const isFullSlideShape = (shape: Element, slideSize: SlideSize): boolean => {
  const xfrm = firstByLocal(shape, 'xfrm')
  if (!xfrm) return false
  const off = firstDirectChild(xfrm, 'off') || firstByLocal(xfrm, 'off')
  const ext = firstDirectChild(xfrm, 'ext') || firstByLocal(xfrm, 'ext')
  const x = readEmu(off, 'x')
  const y = readEmu(off, 'y')
  const cx = readEmu(ext, 'cx')
  const cy = readEmu(ext, 'cy')

  const coversWidth = cx >= slideSize.widthEmu * 0.9
  const coversHeight = cy >= slideSize.heightEmu * 0.9
  const nearOrigin = x <= slideSize.widthEmu * 0.05 && y <= slideSize.heightEmu * 0.05
  return coversWidth && coversHeight && nearOrigin
}

const findShapeFill = (shape: Element): Element | null => {
  const spPr = firstByLocal(shape, 'spPr')
  if (!spPr) return null
  return (
    firstDirectChild(spPr, 'solidFill') ||
    firstDirectChild(spPr, 'gradFill') ||
    firstDirectChild(spPr, 'blipFill') ||
    firstDirectChild(spPr, 'pattFill') ||
    firstByLocal(spPr, 'solidFill') ||
    firstByLocal(spPr, 'gradFill') ||
    firstByLocal(spPr, 'blipFill') ||
    firstByLocal(spPr, 'pattFill')
  )
}

const parseShapeBackground = async (
  doc: Document | null,
  zip: JSZip,
  partPath: string | null,
  scheme: ColorMap,
  slideSize: SlideSize
): Promise<SlideBackground | null> => {
  if (!doc || !partPath) return null
  const rels = await loadRelationships(zip, partPath)
  const shapes = [
    ...elementsByLocal(doc, 'sp'),
    ...elementsByLocal(doc, 'pic'),
  ]

  for (const shape of shapes) {
    if (!isFullSlideShape(shape, slideSize)) continue
    const fill = findShapeFill(shape) || firstByLocal(shape, 'blipFill')
    const parsed = await parseFillElement(fill, zip, rels, scheme, slideSize)
    if (parsed && parsed.type !== 'none') {
      console.log(`Background fill found on full-slide shape in ${partPath}`)
      return parsed
    }
  }

  return null
}

const isEffectivelyDefaultWhite = (bg: SlideBackground): boolean => {
  if (bg.type !== 'solid') return false
  const normalized = bg.color.replace(/\s/g, '').toUpperCase()
  return normalized === '#FFFFFF' || normalized === '#FFF' || normalized === 'WHITE'
}

export const resolveSlideBackground = async (
  zip: JSZip,
  slidePath: string,
  slideDoc: Document,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<SlideBackground> => {
  const slideRels = await loadRelationships(zip, slidePath)
  const layoutPath = findRelByType(slideRels, REL_NS.slideLayout)
  const layoutDoc = layoutPath ? await loadXml(zip, layoutPath) : null

  const layoutRels = layoutPath ? await loadRelationships(zip, layoutPath) : new Map()
  const masterPath = findRelByType(layoutRels, REL_NS.slideMaster)
  const masterDoc = masterPath ? await loadXml(zip, masterPath) : null

  const masterRels = masterPath ? await loadRelationships(zip, masterPath) : new Map()
  const themePath = findRelByType(masterRels, REL_NS.theme)
  const themeDoc = themePath ? await loadXml(zip, themePath) : null

  const themeColors = parseThemeColors(themeDoc)
  const masterClrMap = parseClrMap(masterDoc)
  const slideClrMap = parseClrMap(slideDoc)
  const scheme = applyClrMapToScheme(
    applyClrMapToScheme(themeColors, masterClrMap),
    slideClrMap
  )

  console.log('Background resolution paths:', {
    slidePath,
    layoutPath,
    masterPath,
    themePath,
    slideRels: slideRels.size,
    masterClrMap,
    slideClrMap,
    slideSize,
  })

  // Prefer the most specific visual fill:
  // 1) slide p:bg, 2) full-slide shape fill on the slide,
  // then layout/master equivalents. Shape fills often are the
  // visible "background fill" even when a master p:bg exists.
  const orderedLookups: Array<{
    doc: Document | null
    path: string | null
    label: string
    kind: 'bg' | 'shape'
  }> = [
    { doc: slideDoc, path: slidePath, label: 'slide', kind: 'bg' },
    { doc: slideDoc, path: slidePath, label: 'slide', kind: 'shape' },
    { doc: layoutDoc, path: layoutPath, label: 'layout', kind: 'bg' },
    { doc: layoutDoc, path: layoutPath, label: 'layout', kind: 'shape' },
    { doc: masterDoc, path: masterPath, label: 'master', kind: 'bg' },
    { doc: masterDoc, path: masterPath, label: 'master', kind: 'shape' },
  ]

  for (const level of orderedLookups) {
    const resolved =
      level.kind === 'bg'
        ? await parseExplicitBackground(level.doc, zip, level.path, scheme, themeDoc, slideSize)
        : await parseShapeBackground(level.doc, zip, level.path, scheme, slideSize)

    if (resolved && resolved.type !== 'none') {
      console.log(
        `Background for ${slidePath} from ${level.label} ${level.kind}:`,
        resolved.type,
        resolved.type === 'solid' ? resolved.color : ''
      )
      return resolved
    }
  }

  const fallback = scheme.bg1 || scheme.lt1 || 'FFFFFF'
  const color = `#${fallback.replace('#', '')}`
  const result: SlideBackground = {
    type: 'solid',
    color,
    css: `background-color: ${color}; background-image: none;`,
  }
  console.log(`Background for ${slidePath} fell back to theme bg1:`, result.color, {
    isDefaultWhite: isEffectivelyDefaultWhite(result),
  })
  return result
}

export const backgroundInlineStyle = (background: SlideBackground): string =>
  background.css
    .split(';')
    .map((part) => part.trim())
    .filter(Boolean)
    .join('; ')

export const backgroundCssForSlide = (slideNumber: string, background: SlideBackground): string => {
  const important = background.css
    .split(';')
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => `${part} !important`)
    .join('; ')

  return `
      .slide.slide-${slideNumber},
      .slide[data-slide="${slideNumber}"] {
        ${important};
      }
`
}
