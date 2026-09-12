import type JSZip from 'jszip'
import {
  loadRelationships,
  loadSlideColorScheme,
  loadXml,
  resolveColorElement,
  type ThemeColorMap,
} from './pptxBackground'
import { DEFAULT_SLIDE_SIZE, emuToPx, type SlideSize } from './pptxSlideSize'

const REL_LAYOUT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const REL_MASTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const REL_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

export interface TextRun {
  text: string
  fontSizePx: number
  bold: boolean
  italic: boolean
  underline: boolean
  color: string
  fontFamily: string
}

export interface TextParagraph {
  level: number
  align: string
  bullet: string | null
  marginLeftPx: number
  indentPx: number
  runs: TextRun[]
}

export interface TextShapeExport {
  id: string
  kind: 'title' | 'body' | 'other'
  leftPx: number
  topPx: number
  widthPx: number
  heightPx: number
  padding: { left: number; top: number; right: number; bottom: number }
  verticalAlign: string
  paragraphs: TextParagraph[]
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

const directChildrenByLocal = (parent: Element | null | undefined, name: string): Element[] => {
  if (!parent) return []
  return Array.from(parent.children).filter((el) => localName(el) === name) as Element[]
}

const firstByLocal = (parent: ParentNode | null | undefined, name: string): Element | null =>
  elementsByLocal(parent, name)[0] || null

const firstDirect = (parent: Element | null | undefined, name: string): Element | null =>
  directChildrenByLocal(parent, name)[0] || null

const readEmu = (el: Element | null, attr: string): number => {
  if (!el) return 0
  return parseInt(el.getAttribute(attr) || '0', 10) || 0
}

const escapeHtml = (text: string): string =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

const findRelByType = (
  rels: Map<string, { type: string; target: string }>,
  type: string
): string | null => {
  const suffix = type.split('/').pop() || ''
  for (const value of rels.values()) {
    if (value.type === type || value.type.endsWith(suffix)) return value.target
  }
  return null
}

const szToPx = (sz: number): number => (sz / 100) * (96 / 72)

const emuAttrToPx = (val: string | null | undefined, fallback = 0): number => {
  if (!val) return fallback
  const n = parseInt(val, 10)
  return Number.isFinite(n) ? emuToPx(n) : fallback
}

interface Box {
  x: number
  y: number
  cx: number
  cy: number
}

interface LevelStyle {
  marL?: number
  indent?: number
  algn?: string
  bullet?: string | null
  bulletNone?: boolean
  sz?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  fontFamily?: string
}

interface ThemeFonts {
  major: string
  minor: string
}

const readXfrmBox = (container: Element | null): Box | null => {
  if (!container) return null
  const xfrm =
    firstDirect(firstDirect(container, 'spPr'), 'xfrm') ||
    firstByLocal(firstByLocal(container, 'spPr'), 'xfrm') ||
    firstDirect(container, 'xfrm') ||
    firstByLocal(container, 'xfrm')
  if (!xfrm) return null
  const off = firstDirect(xfrm, 'off') || firstByLocal(xfrm, 'off')
  const ext = firstDirect(xfrm, 'ext') || firstByLocal(xfrm, 'ext')
  const cx = readEmu(ext, 'cx')
  const cy = readEmu(ext, 'cy')
  if (cx <= 0 || cy <= 0) return null
  return { x: readEmu(off, 'x'), y: readEmu(off, 'y'), cx, cy }
}

const readPh = (sp: Element): { type: string | null; idx: string | null } => {
  const ph = firstByLocal(sp, 'ph')
  if (!ph) return { type: null, idx: null }
  return {
    type: ph.getAttribute('type'),
    idx: ph.getAttribute('idx'),
  }
}

const classifyPlaceholder = (type: string | null, idx: string | null): TextShapeExport['kind'] => {
  const t = (type || '').toLowerCase()
  if (t === 'title' || t === 'ctrtitle') return 'title'
  if (t === 'body' || t === 'subtitle' || t === 'obj') return 'body'
  if (!t && idx) return 'body'
  return 'other'
}

const parseThemeFonts = (themeDoc: Document | null): ThemeFonts => {
  const fonts: ThemeFonts = { major: 'Calibri Light', minor: 'Calibri' }
  if (!themeDoc) return fonts
  const scheme = firstByLocal(themeDoc, 'fontScheme')
  const major = firstByLocal(scheme, 'majorFont')
  const minor = firstByLocal(scheme, 'minorFont')
  const majorLatin = firstDirect(major, 'latin') || firstByLocal(major, 'latin')
  const minorLatin = firstDirect(minor, 'latin') || firstByLocal(minor, 'latin')
  if (majorLatin?.getAttribute('typeface')) fonts.major = majorLatin.getAttribute('typeface')!
  if (minorLatin?.getAttribute('typeface')) fonts.minor = minorLatin.getAttribute('typeface')!
  return fonts
}

const resolveTypeface = (raw: string | null | undefined, themeFonts: ThemeFonts): string | undefined => {
  if (!raw) return undefined
  if (raw === '+mj-lt' || raw === '+mj-ea' || raw === '+mj-cs') return themeFonts.major
  if (raw === '+mn-lt' || raw === '+mn-ea' || raw === '+mn-cs') return themeFonts.minor
  return raw
}

const parseLevelStyle = (
  lvlEl: Element | null,
  scheme: ThemeColorMap,
  themeFonts: ThemeFonts
): LevelStyle => {
  if (!lvlEl) return {}
  const def = firstDirect(lvlEl, 'defRPr') || firstByLocal(lvlEl, 'defRPr')
  const solid = def ? firstDirect(def, 'solidFill') || firstByLocal(def, 'solidFill') : null
  const latin = def ? firstDirect(def, 'latin') || firstByLocal(def, 'latin') : null
  const buNone = !!firstDirect(lvlEl, 'buNone') || !!firstByLocal(lvlEl, 'buNone')
  const buChar = firstDirect(lvlEl, 'buChar') || firstByLocal(lvlEl, 'buChar')
  const sz = parseInt(def?.getAttribute('sz') || '', 10)

  return {
    marL: (() => {
      const v = parseInt(lvlEl.getAttribute('marL') || '', 10)
      return Number.isFinite(v) ? v : undefined
    })(),
    indent: (() => {
      const v = parseInt(lvlEl.getAttribute('indent') || '', 10)
      return Number.isFinite(v) ? v : undefined
    })(),
    algn: lvlEl.getAttribute('algn') || undefined,
    bulletNone: buNone,
    bullet: buNone ? null : buChar?.getAttribute('char') || undefined,
    sz: Number.isFinite(sz) && sz > 0 ? sz : undefined,
    bold: def?.getAttribute('b') === '1' || def?.getAttribute('b') === 'true',
    italic: def?.getAttribute('i') === '1' || def?.getAttribute('i') === 'true',
    underline: !!def?.getAttribute('u') && def.getAttribute('u') !== 'none',
    color: solid ? resolveColorElement(solid, scheme) || undefined : undefined,
    fontFamily: resolveTypeface(latin?.getAttribute('typeface'), themeFonts),
  }
}

const levelName = (level: number): string => (level <= 0 ? 'lvl1pPr' : `lvl${level + 1}pPr`)

const getLstLevel = (lstStyle: Element | null, level: number): Element | null => {
  if (!lstStyle) return null
  const name = levelName(level)
  return directChildrenByLocal(lstStyle, name)[0] || elementsByLocal(lstStyle, name)[0] || null
}

const getTxStyleLevel = (
  masterDoc: Document | null,
  kind: TextShapeExport['kind'],
  level: number
): Element | null => {
  if (!masterDoc) return null
  const txStyles = firstByLocal(masterDoc, 'txStyles')
  const styleName = kind === 'title' ? 'titleStyle' : kind === 'body' ? 'bodyStyle' : 'otherStyle'
  const style = firstByLocal(txStyles, styleName)
  return getLstLevel(style, level) || firstDirect(style, 'defPPr') || firstByLocal(style, 'defPPr')
}

const findPlaceholderShape = (
  doc: Document | null,
  type: string | null,
  idx: string | null
): Element | null => {
  if (!doc) return null
  const shapes = elementsByLocal(doc, 'sp')
  const t = (type || '').toLowerCase()

  if (t) {
    const byType = shapes.find((sp) => (readPh(sp).type || '').toLowerCase() === t)
    if (byType) return byType
  }
  if (idx != null) {
    const byIdx = shapes.find((sp) => readPh(sp).idx === idx)
    if (byIdx) return byIdx
  }
  return null
}

const mergeLevel = (base: LevelStyle, over: LevelStyle): LevelStyle => ({
  marL: over.marL ?? base.marL,
  indent: over.indent ?? base.indent,
  algn: over.algn ?? base.algn,
  bullet: over.bulletNone ? null : over.bullet !== undefined ? over.bullet : base.bullet,
  bulletNone: over.bulletNone ?? base.bulletNone,
  sz: over.sz ?? base.sz,
  bold: over.bold || base.bold,
  italic: over.italic || base.italic,
  underline: over.underline || base.underline,
  color: over.color ?? base.color,
  fontFamily: over.fontFamily ?? base.fontFamily,
})

const resolvePadding = (txBody: Element | null) => {
  const bodyPr = firstDirect(txBody, 'bodyPr') || firstByLocal(txBody, 'bodyPr')
  return {
    left: emuAttrToPx(bodyPr?.getAttribute('lIns'), emuToPx(91440)),
    top: emuAttrToPx(bodyPr?.getAttribute('tIns'), emuToPx(45720)),
    right: emuAttrToPx(bodyPr?.getAttribute('rIns'), emuToPx(91440)),
    bottom: emuAttrToPx(bodyPr?.getAttribute('bIns'), emuToPx(45720)),
    anchor: bodyPr?.getAttribute('anchor') || null,
  }
}

const paragraphHasText = (p: Element): boolean =>
  directChildrenByLocal(p, 'r').some((r) => (firstDirect(r, 't')?.textContent || '').length > 0) ||
  elementsByLocal(p, 't').some((t) => (t.textContent || '').trim().length > 0)

const collectParagraphs = (txBody: Element): Element[] =>
  Array.from(txBody.children).filter((el) => localName(el) === 'p') as Element[]

const parseRun = (
  r: Element,
  level: LevelStyle,
  scheme: ThemeColorMap,
  themeFonts: ThemeFonts
): TextRun | null => {
  const textNode = firstDirect(r, 't') || firstByLocal(r, 't')
  if (!textNode) return null
  const text = textNode.textContent ?? ''

  const rPr = firstDirect(r, 'rPr') || firstByLocal(r, 'rPr')
  const solid = rPr ? firstDirect(rPr, 'solidFill') || firstByLocal(rPr, 'solidFill') : null
  const latin = rPr ? firstDirect(rPr, 'latin') || firstByLocal(rPr, 'latin') : null
  const szAttr = parseInt(rPr?.getAttribute('sz') || '', 10)
  const sz = Number.isFinite(szAttr) && szAttr > 0 ? szAttr : level.sz || 1800

  return {
    text,
    fontSizePx: Math.max(8, szToPx(sz)),
    bold: rPr?.getAttribute('b') === '1' || rPr?.getAttribute('b') === 'true' || !!level.bold,
    italic: rPr?.getAttribute('i') === '1' || rPr?.getAttribute('i') === 'true' || !!level.italic,
    underline:
      (!!rPr?.getAttribute('u') && rPr.getAttribute('u') !== 'none') || !!level.underline,
    color: (solid ? resolveColorElement(solid, scheme) : null) || level.color || '#000000',
    fontFamily:
      resolveTypeface(latin?.getAttribute('typeface'), themeFonts) ||
      level.fontFamily ||
      themeFonts.minor,
  }
}

const cssAlign = (algn: string | undefined): string => {
  switch ((algn || 'l').toLowerCase()) {
    case 'ctr':
      return 'center'
    case 'r':
      return 'right'
    case 'just':
      return 'justify'
    default:
      return 'left'
  }
}

const cssVerticalAlign = (anchor: string): string => {
  switch (anchor) {
    case 'ctr':
      return 'middle'
    case 'b':
      return 'bottom'
    default:
      return 'top'
  }
}

export const extractSlideTextShapes = async (
  zip: JSZip,
  slidePath: string,
  slideDoc: Document,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<TextShapeExport[]> => {
  const slideRels = await loadRelationships(zip, slidePath)
  const layoutPath = findRelByType(slideRels, REL_LAYOUT)
  const layoutDoc = layoutPath ? await loadXml(zip, layoutPath) : null
  const layoutRels = layoutPath ? await loadRelationships(zip, layoutPath) : new Map()
  const masterPath = findRelByType(layoutRels, REL_MASTER)
  const masterDoc = masterPath ? await loadXml(zip, masterPath) : null
  const masterRels = masterPath ? await loadRelationships(zip, masterPath) : new Map()
  const themePath = findRelByType(masterRels, REL_THEME)
  const themeDoc = themePath ? await loadXml(zip, themePath) : null

  let scheme: ThemeColorMap = {}
  try {
    scheme = await loadSlideColorScheme(zip, slidePath)
  } catch (error) {
    console.warn('Text theme colours unavailable; using defaults', error)
  }
  const themeFonts = parseThemeFonts(themeDoc)

  const results: TextShapeExport[] = []
  const shapes = elementsByLocal(slideDoc, 'sp')

  for (let i = 0; i < shapes.length; i++) {
    const sp = shapes[i]
    const txBody = firstDirect(sp, 'txBody') || firstByLocal(sp, 'txBody')
    if (!txBody) continue

    const paras = collectParagraphs(txBody).filter(paragraphHasText)
    if (paras.length === 0) continue

    const ph = readPh(sp)
    const phType = (ph.type || '').toLowerCase()
    if (phType === 'dt' || phType === 'ftr' || phType === 'sldnum') continue

    const kind = classifyPlaceholder(ph.type, ph.idx)
    const layoutPh = findPlaceholderShape(layoutDoc, ph.type, ph.idx)
    const masterPh =
      findPlaceholderShape(masterDoc, ph.type, ph.idx) ||
      (kind === 'title'
        ? findPlaceholderShape(masterDoc, 'title', null)
        : kind === 'body'
          ? findPlaceholderShape(masterDoc, 'body', ph.idx || '1')
          : null)

    const box =
      readXfrmBox(sp) ||
      readXfrmBox(layoutPh) ||
      readXfrmBox(masterPh) || {
        x: Math.round(slideSize.widthEmu * 0.05),
        y: Math.round(slideSize.heightEmu * 0.05),
        cx: Math.round(slideSize.widthEmu * 0.9),
        cy: Math.round(slideSize.heightEmu * 0.2),
      }

    const layoutTx = layoutPh ? firstDirect(layoutPh, 'txBody') || firstByLocal(layoutPh, 'txBody') : null
    const masterTx = masterPh ? firstDirect(masterPh, 'txBody') || firstByLocal(masterPh, 'txBody') : null
    const hasBodyPr = !!(firstDirect(txBody, 'bodyPr') || firstByLocal(txBody, 'bodyPr'))
    const slidePad = resolvePadding(txBody)
    const layoutPad = resolvePadding(layoutTx)
    const masterPad = resolvePadding(masterTx)
    const padding = {
      left: hasBodyPr ? slidePad.left : layoutPad.left || masterPad.left || slidePad.left,
      top: hasBodyPr ? slidePad.top : layoutPad.top || masterPad.top || slidePad.top,
      right: hasBodyPr ? slidePad.right : layoutPad.right || masterPad.right || slidePad.right,
      bottom: hasBodyPr ? slidePad.bottom : layoutPad.bottom || masterPad.bottom || slidePad.bottom,
      anchor: slidePad.anchor || layoutPad.anchor || masterPad.anchor || 't',
    }

    const shapeLst = firstDirect(txBody, 'lstStyle') || firstByLocal(txBody, 'lstStyle')
    const layoutLst = layoutTx
      ? firstDirect(layoutTx, 'lstStyle') || firstByLocal(layoutTx, 'lstStyle')
      : null
    const masterLst = masterTx
      ? firstDirect(masterTx, 'lstStyle') || firstByLocal(masterTx, 'lstStyle')
      : null

    const paragraphs: TextParagraph[] = []
    for (const p of paras) {
      const pPr = firstDirect(p, 'pPr') || firstByLocal(p, 'pPr')
      const lvl = Math.max(0, parseInt(pPr?.getAttribute('lvl') || '0', 10) || 0)

      let levelStyle = parseLevelStyle(getTxStyleLevel(masterDoc, kind, lvl), scheme, themeFonts)
      levelStyle = mergeLevel(levelStyle, parseLevelStyle(getLstLevel(masterLst, lvl), scheme, themeFonts))
      levelStyle = mergeLevel(levelStyle, parseLevelStyle(getLstLevel(layoutLst, lvl), scheme, themeFonts))
      levelStyle = mergeLevel(levelStyle, parseLevelStyle(getLstLevel(shapeLst, lvl), scheme, themeFonts))
      levelStyle = mergeLevel(levelStyle, parseLevelStyle(pPr, scheme, themeFonts))

      if (pPr && (firstDirect(pPr, 'buNone') || firstByLocal(pPr, 'buNone'))) {
        levelStyle = { ...levelStyle, bullet: null, bulletNone: true }
      } else if (pPr) {
        const buChar = firstDirect(pPr, 'buChar') || firstByLocal(pPr, 'buChar')
        if (buChar?.getAttribute('char')) {
          levelStyle = { ...levelStyle, bullet: buChar.getAttribute('char'), bulletNone: false }
        }
      }

      if (kind === 'title') {
        levelStyle = { ...levelStyle, bullet: null, bulletNone: true }
      }

      const defRPr = pPr ? firstDirect(pPr, 'defRPr') || firstByLocal(pPr, 'defRPr') : null
      if (defRPr) {
        const solid = firstDirect(defRPr, 'solidFill') || firstByLocal(defRPr, 'solidFill')
        const latin = firstDirect(defRPr, 'latin') || firstByLocal(defRPr, 'latin')
        const sz = parseInt(defRPr.getAttribute('sz') || '', 10)
        levelStyle = mergeLevel(levelStyle, {
          sz: Number.isFinite(sz) && sz > 0 ? sz : undefined,
          bold: defRPr.getAttribute('b') === '1' || defRPr.getAttribute('b') === 'true',
          italic: defRPr.getAttribute('i') === '1' || defRPr.getAttribute('i') === 'true',
          underline: !!defRPr.getAttribute('u') && defRPr.getAttribute('u') !== 'none',
          color: solid ? resolveColorElement(solid, scheme) || undefined : undefined,
          fontFamily: resolveTypeface(latin?.getAttribute('typeface'), themeFonts),
        })
      }

      let runs = directChildrenByLocal(p, 'r')
        .map((r) => parseRun(r, levelStyle, scheme, themeFonts))
        .filter((r): r is TextRun => !!r)

      if (runs.length === 0) {
        const raw = (p.textContent || '').replace(/\s+/g, ' ').trim()
        if (!raw) continue
        runs = [
          {
            text: raw,
            fontSizePx: szToPx(levelStyle.sz || 1800),
            bold: !!levelStyle.bold,
            italic: !!levelStyle.italic,
            underline: !!levelStyle.underline,
            color: levelStyle.color || '#000000',
            fontFamily: levelStyle.fontFamily || themeFonts.minor,
          },
        ]
      }

      paragraphs.push({
        level: lvl,
        align: cssAlign(levelStyle.algn || pPr?.getAttribute('algn') || undefined),
        bullet: levelStyle.bulletNone ? null : levelStyle.bullet || null,
        marginLeftPx: emuToPx(levelStyle.marL || 0),
        indentPx: emuToPx(levelStyle.indent || 0),
        runs,
      })
    }

    if (paragraphs.length === 0) continue

    results.push({
      id: `text-${i + 1}`,
      kind,
      leftPx: Math.max(0, emuToPx(box.x)),
      topPx: Math.max(0, emuToPx(box.y)),
      widthPx: Math.max(20, emuToPx(box.cx)),
      heightPx: Math.max(20, emuToPx(box.cy)),
      padding: {
        left: padding.left,
        top: padding.top,
        right: padding.right,
        bottom: padding.bottom,
      },
      verticalAlign: cssVerticalAlign(padding.anchor),
      paragraphs,
    })
  }

  console.log(`Extracted ${results.length} text shape(s) on ${slidePath}`)
  return results
}

export const renderTextShapesHtml = (shapes: TextShapeExport[]): string => {
  if (shapes.length === 0) return ''

  return shapes
    .map((shape) => {
      const justify =
        shape.verticalAlign === 'middle'
          ? 'center'
          : shape.verticalAlign === 'bottom'
            ? 'flex-end'
            : 'flex-start'
      const box = `position:absolute;left:${shape.leftPx.toFixed(1)}px;top:${shape.topPx.toFixed(1)}px;width:${shape.widthPx.toFixed(1)}px;height:${shape.heightPx.toFixed(1)}px;padding:${shape.padding.top.toFixed(1)}px ${shape.padding.right.toFixed(1)}px ${shape.padding.bottom.toFixed(1)}px ${shape.padding.left.toFixed(1)}px;box-sizing:border-box;display:flex;flex-direction:column;justify-content:${justify};overflow:hidden;z-index:3;`

      const paras = shape.paragraphs
        .map((p) => {
          const runs = p.runs
            .map((r) => {
              const styles = [
                `font-size:${r.fontSizePx.toFixed(1)}px`,
                `color:${r.color}`,
                `font-family:${JSON.stringify(r.fontFamily)}, Calibri, sans-serif`,
                r.bold ? 'font-weight:700' : 'font-weight:400',
                r.italic ? 'font-style:italic' : '',
                r.underline ? 'text-decoration:underline' : '',
              ]
                .filter(Boolean)
                .join(';')
              return `<span style="${styles}">${escapeHtml(r.text)}</span>`
            })
            .join('')

          const bullet =
            p.bullet != null
              ? `<span class="slide-text-bullet" aria-hidden="true" style="font-size:${(p.runs[0]?.fontSizePx || 16).toFixed(1)}px;color:${p.runs[0]?.color || '#000000'}">${escapeHtml(p.bullet)}</span>`
              : ''

          return `<div class="slide-text-p" style="text-align:${p.align};padding-left:${Math.max(0, p.marginLeftPx).toFixed(1)}px;text-indent:${p.indentPx.toFixed(1)}px;">${bullet}<span class="slide-text-runs">${runs}</span></div>`
        })
        .join('')

      return `
    <div class="slide-text-shape slide-text-${shape.kind}" data-text="${shape.id}" style="${box}">
      ${paras}
    </div>`
    })
    .join('\n')
}

export const textShapesCss = `
      .slide-text-shape {
        pointer-events: none;
      }
      .slide-text-p {
        position: relative;
        margin: 0 0 0.15em;
        line-height: 1.2;
        white-space: pre-wrap;
        word-break: break-word;
      }
      .slide-text-bullet {
        position: absolute;
        left: 0;
        width: 1em;
        text-align: center;
        font-family: Arial, sans-serif;
      }
      .slide-text-runs {
        display: inline;
      }
`
