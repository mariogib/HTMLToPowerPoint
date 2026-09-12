export type ImportedHtmlPage = {
  html: string
  backgroundColor?: string
}

export type ImportedHtmlDocument = {
  title?: string
  pages: ImportedHtmlPage[]
}

const SLIDE_SELECTOR = '.slide[data-slide], .slides-container > .slide, section.slide'
const MODEL_VIEWER_SRC = 'https://ajax.googleapis.com/ajax/libs/model-viewer/3.5.0/model-viewer.min.js'

export const isImportedSlideHtml = (html: string) =>
  /imported-slide|slide-text-layer|slide-media-layer|data-slide\s*=/.test(html)

export const ensureModelViewerLoaded = async () => {
  if (typeof document === 'undefined') return
  if (!document.querySelector(`script[data-imported-model-viewer]`)) {
    const script = document.createElement('script')
    script.type = 'module'
    script.src = MODEL_VIEWER_SRC
    script.dataset.importedModelViewer = 'true'
    document.head.appendChild(script)
  }
  if (typeof customElements === 'undefined') return
  if (customElements.get('model-viewer')) return
  await Promise.race([
    customElements.whenDefined('model-viewer'),
    new Promise<void>((resolve) => window.setTimeout(resolve, 20000)),
  ])
}

export const SHARED_SLIDE_CSS = `
.imported-slide {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  overflow: hidden;
  box-sizing: border-box;
  font-family: Calibri, Aptos, Arial, sans-serif;
  background-repeat: no-repeat;
  background-position: center center;
}
.imported-slide .imported-slide-bg {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  z-index: 0;
  pointer-events: none;
  background-repeat: no-repeat;
  background-position: center center;
}
.imported-slide .slide-text-layer,
.imported-slide .slide-media-layer {
  position: absolute;
  inset: 0;
}
.imported-slide .slide-text-layer { pointer-events: none; }
.imported-slide .slide-text-shape { pointer-events: none; }
.imported-slide .slide-text-p {
  position: relative;
  margin: 0 0 0.15em;
  line-height: 1.2;
  white-space: pre-wrap;
  word-break: break-word;
}
.imported-slide .slide-text-bullet {
  position: absolute;
  left: 0;
  width: 1em;
  text-align: center;
  font-family: Arial, sans-serif;
}
.imported-slide .slide-3d-model {
  overflow: hidden;
  pointer-events: auto;
}
.imported-slide .slide-3d-model model-viewer {
  display: block;
  width: 100%;
  height: 100%;
  background: transparent;
  --poster-color: transparent;
}
.imported-slide .slide-3d-poster,
.imported-slide .slide-3d-model img {
  width: 100%;
  height: 100%;
  object-fit: contain;
  display: block;
}
.imported-slide .smartart {
  box-sizing: border-box;
  overflow: hidden;
  position: absolute;
}
.imported-slide .smartart-shape {
  position: absolute;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 4px 6px;
  overflow: hidden;
}
.imported-slide .smartart-shape-text {
  font-size: inherit;
  font-weight: 600;
  line-height: 1.15;
  text-align: center;
  width: 100%;
  word-break: break-word;
}
`

const escapeAttr = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')

const normalizeDataUrl = (value: string) =>
  value.trim().replace(/^(data:[a-z0-9.+/-]+);\s*base64,/i, '$1;base64,')

const usableColor = (value: string) => {
  const color = value.trim().toLowerCase()
  if (!color || color === 'transparent' || color === 'inherit' || color === 'initial') {
    return undefined
  }
  if (color.startsWith('rgba') && /rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*0(?:\.0+)?\s*\)/.test(color)) {
    return undefined
  }
  return value.trim()
}

const readCssUrl = (style: string, property = 'background-image') => {
  const marker = style.search(new RegExp(`${property}\\s*:\\s*url\\(`, 'i'))
  if (marker < 0) return undefined
  let i = style.indexOf('url(', marker) + 4
  while (style[i] === ' ') i++
  const quote = style[i] === '"' || style[i] === "'" ? style[i] : ''
  if (quote) {
    const end = style.indexOf(quote, i + 1)
    if (end < 0) return undefined
    return normalizeDataUrl(style.slice(i + 1, end))
  }
  const end = style.indexOf(')', i)
  if (end < 0) return undefined
  return normalizeDataUrl(style.slice(i, end))
}

const readCssProp = (style: string, property: string) => {
  const match = style.match(new RegExp(`${property}\\s*:\\s*([^;]+)`, 'i'))
  if (!match?.[1] || /url\(/i.test(match[1])) return undefined
  return match[1].replace(/!important/gi, '').trim()
}

const collectGlbByViewerId = (doc: Document) => {
  const map = new Map<string, string>()
  for (const script of Array.from(doc.querySelectorAll('script'))) {
    const text = script.textContent || ''
    const idMatch = text.match(/getElementById\(\s*['"](mv-[^'"]+)['"]\s*\)/)
    const b64Match = text.match(/var b64 = ['"]([A-Za-z0-9+/=\s]+)['"]/)
    if (idMatch?.[1] && b64Match?.[1]) {
      map.set(idMatch[1], b64Match[1].replace(/\s+/g, ''))
    }
  }
  return map
}

const collectSlideBackgroundCss = (doc: Document) => {
  const map = new Map<string, string>()
  const css = Array.from(doc.querySelectorAll('style'))
    .map((node) => node.textContent || '')
    .join('\n')
  const rule = /\.slide\[data-slide=["'](\d+)["']\]\s*\{([^}]+)\}/g
  let match: RegExpExecArray | null
  while ((match = rule.exec(css))) {
    const number = match[1]
    const body = match[2]
    if (number && body && /background-image\s*:/i.test(body)) {
      map.set(number, body)
    }
  }
  return map
}

const restoreModelViewers = (root: HTMLElement, glbById: Map<string, string>) => {
  root.querySelectorAll('model-viewer').forEach((viewer) => {
    const poster = normalizeDataUrl(
      viewer.getAttribute('poster') || viewer.querySelector('img')?.getAttribute('src') || ''
    )
    if (poster) {
      viewer.setAttribute('poster', poster)
      const existing = viewer.querySelector('img')
      if (existing) {
        existing.setAttribute('src', poster)
      } else {
        const image = root.ownerDocument.createElement('img')
        image.className = 'slide-3d-poster'
        image.setAttribute('slot', 'poster')
        image.alt = viewer.getAttribute('alt') || ''
        image.src = poster
        viewer.appendChild(image)
      }
    }

    const id = viewer.getAttribute('id')
    const b64 = id ? glbById.get(id) : undefined
    if (b64) {
      viewer.setAttribute('src', `data:model/gltf-binary;base64,${b64}`)
    }

    if (!viewer.getAttribute('src') && !poster) {
      viewer.remove()
    }
  })
}

const readBackground = (inlineStyle: string, cssStyle: string) => {
  const imageUrl = readCssUrl(inlineStyle) || readCssUrl(cssStyle)
  const size =
    readCssProp(inlineStyle, 'background-size') || readCssProp(cssStyle, 'background-size')
  const repeat =
    readCssProp(inlineStyle, 'background-repeat') ||
    readCssProp(cssStyle, 'background-repeat') ||
    'no-repeat'
  const position =
    readCssProp(inlineStyle, 'background-position') ||
    readCssProp(cssStyle, 'background-position') ||
    'center center'
  const tiled = /repeat/i.test(repeat) && !/no-repeat/i.test(repeat)

  return {
    imageUrl,
    size: size || (tiled ? 'auto' : '100% 100%'),
    repeat,
    position,
    tiled,
  }
}

const backgroundLayer = (
  imageUrl: string,
  size: string,
  repeat: string,
  position: string
) =>
  `<div class="imported-slide-bg" style="background-image:url('${imageUrl.replace(/'/g, '%27')}');background-size:${escapeAttr(size)};background-repeat:${escapeAttr(repeat)};background-position:${escapeAttr(position)}"></div>`

const pageFromSlide = (
  slide: HTMLElement,
  glbById: Map<string, string>,
  backgroundCss: Map<string, string>
): ImportedHtmlPage => {
  const clone = slide.cloneNode(true) as HTMLElement
  restoreModelViewers(clone, glbById)
  clone.querySelectorAll('script').forEach((node) => node.remove())

  const dataSlide = slide.getAttribute('data-slide') || ''
  const inlineStyle = slide.getAttribute('style') || ''
  const cssStyle = dataSlide ? backgroundCss.get(dataSlide) || '' : ''
  const sourceStyle = [inlineStyle, cssStyle].filter(Boolean).join(';')
  const background = readBackground(inlineStyle, cssStyle)

  const backgroundColor = usableColor(
    readCssProp(sourceStyle, 'background-color') || slide.style.backgroundColor || ''
  )

  clone.querySelectorAll('img.slide-image').forEach((node) => {
    const positioned = /position\s*:\s*absolute/i.test(node.getAttribute('style') || '')
    if (!positioned) {
      node.remove()
    }
  })

  const gradient = /gradient\(/i.test(sourceStyle)
    ? sourceStyle.match(/background-image\s*:\s*[^;]+/i)?.[0]
    : undefined

  const wrapperStyle = [
    backgroundColor ? `background-color:${backgroundColor}` : '',
    gradient,
    'position:absolute',
    'inset:0',
    'width:100%',
    'height:100%',
    'overflow:hidden',
    'box-sizing:border-box',
  ]
    .filter(Boolean)
    .join(';')

  const layer = background.imageUrl
    ? backgroundLayer(background.imageUrl, background.size, background.repeat, background.position)
    : ''

  return {
    html: `<style>${SHARED_SLIDE_CSS}</style><div class="imported-slide" style="${escapeAttr(wrapperStyle)}">${layer}${clone.innerHTML}</div>`,
    backgroundColor,
  }
}

export const parseImportedHtml = (source: string): ImportedHtmlDocument => {
  if (!source.trim()) {
    return { pages: [{ html: source }] }
  }

  const doc = new DOMParser().parseFromString(source, 'text/html')
  const title =
    doc.querySelector('.presentation-header h1')?.textContent?.trim() ||
    doc.querySelector('title')?.textContent?.trim() ||
    undefined
  const glbById = collectGlbByViewerId(doc)
  const backgroundCss = collectSlideBackgroundCss(doc)

  const seen = new Set<Element>()
  const slides: HTMLElement[] = []
  for (const node of Array.from(doc.querySelectorAll(SLIDE_SELECTOR))) {
    if (seen.has(node) || !(node instanceof HTMLElement)) continue
    seen.add(node)
    slides.push(node)
  }

  if (slides.length === 0) {
    const imported = doc.querySelector('.imported-slide')
    if (imported instanceof HTMLElement) {
      const backgroundColor = usableColor(
        readCssProp(imported.getAttribute('style') || '', 'background-color') ||
          imported.style.backgroundColor ||
          ''
      )
      return { title, pages: [{ html: source, backgroundColor }] }
    }
    return { title, pages: [{ html: source }] }
  }

  return { title, pages: slides.map((slide) => pageFromSlide(slide, glbById, backgroundCss)) }
}

export const splitImportedHtml = (source: string): ImportedHtmlPage[] =>
  parseImportedHtml(source).pages
