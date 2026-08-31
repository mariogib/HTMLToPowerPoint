import PptxGenJS from 'pptxgenjs'
import { detectBulletStyle, pptxBulletOption, type BulletStyleId } from './bulletStyles'
import { containRect, loadImageNaturalSize } from './imageFit'
import { pxToInches, slideSizePx } from './slideLayout'

export type HtmlToPptxOptions = {
  slideWidthIn?: number
  slideHeightIn?: number
  background?: string
  embedImages?: boolean
  baseUrl?: string
}

const DEFAULT_SLIDE_WIDTH_IN = 13.333
const DEFAULT_SLIDE_HEIGHT_IN = 7.5

const CONTENT_TAGS = new Set([
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'p',
  'img',
  'ul',
  'ol',
  'pre',
  'blockquote',
])

type TextStyle = {
  fontSize?: number
  fontFace?: string
  bold?: boolean
  color?: string
  italic?: boolean
  align?: 'left' | 'center' | 'right'
}

export type MeasuredBlock = {
  kind: 'heading' | 'text' | 'image' | 'list'
  content: string
  x: number
  y: number
  width: number
  height: number
  styles: TextStyle
  listType?: 'ul' | 'ol'
  bulletIndentPt?: number
  bulletStyle?: BulletStyleId
  aspectRatio?: number
}

const parseColor = (color: string): string => {
  if (color.startsWith('#')) {
    return color.substring(1)
  }
  if (color.startsWith('rgb')) {
    const matches = color.match(/\d+/g)
    if (matches && matches.length >= 3) {
      const [r, g, b] = matches.map(Number)
      return ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0')
    }
  }
  return '000000'
}

const parseAlign = (align: string): TextStyle['align'] => {
  if (align === 'center' || align === 'right') {
    return align
  }
  return 'left'
}

const getComputedStyles = (element: HTMLElement): TextStyle => {
  const computed = window.getComputedStyle(element)
  const fontSize = parseFloat(computed.fontSize)
  const fontFace = computed.fontFamily.split(',')[0]?.replace(/['"]/g, '').trim()

  return {
    fontSize: fontSize ? Math.round(fontSize * 0.75) : undefined,
    fontFace: fontFace || 'Arial',
    bold: computed.fontWeight === 'bold' || parseInt(computed.fontWeight, 10) >= 600,
    italic: computed.fontStyle === 'italic',
    color: computed.color ? parseColor(computed.color) : undefined,
    align: parseAlign(computed.textAlign),
  }
}

const blockKind = (tagName: string): MeasuredBlock['kind'] => {
  if (tagName === 'img') return 'image'
  if (tagName === 'ul' || tagName === 'ol') return 'list'
  if (/^h[1-6]$/.test(tagName)) return 'heading'
  return 'text'
}

const blockContent = (element: HTMLElement): string => {
  const tagName = element.tagName.toLowerCase()
  if (tagName === 'img') {
    return (element as HTMLImageElement).getAttribute('src') || ''
  }
  if (tagName === 'ul' || tagName === 'ol') {
    return Array.from(element.children)
      .filter((child) => child.tagName.toLowerCase() === 'li')
      .map((li) => (li as HTMLElement).innerText.replace(/\u00a0/g, ' ').replace(/\r\n/g, '\n'))
      .join('\n')
  }
  return element.innerText.replace(/\u00a0/g, ' ').replace(/\r\n/g, '\n').replace(/^\n+|\n+$/g, '')
}

const collectContentElements = (root: HTMLElement): HTMLElement[] => {
  const leaves: HTMLElement[] = []

  const walk = (element: HTMLElement) => {
    const tagName = element.tagName.toLowerCase()
    if (tagName === 'script' || tagName === 'style' || tagName === 'br' || tagName === 'hr') {
      return
    }

    const computed = window.getComputedStyle(element)
    if (computed.display === 'none' || computed.visibility === 'hidden') {
      return
    }

    if (CONTENT_TAGS.has(tagName)) {
      leaves.push(element)
      return
    }

    const children = Array.from(element.children) as HTMLElement[]
    if (children.length === 0) {
      if (element.innerText.trim()) {
        leaves.push(element)
      }
      return
    }

    for (const child of children) {
      walk(child)
    }
  }

  const roots =
    root.children.length > 0 ? (Array.from(root.children) as HTMLElement[]) : [root]
  for (const child of roots) {
    walk(child)
  }

  return leaves
}

const waitForLayout = async (root: HTMLElement) => {
  const images = Array.from(root.querySelectorAll('img'))
  await Promise.all(
    images.map((img) =>
      img.complete
        ? Promise.resolve()
        : new Promise<void>((resolve) => {
            img.addEventListener('load', () => resolve(), { once: true })
            img.addEventListener('error', () => resolve(), { once: true })
          })
    )
  )

  if (document.fonts?.ready) {
    await document.fonts.ready
  }

  await new Promise<void>((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()))
  })
}

const createSlideContainer = (html: string, widthPx: number, heightPx: number) => {
  const container = document.createElement('div')
  container.style.position = 'absolute'
  container.style.left = '-10000px'
  container.style.top = '0'
  container.style.width = `${widthPx}px`
  container.style.height = `${heightPx}px`
  container.style.overflow = 'hidden'
  container.style.boxSizing = 'border-box'
  container.style.margin = '0'
  container.style.padding = '0'
  container.innerHTML = html
  document.body.appendChild(container)
  return container
}

export const measureHtmlBlocks = async (
  html: string,
  slideWidthIn = DEFAULT_SLIDE_WIDTH_IN,
  slideHeightIn = DEFAULT_SLIDE_HEIGHT_IN
): Promise<MeasuredBlock[]> => {
  const { width: widthPx, height: heightPx } = slideSizePx(slideWidthIn, slideHeightIn)
  const container = createSlideContainer(html, widthPx, heightPx)

  try {
    await waitForLayout(container)
    const origin = container.getBoundingClientRect()

    return collectContentElements(container)
      .map((element) => {
        const tagName = element.tagName.toLowerCase()
        const rect = element.getBoundingClientRect()
        const styles = getComputedStyles(element)
        const kind = blockKind(tagName)
        const content = blockContent(element)
        const computed = window.getComputedStyle(element)
        const paddingLeft = parseFloat(computed.paddingLeft) || 0
        const itemCount = kind === 'list' ? content.split('\n').filter(Boolean).length : 0
        const minListHeight =
          kind === 'list' ? itemCount * ((styles.fontSize || 14) * (96 / 72) * 1.25) : 0
        const image = kind === 'image' ? (element as HTMLImageElement) : null
        const aspectRatio =
          image && image.naturalWidth > 0 && image.naturalHeight > 0
            ? image.naturalWidth / image.naturalHeight
            : undefined

        return {
          kind,
          content,
          x: rect.left - origin.left,
          y: rect.top - origin.top,
          width: Math.max(rect.width, 1),
          height: Math.max(Math.max(rect.height, 1), minListHeight),
          styles,
          listType: tagName === 'ol' ? 'ol' : tagName === 'ul' ? 'ul' : undefined,
          bulletIndentPt:
            kind === 'list' ? Math.max(12, Math.round(paddingLeft * 0.75) || 14) : undefined,
          bulletStyle: kind === 'list' ? detectBulletStyle(element) : undefined,
          aspectRatio,
        } satisfies MeasuredBlock
      })
      .filter((block) => block.content && block.width >= 1 && block.height >= 1)
  } finally {
    document.body.removeChild(container)
  }
}

const fetchAsDataUrl = async (url: string, baseUrl: string) => {
  try {
    const resolved = new URL(url, baseUrl).toString()
    const response = await fetch(resolved, { mode: 'cors' })
    if (!response.ok) {
      throw new Error(`Failed to fetch image: ${response.status}`)
    }
    const blob = await response.blob()
    return new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.onloadend = () => resolve(reader.result as string)
      reader.onerror = () => reject(reader.error)
      reader.readAsDataURL(blob)
    })
  } catch {
    return null
  }
}

const svgToPng = async (svgDataUrl: string, width: number, height: number): Promise<string> => {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.round(width))
      canvas.height = Math.max(1, Math.round(height))
      const ctx = canvas.getContext('2d')
      if (ctx) {
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
        resolve(canvas.toDataURL('image/png'))
      } else {
        reject(new Error('Could not get canvas context'))
      }
    }
    img.onerror = () => reject(new Error('Failed to load SVG'))
    img.src = svgDataUrl
  })
}

const textOptions = (block: MeasuredBlock): PptxGenJS.TextPropsOptions => ({
  x: pxToInches(block.x),
  y: pxToInches(block.y),
  w: pxToInches(block.width),
  h: pxToInches(block.height),
  fontSize: block.styles.fontSize || (block.kind === 'heading' ? 24 : 14),
  fontFace: block.styles.fontFace || 'Arial',
  bold: block.kind === 'heading' ? true : block.styles.bold,
  italic: block.styles.italic,
  color: block.styles.color || '000000',
  align: block.styles.align,
  valign: 'top',
  wrap: true,
  margin: 0,
})

const addWrappedText = (
  slide: PptxGenJS.Slide,
  text: string,
  options: PptxGenJS.TextPropsOptions
) => {
  const lines = text.replace(/\r\n/g, '\n').split('\n')
  if (lines.length <= 1) {
    slide.addText(text, options)
    return
  }

  slide.addText(
    lines.map((line, index) => ({
      text: line,
      options: { breakLine: index < lines.length - 1 },
    })),
    options
  )
}

const paintBlocksOnSlide = async (
  pptxSlide: PptxGenJS.Slide,
  html: string,
  options: {
    widthIn: number
    heightIn: number
    background: string
    baseUrl: string
    embedImages: boolean
  }
) => {
  pptxSlide.background = { color: parseColor(options.background) }
  const blocks = await measureHtmlBlocks(html, options.widthIn, options.heightIn)

  for (const block of blocks) {
    if (block.kind === 'image') {
      if (!options.embedImages) {
        continue
      }

      let dataUrl = block.content.startsWith('data:')
        ? block.content
        : await fetchAsDataUrl(block.content, options.baseUrl)

      if (!dataUrl) {
        continue
      }

      let natural = { width: block.width, height: block.height }
      try {
        natural = await loadImageNaturalSize(dataUrl)
      } catch {
        if (block.aspectRatio) {
          natural = { width: block.aspectRatio, height: 1 }
        }
      }

      const fitted = containRect(
        block.x,
        block.y,
        block.width,
        block.height,
        natural.width,
        natural.height
      )

      if (dataUrl.startsWith('data:image/svg+xml')) {
        try {
          dataUrl = await svgToPng(dataUrl, natural.width, natural.height)
        } catch (error) {
          console.warn('Failed to convert SVG to PNG:', error)
          continue
        }
      }

      pptxSlide.addImage({
        data: dataUrl,
        x: pxToInches(fitted.x),
        y: pxToInches(fitted.y),
        w: pxToInches(fitted.width),
        h: pxToInches(fitted.height),
      })
      continue
    }

    if (block.kind === 'list') {
      const items = block.content
        .split('\n')
        .map((line) => line.trim())
        .filter(Boolean)

      if (items.length === 0) {
        continue
      }

      pptxSlide.addText(items.join('\n'), {
        ...textOptions(block),
        bullet: pptxBulletOption(block.bulletStyle ?? (block.listType === 'ol' ? 'number' : 'disc'), block.bulletIndentPt || 14),
        paraSpaceAfter: 0,
        paraSpaceBefore: 0,
        valign: 'top',
      })
      continue
    }

    addWrappedText(pptxSlide, block.content, textOptions(block))
  }
}

export const htmlToPptx = async (html: string | string[], options: HtmlToPptxOptions = {}) => {
  const widthIn = options.slideWidthIn ?? DEFAULT_SLIDE_WIDTH_IN
  const heightIn = options.slideHeightIn ?? DEFAULT_SLIDE_HEIGHT_IN
  const background = options.background ?? '#ffffff'
  const baseUrl = options.baseUrl ?? window.location.href
  const embedImages = options.embedImages ?? true
  const pages = (Array.isArray(html) ? html : [html]).filter((page) => page != null)

  const pptx = new PptxGenJS()
  const layoutName = `CUSTOM_${widthIn}x${heightIn}`
  pptx.defineLayout({ name: layoutName, width: widthIn, height: heightIn })
  pptx.layout = layoutName

  for (const pageHtml of pages) {
    await paintBlocksOnSlide(pptx.addSlide(), pageHtml, {
      widthIn,
      heightIn,
      background,
      baseUrl,
      embedImages,
    })
  }

  return (await pptx.write({ outputType: 'blob' })) as Blob
}
