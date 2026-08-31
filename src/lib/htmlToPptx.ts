import PptxGenJS from 'pptxgenjs'

export type HtmlToPptxOptions = {
  slideWidthIn?: number
  slideHeightIn?: number
  background?: string
  embedImages?: boolean
  baseUrl?: string
}

const DEFAULT_SLIDE_WIDTH_IN = 13.333
const DEFAULT_SLIDE_HEIGHT_IN = 7.5

type TextStyle = {
  fontSize?: number
  bold?: boolean
  color?: string
  italic?: boolean
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

const getComputedStyles = (element: HTMLElement): TextStyle => {
  const computed = window.getComputedStyle(element)
  const fontSize = parseFloat(computed.fontSize)
  
  return {
    fontSize: fontSize ? Math.round(fontSize * 0.75) : undefined,
    bold: computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600,
    italic: computed.fontStyle === 'italic',
    color: computed.color ? parseColor(computed.color) : undefined,
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
      canvas.width = width
      canvas.height = height
      const ctx = canvas.getContext('2d')
      if (ctx) {
        ctx.drawImage(img, 0, 0, width, height)
        resolve(canvas.toDataURL('image/png'))
      } else {
        reject(new Error('Could not get canvas context'))
      }
    }
    img.onerror = () => reject(new Error('Failed to load SVG'))
    img.src = svgDataUrl
  })
}

const processNode = async (
  node: Node,
  slide: PptxGenJS.Slide,
  position: { x: number; y: number },
  baseUrl: string,
  slideWidth: number
): Promise<number> => {
  let currentY = position.y

  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent?.trim()
    if (text) {
      const parent = node.parentElement
      if (parent) {
        const styles = getComputedStyles(parent)
        slide.addText(text, {
          x: position.x,
          y: currentY,
          w: slideWidth - position.x - 0.5,
          fontSize: styles.fontSize || 14,
          bold: styles.bold,
          italic: styles.italic,
          color: styles.color || '000000',
        })
        currentY += ((styles.fontSize || 14) / 72) * 1.5
      }
    }
    return currentY
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return currentY
  }

  const element = node as HTMLElement
  const tagName = element.tagName.toLowerCase()

  if (tagName === 'h1' || tagName === 'h2' || tagName === 'h3') {
    const text = element.textContent?.trim()
    if (text) {
      const styles = getComputedStyles(element)
      const fontSize = tagName === 'h1' ? 32 : tagName === 'h2' ? 24 : 18
      slide.addText(text, {
        x: position.x,
        y: currentY,
        w: slideWidth - position.x - 0.5,
        fontSize,
        bold: true,
        color: styles.color || '000000',
      })
      currentY += (fontSize / 72) * 1.5
    }
  } else if (tagName === 'p') {
    const text = element.textContent?.trim()
    if (text) {
      const styles = getComputedStyles(element)
      slide.addText(text, {
        x: position.x,
        y: currentY,
        w: slideWidth - position.x - 0.5,
        fontSize: styles.fontSize || 14,
        bold: styles.bold,
        italic: styles.italic,
        color: styles.color || '000000',
      })
      currentY += ((styles.fontSize || 14) / 72) * 2
    }
  } else if (tagName === 'img') {
    const src = element.getAttribute('src')
    if (src) {
      let dataUrl = src.startsWith('data:') ? src : await fetchAsDataUrl(src, baseUrl)
      
      if (dataUrl) {
        const width = element.getAttribute('width')
        const height = element.getAttribute('height')
        const widthPx = width ? parseFloat(width) : 140
        const heightPx = height ? parseFloat(height) : 140
        const widthIn = widthPx / 96
        const heightIn = heightPx / 96
        
        // Convert SVG data URLs to PNG
        if (dataUrl.startsWith('data:image/svg+xml')) {
          try {
            dataUrl = await svgToPng(dataUrl, widthPx, heightPx)
          } catch (error) {
            console.warn('Failed to convert SVG to PNG:', error)
            return currentY
          }
        }
        
        slide.addImage({
          data: dataUrl,
          x: position.x,
          y: currentY,
          w: widthIn,
          h: heightIn,
        })
        currentY += heightIn + 0.2
      }
    }
  } else if (tagName === 'ul' || tagName === 'ol') {
    const items = Array.from(element.querySelectorAll('li'))
    for (let i = 0; i < items.length; i++) {
      const li = items[i]
      const text = li.textContent?.trim()
      if (text) {
        const styles = getComputedStyles(li)
        const bullet = tagName === 'ul' ? '• ' : `${i + 1}. `
        slide.addText(bullet + text, {
          x: position.x + 0.3,
          y: currentY,
          w: slideWidth - position.x - 0.8,
          fontSize: styles.fontSize || 14,
          color: styles.color || '000000',
        })
        currentY += ((styles.fontSize || 14) / 72) * 1.5
      }
    }
    currentY += 0.2
  } else if (tagName === 'div') {
    for (const child of Array.from(element.childNodes)) {
      currentY = await processNode(child, slide, { x: position.x, y: currentY }, baseUrl, slideWidth)
    }
  } else {
    for (const child of Array.from(element.childNodes)) {
      currentY = await processNode(child, slide, { x: position.x, y: currentY }, baseUrl, slideWidth)
    }
  }

  return currentY
}

export const htmlToPptx = async (html: string, options: HtmlToPptxOptions = {}) => {
  const widthIn = options.slideWidthIn ?? DEFAULT_SLIDE_WIDTH_IN
  const heightIn = options.slideHeightIn ?? DEFAULT_SLIDE_HEIGHT_IN
  const background = options.background ?? '#ffffff'
  const baseUrl = options.baseUrl ?? window.location.href

  const container = document.createElement('div')
  container.style.position = 'absolute'
  container.style.left = '-10000px'
  container.style.top = '0'
  container.innerHTML = html
  document.body.appendChild(container)

  try {
    const pptx = new PptxGenJS()
    const layoutName = `CUSTOM_${widthIn}x${heightIn}`
    pptx.defineLayout({ name: layoutName, width: widthIn, height: heightIn })
    pptx.layout = layoutName

    const slide = pptx.addSlide()
    slide.background = { color: parseColor(background) }

    await processNode(container, slide, { x: 0.5, y: 0.5 }, baseUrl, widthIn)

    return (await pptx.write({ outputType: 'blob' })) as Blob
  } finally {
    document.body.removeChild(container)
  }
}
