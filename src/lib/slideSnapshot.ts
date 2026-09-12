import { toPng } from 'html-to-image'
import html2canvas from 'html2canvas'

export const SNAPSHOT_HIDE_3D =
  '.slide-3d-model, .slide-3d-poster, .slide-3d-poster-orbit, model-viewer'

const UNSUPPORTED_COLOR = /oklch\([^)]*\)|oklab\([^)]*\)|lab\([^)]*\)|lch\([^)]*\)|color\([^)]*\)/gi

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

const stripHeavyNodes = (root: ParentNode, hideSelectors?: string) => {
  root.querySelectorAll('script, link, iframe, object, embed, model-viewer').forEach((node) => node.remove())
  if (hideSelectors) {
    root.querySelectorAll(hideSelectors).forEach((node) => node.remove())
  }
  root.querySelectorAll('[src], [poster], [style]').forEach((node) => {
    if (!(node instanceof Element)) return
    for (const attr of ['src', 'poster', 'style']) {
      const value = node.getAttribute(attr)
      if (value && /data:model\//i.test(value)) {
        if (attr === 'style') {
          node.setAttribute(attr, value.replace(/url\(\s*['"]?data:model\/[^)]+\)/gi, 'none'))
        } else {
          node.removeAttribute(attr)
        }
      }
    }
  })
}

const sanitizeSnapshotHtml = (html: string, hideSelectors?: string) => {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  stripHeavyNodes(doc, hideSelectors)
  doc.querySelectorAll('style').forEach((style) => {
    style.textContent = (style.textContent || '').replace(UNSUPPORTED_COLOR, '#64748b')
  })
  doc.querySelectorAll('[style]').forEach((node) => {
    const style = node.getAttribute('style')
    if (style && UNSUPPORTED_COLOR.test(style)) {
      node.setAttribute('style', style.replace(UNSUPPORTED_COLOR, '#64748b'))
    }
  })
  return doc.body.innerHTML
}

const snapshotWithHtmlToImage = async (host: HTMLElement, widthPx: number, heightPx: number) =>
  toPng(host, {
    width: widthPx,
    height: heightPx,
    pixelRatio: 1,
    cacheBust: false,
    skipFonts: true,
  })

const snapshotWithHtml2Canvas = async (host: HTMLElement, widthPx: number, heightPx: number) => {
  const canvas = await html2canvas(host, {
    width: widthPx,
    height: heightPx,
    scale: 2,
    useCORS: true,
    allowTaint: true,
    backgroundColor: null,
    logging: false,
    windowWidth: widthPx,
    windowHeight: heightPx,
  })
  return canvas.toDataURL('image/png')
}

export const snapshotHtmlSlide = async (
  html: string,
  widthPx: number,
  heightPx: number,
  options: { hideSelectors?: string } = {}
): Promise<string> => {
  const host = document.createElement('div')
  host.setAttribute('data-slide-snapshot', 'true')
  host.style.cssText = [
    'position:fixed',
    'left:0',
    'top:0',
    `width:${widthPx}px`,
    `height:${heightPx}px`,
    'overflow:hidden',
    'z-index:0',
    'pointer-events:none',
    'box-sizing:border-box',
  ].join(';')
  host.innerHTML = sanitizeSnapshotHtml(html, options.hideSelectors)
  stripHeavyNodes(host, options.hideSelectors)
  document.body.appendChild(host)

  const target =
    (host.querySelector('.imported-slide') as HTMLElement | null) ??
    (host.firstElementChild as HTMLElement | null) ??
    host
  target.style.position = 'absolute'
  target.style.left = '0'
  target.style.top = '0'
  target.style.right = '0'
  target.style.bottom = '0'
  target.style.width = `${widthPx}px`
  target.style.height = `${heightPx}px`
  target.style.maxWidth = 'none'
  target.style.maxHeight = 'none'
  target.style.transform = 'none'
  target.style.margin = '0'
  target.style.boxSizing = 'border-box'

  try {
    await waitForLayout(host)
    try {
      return await snapshotWithHtmlToImage(host, widthPx, heightPx)
    } catch (firstError) {
      console.warn('html-to-image snapshot failed, trying html2canvas:', firstError)
      return await snapshotWithHtml2Canvas(host, widthPx, heightPx)
    }
  } finally {
    host.remove()
  }
}
