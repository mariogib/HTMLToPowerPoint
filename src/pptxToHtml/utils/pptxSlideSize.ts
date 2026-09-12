import type JSZip from 'jszip'

/** Office EMUs per CSS pixel at 96 DPI (914400 EMU/inch ÷ 96). */
export const EMU_PER_PX = 9525

export interface SlideSize {
  widthEmu: number
  heightEmu: number
  widthPx: number
  heightPx: number
}

export const DEFAULT_SLIDE_SIZE: SlideSize = {
  widthEmu: 12192000, // 16:9 widescreen (13.333" × 7.5")
  heightEmu: 6858000,
  widthPx: Math.round(12192000 / EMU_PER_PX), // 1280
  heightPx: Math.round(6858000 / EMU_PER_PX), // 720
}

const localName = (el: Element | null | undefined): string => {
  if (!el) return ''
  return el.localName || el.tagName.replace(/^.*:/, '')
}

const firstByLocal = (root: Document | Element, name: string): Element | null => {
  const scope = root instanceof Document ? root.documentElement : root
  if (!scope) return null
  return (
    (Array.from(scope.getElementsByTagName('*')).find((el) => localName(el) === name) as Element) ||
    null
  )
}

const getZipEntry = (zip: JSZip, path: string) => {
  const candidates = [path, path.replace(/^\//, ''), `/${path.replace(/^\//, '')}`]
  for (const candidate of candidates) {
    const entry = zip.file(candidate) || zip.files[candidate]
    if (entry && !entry.dir) return entry
  }
  return null
}

export const emuToPx = (emu: number): number => emu / EMU_PER_PX

export const roundPx = (emu: number): number => Math.round(emuToPx(emu))

/**
 * Read the presentation slide size from ppt/presentation.xml (p:sldSz).
 * Falls back to widescreen 16:9 if missing.
 */
export const readPresentationSlideSize = async (zip: JSZip): Promise<SlideSize> => {
  const entry = getZipEntry(zip, 'ppt/presentation.xml')
  if (!entry) {
    console.warn('presentation.xml missing; using default widescreen slide size')
    return { ...DEFAULT_SLIDE_SIZE }
  }

  try {
    const text = await entry.async('text')
    const doc = new DOMParser().parseFromString(text, 'text/xml')
    const sldSz = firstByLocal(doc, 'sldSz')
    const cx = parseInt(sldSz?.getAttribute('cx') || '', 10)
    const cy = parseInt(sldSz?.getAttribute('cy') || '', 10)

    if (!sldSz || !Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) {
      console.warn('Invalid sldSz; using default widescreen slide size')
      return { ...DEFAULT_SLIDE_SIZE }
    }

    const size: SlideSize = {
      widthEmu: cx,
      heightEmu: cy,
      widthPx: roundPx(cx),
      heightPx: roundPx(cy),
    }
    console.log('Presentation slide size:', size)
    return size
  } catch (error) {
    console.error('Failed to read slide size:', error)
    return { ...DEFAULT_SLIDE_SIZE }
  }
}
