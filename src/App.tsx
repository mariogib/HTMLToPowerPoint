import { useEffect, useMemo, useRef, useState, type RefObject } from 'react'
import { CanvasContextMenu } from './CanvasContextMenu'
import { PropertiesPanel } from './PropertiesPanel'
import { SlideRail } from './SlideRail'
import type { CanvasElement, DeckSlide } from './lib/canvasTypes'
import { htmlToPptx, measureHtmlBlocks } from './lib/htmlToPptx'
import { bulletPreview, getBulletStyle } from './lib/bulletStyles'
import type { BulletStyleId } from './lib/bulletStyles'
import { createDeckSlide, emptySlideHtml, moveById, reorderById } from './lib/deck'
import { captureDeck, cloneCanvasElement, cloneDeck, type DeckSnapshot } from './lib/deckSnapshot'
import { containRect, resizeProportional } from './lib/imageFit'
import { defaultColor, defaultFontSize, withAutoSize } from './lib/measureElement'
import { clientToSlidePx, slideSizePx } from './lib/slideLayout'
import './App.css'

const sampleSvg =
  '<svg xmlns="http://www.w3.org/2000/svg" width="140" height="140">' +
  '<rect width="140" height="140" fill="#4f46e5"/>' +
  '<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" ' +
  'font-family="Arial" font-size="18" fill="#ffffff">IMG</text>' +
  '</svg>'

const SAMPLE_IMAGE_DATA_URL = `data:image/svg+xml;utf8,${encodeURIComponent(sampleSvg)}`

const SAMPLE_HTML = `
<div style="font-family: Arial, sans-serif; padding: 28px; height: 100%; box-sizing: border-box; background: #ffffff;">
  <h1 style="margin: 0 0 12px; color: #1f2a44;">HTML to PowerPoint</h1>
  <p style="font-size: 18px; margin: 0 0 20px; color: #3e4c66;">
    Generate a PPTX from HTML and embed images automatically.
  </p>
  <div style="display: flex; gap: 20px; align-items: center;">
    <img src="${SAMPLE_IMAGE_DATA_URL}" width="140" height="140" alt="Sample" />
    <div style="font-size: 18px; color: #1f2a44;">
      <ul style="margin: 0; padding-left: 20px;">
        <li>Embedded HTML images</li>
        <li>Custom slide size</li>
        <li>Single-click export</li>
      </ul>
    </div>
  </div>
</div>
`.trim()

const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

const saveTextToChosenFile = async (contents: string, defaultName: string) => {
  const picker = (
    window as Window & {
      showSaveFilePicker?: (options: {
        suggestedName?: string
        types?: { description: string; accept: Record<string, string[]> }[]
      }) => Promise<{
        createWritable: () => Promise<{ write: (data: string) => Promise<void>; close: () => Promise<void> }>
      }>
    }
  ).showSaveFilePicker
  if (typeof picker === 'function') {
    try {
      const handle = await picker.call(window, {
        suggestedName: defaultName,
        types: [
          {
            description: 'HTML file',
            accept: { 'text/html': ['.html', '.htm'] },
          },
        ],
      })
      const writable = await handle.createWritable()
      await writable.write(contents)
      await writable.close()
      return
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') {
        return
      }
      throw error
    }
  }

  downloadBlob(new Blob([contents], { type: 'text/html;charset=utf-8' }), defaultName)
}

const ListEditor = ({
  content,
  onChange,
  bulletStyle,
}: {
  content: string
  onChange: (value: string) => void
  bulletStyle?: BulletStyleId
}) => {
  const items = content.length > 0 ? content.split('\n') : ['']
  const marker = getBulletStyle(bulletStyle)

  const updateItem = (index: number, value: string) => {
    const next = [...items]
    next[index] = value
    onChange(next.join('\n'))
  }

  return (
    <ul className="element-list">
      {items.map((item, index) => (
        <li key={index}>
          {marker.id !== 'none' && (
            <span className="element-list-bullet" aria-hidden="true">
              {bulletPreview(marker.id, index)}
            </span>
          )}
          <input
            className="element-list-input"
            value={item}
            aria-label={`List item ${index + 1}`}
            onMouseDown={(event) => event.stopPropagation()}
            onChange={(event) => updateItem(index, event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.preventDefault()
                const next = [...items]
                next.splice(index + 1, 0, '')
                onChange(next.join('\n'))
              }
              if (event.key === 'Backspace' && item === '' && items.length > 1) {
                event.preventDefault()
                onChange(items.filter((_, itemIndex) => itemIndex !== index).join('\n'))
              }
            }}
          />
        </li>
      ))}
    </ul>
  )
}

const useContainScale = (
  ref: RefObject<HTMLElement | null>,
  contentW: number,
  contentH: number,
  enabled: boolean
) => {
  const [scale, setScale] = useState(1)

  useEffect(() => {
    if (!enabled) return
    const el = ref.current
    if (!el) return

    const update = () => {
      const { width, height } = el.getBoundingClientRect()
      if (width < 1 || height < 1) return
      setScale(Math.min(width / contentW, height / contentH))
    }

    update()
    const observer = new ResizeObserver(update)
    observer.observe(el)
    return () => observer.disconnect()
  }, [contentW, contentH, enabled, ref])

  return scale
}

const escapeHtml = (value: string) =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

const toHtmlMultiline = (value: string) =>
  escapeHtml(value).replace(/\r\n/g, '\n').replace(/\n/g, '<br>')

const buildHtmlFromCanvas = (elements: CanvasElement[], background: string) => {
  const nodes = elements
    .map((el) => {
      const fontSize = el.fontSize ?? defaultFontSize(el.type)
      const color = el.color ?? defaultColor(el.type)
      const fontFamily = el.fontFamily ?? 'Arial'
      const weight = el.bold ?? el.type === 'heading' ? 'bold' : 'normal'
      const style = [
        `position: absolute`,
        `left: ${el.x}px`,
        `top: ${el.y}px`,
        `width: ${el.width}px`,
        `height: ${el.height}px`,
        `margin: 0`,
        `box-sizing: border-box`,
        `font-size: ${fontSize}px`,
        `font-family: ${fontFamily}, sans-serif`,
        `color: ${color}`,
        `font-weight: ${weight}`,
        el.italic ? 'font-style: italic' : '',
        `text-align: ${el.align ?? 'left'}`,
        el.type === 'image' ? '' : 'white-space: pre-wrap',
      ]
        .filter(Boolean)
        .join('; ')

      switch (el.type) {
        case 'heading':
          return `<h1 style="${style}">${toHtmlMultiline(el.content)}</h1>`
        case 'text':
          return `<p style="${style}">${toHtmlMultiline(el.content)}</p>`
        case 'image':
          return `<img src="${el.content}" style="${style}; object-fit: contain;" alt="Image" />`
        case 'list': {
          const items = el.content.split('\n')
          const bullet = getBulletStyle(el.bulletStyle)
          const tag = bullet.id === 'number' ? 'ol' : 'ul'
          return `<${tag} data-bullet="${bullet.id}" style="${style}; padding-left: 20px; list-style-type: ${bullet.cssType};">${items
            .map((item) => `<li>${toHtmlMultiline(item)}</li>`)
            .join('')}</${tag}>`
        }
        default:
          return ''
      }
    })
    .join('\n')

  return `<div style="font-family: Arial, sans-serif; width: 100%; height: 100%; box-sizing: border-box; background: ${background}; position: relative;">
${nodes}
</div>`
}

function App() {
  const [slides, setSlides] = useState<DeckSlide[]>(() => [
    { id: 'slide-1', html: SAMPLE_HTML, canvasElements: [] },
  ])
  const [activeSlideId, setActiveSlideId] = useState('slide-1')
  const [activeTab, setActiveTab] = useState<'editor' | 'preview' | 'html'>('editor')
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [propertiesId, setPropertiesId] = useState<string | null>(null)
  const [propertiesPos, setPropertiesPos] = useState({ x: 0, y: 0 })
  const [draggedElement, setDraggedElement] = useState<string | null>(null)
  const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 })
  const [resizingElement, setResizingElement] = useState<string | null>(null)
  const [resizeHandle, setResizeHandle] = useState<string | null>(null)
  const [resizeStart, setResizeStart] = useState({
    x: 0,
    y: 0,
    width: 0,
    height: 0,
    originX: 0,
    originY: 0,
  })
  const previewViewportRef = useRef<HTMLDivElement>(null)
  const builderViewportRef = useRef<HTMLDivElement>(null)
  const builderStageRef = useRef<HTMLDivElement>(null)
  const htmlFileInputRef = useRef<HTMLInputElement>(null)
  const [showSettings, setShowSettings] = useState(false)
  const [filename, setFilename] = useState('html-slide.pptx')
  const [slideWidth, setSlideWidth] = useState(13.333)
  const [slideHeight, setSlideHeight] = useState(7.5)
  const [background, setBackground] = useState('#ffffff')
  const [embedImages, setEmbedImages] = useState(true)
  const [baseUrl, setBaseUrl] = useState(() => window.location.href)
  const [isConverting, setIsConverting] = useState(false)
  const [status, setStatus] = useState<string | null>(null)
  const [clipboard, setClipboard] = useState<CanvasElement | null>(null)
  const [canvasMenu, setCanvasMenu] = useState<{ x: number; y: number; elementId: string | null } | null>(
    null
  )
  const [historyTick, setHistoryTick] = useState(0)
  const historyRef = useRef<{ past: DeckSnapshot[]; future: DeckSnapshot[] }>({
    past: [],
    future: [],
  })
  const applyingHistoryRef = useRef(false)
  const textBurstRef = useRef(false)
  const textBurstTimerRef = useRef<number>(0)
  const slidesRef = useRef(slides)
  const activeSlideIdRef = useRef(activeSlideId)
  const selectedIdRef = useRef(selectedId)
  const clipboardRef = useRef(clipboard)
  slidesRef.current = slides
  activeSlideIdRef.current = activeSlideId
  selectedIdRef.current = selectedId
  clipboardRef.current = clipboard

  const activeSlide = slides.find((slide) => slide.id === activeSlideId) ?? slides[0]
  const html = activeSlide?.html ?? ''
  const canvasElements = activeSlide?.canvasElements ?? []

  const setHtml = (value: string | ((prev: string) => string)) => {
    setSlides((prev) =>
      prev.map((slide) => {
        if (slide.id !== activeSlideId) return slide
        const html = typeof value === 'function' ? value(slide.html) : value
        return { ...slide, html }
      })
    )
  }

  const setCanvasElements = (
    value: CanvasElement[] | ((prev: CanvasElement[]) => CanvasElement[])
  ) => {
    setSlides((prev) =>
      prev.map((slide) => {
        if (slide.id !== activeSlideId) return slide
        const canvasElements = typeof value === 'function' ? value(slide.canvasElements) : value
        return { ...slide, canvasElements }
      })
    )
  }

  const persistActiveCanvas = (items: DeckSlide[]) =>
    items.map((slide) =>
      slide.id === activeSlideId && slide.canvasElements.length > 0
        ? { ...slide, html: buildHtmlFromCanvas(slide.canvasElements, background) }
        : slide
    )

  const selectSlide = (id: string) => {
    if (id === activeSlideId) return
    setSlides(persistActiveCanvas)
    setActiveSlideId(id)
    setSelectedId(null)
    setPropertiesId(null)
  }

  const addSlide = () => {
    const slide = createDeckSlide(emptySlideHtml(background))
    setSlides((prev) => {
      const persisted = persistActiveCanvas(prev)
      const index = persisted.findIndex((item) => item.id === activeSlideId)
      const next = [...persisted]
      next.splice(Math.max(0, index) + 1, 0, slide)
      return next
    })
    setActiveSlideId(slide.id)
    setSelectedId(null)
    setPropertiesId(null)
  }

  const deleteSlide = (id: string) => {
    setSlides((prev) => {
      if (prev.length <= 1) return prev
      return prev.filter((slide) => slide.id !== id)
    })
    if (id === activeSlideId) {
      const index = slides.findIndex((slide) => slide.id === id)
      const remaining = slides.filter((slide) => slide.id !== id)
      const neighbor = remaining[Math.max(0, index - 1)] ?? remaining[0]
      if (neighbor) setActiveSlideId(neighbor.id)
      setSelectedId(null)
      setPropertiesId(null)
    }
  }

  const moveSlide = (id: string, direction: -1 | 1) => {
    setSlides((prev) => moveById(prev, id, direction))
  }

  const reorderSlides = (fromId: string, toId: string) => {
    setSlides((prev) => reorderById(prev, fromId, toId))
  }

  const bumpHistory = () => setHistoryTick((tick) => tick + 1)

  const captureNow = () =>
    captureDeck(slidesRef.current, activeSlideIdRef.current, selectedIdRef.current)

  const applySnapshot = (snapshot: DeckSnapshot) => {
    applyingHistoryRef.current = true
    setSlides(cloneDeck(snapshot.slides))
    setActiveSlideId(snapshot.activeSlideId)
    setSelectedId(snapshot.selectedId)
    setPropertiesId(null)
    window.setTimeout(() => {
      applyingHistoryRef.current = false
    }, 0)
  }

  const pushHistory = () => {
    if (applyingHistoryRef.current) return
    historyRef.current.past.push(captureNow())
    if (historyRef.current.past.length > 80) {
      historyRef.current.past.shift()
    }
    historyRef.current.future = []
    bumpHistory()
  }

  const undoCanvas = () => {
    const previous = historyRef.current.past.pop()
    if (!previous) return
    historyRef.current.future.push(captureNow())
    applySnapshot(previous)
    bumpHistory()
  }

  const findSelectedElement = () => {
    const id = selectedIdRef.current
    if (!id) return null
    const slide =
      slidesRef.current.find((item) => item.id === activeSlideIdRef.current) ?? slidesRef.current[0]
    return slide?.canvasElements.find((element) => element.id === id) ?? null
  }

  const copySelected = () => {
    const element = findSelectedElement()
    if (!element) return
    setClipboard(cloneCanvasElement(element))
  }

  const removeElement = (id: string) => {
    setCanvasElements((prev) => prev.filter((el) => el.id !== id))
    setSelectedId((current) => (current === id ? null : current))
    setPropertiesId((current) => (current === id ? null : current))
  }

  const cutSelected = () => {
    const element = findSelectedElement()
    if (!element) return
    setClipboard(cloneCanvasElement(element))
    pushHistory()
    removeElement(element.id)
  }

  const pasteClipboard = (clientX?: number, clientY?: number) => {
    const clip = clipboardRef.current
    if (!clip) return
    pushHistory()
    let x = clip.x + 24
    let y = clip.y + 24
    if (clientX != null && clientY != null && builderStageRef.current) {
      const point = toSlidePx(clientX, clientY)
      x = point.x
      y = point.y
    }
    const pasted: CanvasElement = {
      ...cloneCanvasElement(clip),
      id: `el-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
      x: Math.max(0, x),
      y: Math.max(0, y),
    }
    setCanvasElements((prev) => [...prev, pasted])
    setSelectedId(pasted.id)
  }

  const openCanvasMenu = (event: React.MouseEvent, elementId: string | null = null) => {
    event.preventDefault()
    event.stopPropagation()
    if (elementId) setSelectedId(elementId)
    setPropertiesId(null)
    setCanvasMenu({ x: event.clientX, y: event.clientY, elementId })
  }

  useEffect(() => {
    document.documentElement.style.setProperty('--preview-bg', background)
    return () => {
      document.documentElement.style.removeProperty('--preview-bg')
    }
  }, [background])

  const isValidSize = useMemo(
    () => slideWidth > 0 && slideHeight > 0,
    [slideWidth, slideHeight]
  )
  const slidePx = useMemo(() => slideSizePx(slideWidth, slideHeight), [slideWidth, slideHeight])
  const previewScale = useContainScale(
    previewViewportRef,
    slidePx.width,
    slidePx.height,
    activeTab === 'preview'
  )
  const builderScale = useContainScale(
    builderViewportRef,
    slidePx.width,
    slidePx.height,
    activeTab === 'editor'
  )

  const toSlidePx = (clientX: number, clientY: number) => {
    const stage = builderStageRef.current
    if (!stage) {
      return { x: 0, y: 0, scaleX: 1, scaleY: 1 }
    }
    return clientToSlidePx(clientX, clientY, stage, slidePx.width, slidePx.height)
  }

  const handleConvert = async () => {
    if (!isValidSize) {
      setStatus('Slide dimensions must be greater than 0.')
      return
    }

    const pages = slides.map((slide) =>
      slide.canvasElements.length > 0
        ? buildHtmlFromCanvas(slide.canvasElements, background)
        : slide.html
    )

    setSlides((prev) =>
      prev.map((slide, index) => ({ ...slide, html: pages[index] ?? slide.html }))
    )

    setIsConverting(true)
    setStatus(`Rendering ${pages.length} slide${pages.length === 1 ? '' : 's'}...`)

    try {
      const blob = await htmlToPptx(pages, {
        slideWidthIn: slideWidth,
        slideHeightIn: slideHeight,
        background,
        embedImages,
        baseUrl,
      })

      downloadBlob(blob, filename || 'html-slide.pptx')
      setStatus('PPTX created successfully.')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Failed to create PPTX.')
    } finally {
      setIsConverting(false)
    }
  }

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    const elementType = e.dataTransfer.getData('elementType') as CanvasElement['type']
    if (!elementType) return

    const { x, y } = toSlidePx(e.clientX, e.clientY)

    const newElement = withAutoSize({
      id: `el-${Date.now()}`,
      type: elementType,
      content:
        elementType === 'heading'
          ? 'New Heading'
          : elementType === 'text'
            ? 'New paragraph text'
            : elementType === 'list'
              ? 'List item 1\nList item 2\nList item 3'
              : SAMPLE_IMAGE_DATA_URL,
      x: Math.max(0, x - 50),
      y: Math.max(0, y - 20),
      width: elementType === 'image' ? 140 : 300,
      height: elementType === 'image' ? 140 : 50,
      autoSize: elementType !== 'image',
      aspectRatio: elementType === 'image' ? 1 : undefined,
      fontSize: defaultFontSize(elementType),
      fontFamily: 'Arial',
      color: defaultColor(elementType),
      bold: elementType === 'heading',
      bulletStyle: elementType === 'list' ? 'disc' : undefined,
    })

    pushHistory()
    setCanvasElements([...canvasElements, newElement])
    setSelectedId(newElement.id)
  }

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
  }

  const handleElementUpdate = (id: string, updates: Partial<CanvasElement>) => {
    if (!textBurstRef.current) {
      pushHistory()
      textBurstRef.current = true
    }
    window.clearTimeout(textBurstTimerRef.current)
    textBurstTimerRef.current = window.setTimeout(() => {
      textBurstRef.current = false
    }, 400)
    setCanvasElements((prev) =>
      prev.map((el) => (el.id === id ? withAutoSize({ ...el, ...updates }) : el))
    )
  }

  const handleElementDelete = (id: string) => {
    pushHistory()
    removeElement(id)
  }

  const applyImageToElement = (elementId: string, dataUrl: string) => {
    const img = new Image()
    img.onload = () => {
      const natW = img.naturalWidth || img.width
      const natH = img.naturalHeight || img.height
      setCanvasElements((prev) =>
        prev.map((el) => {
          if (el.id !== elementId) return el
          const fitted = containRect(el.x, el.y, el.width, el.height, natW, natH)
          return {
            ...el,
            content: dataUrl,
            x: fitted.x,
            y: fitted.y,
            width: fitted.width,
            height: fitted.height,
            aspectRatio: natW / natH,
          }
        })
      )
    }
    img.src = dataUrl
  }

  const handleImageDrop = (e: React.DragEvent, elementId: string) => {
    e.preventDefault()
    e.stopPropagation()

    const file = e.dataTransfer.files[0]
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader()
      reader.onload = (event) => {
        const dataUrl = event.target?.result as string
        if (dataUrl) {
          pushHistory()
          applyImageToElement(elementId, dataUrl)
        }
      }
      reader.readAsDataURL(file)
    }
  }

  const handleImageDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
  }

  const parseHtmlToCanvas = async (sourceHtml = html) => {
    const blocks = await measureHtmlBlocks(sourceHtml, slideWidth, slideHeight)
    pushHistory()
    setCanvasElements(
      blocks.map((block, index) => {
        const fitted =
          block.kind === 'image' && block.aspectRatio
            ? containRect(block.x, block.y, block.width, block.height, block.aspectRatio, 1)
            : { x: block.x, y: block.y, width: block.width, height: block.height }

        return withAutoSize({
          id: `parsed-${index}`,
          type: block.kind,
          content: block.content,
          x: fitted.x,
          y: fitted.y,
          width: fitted.width,
          height: fitted.height,
          autoSize: block.kind !== 'image',
          aspectRatio: block.aspectRatio,
          fontSize: block.styles.fontSize
            ? Math.round(block.styles.fontSize / 0.75)
            : defaultFontSize(block.kind),
          fontFamily: block.styles.fontFace ?? 'Arial',
          color: block.styles.color ? `#${block.styles.color}` : defaultColor(block.kind),
          bold: block.styles.bold || block.kind === 'heading',
          italic: block.styles.italic,
          align: block.styles.align,
          bulletStyle: block.kind === 'list' ? block.bulletStyle : undefined,
        })
      })
    )
  }

  const saveCanvasToPreview = () => {
    if (canvasElements.length === 0) return html
    const markup = buildHtmlFromCanvas(canvasElements, background)
    setHtml(markup)
    return markup
  }

  const goToPreview = () => {
    if (activeTab === 'editor') {
      saveCanvasToPreview()
    }
    setActiveTab('preview')
  }

  const htmlDownloadName = () =>
    `${(filename || 'html-slide').replace(/\.(pptx|html|htm)$/i, '')}.html`

  const saveHtmlToFile = async () => {
    const markup = saveCanvasToPreview()
    try {
      await saveTextToChosenFile(markup, htmlDownloadName())
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Failed to save HTML file.')
    }
  }

  const loadHtmlFromFile = (file: File) => {
    const reader = new FileReader()
    reader.onload = async () => {
      const text = typeof reader.result === 'string' ? reader.result : ''
      setHtml(text)
      await parseHtmlToCanvas(text)
    }
    reader.readAsText(file)
  }

  const clearCanvas = () => {
    pushHistory()
    setCanvasElements([])
    setSelectedId(null)
    setPropertiesId(null)
  }

  const beginMove = (e: React.MouseEvent, element: CanvasElement, host: HTMLElement) => {
    e.preventDefault()
    e.stopPropagation()
    pushHistory()
    setSelectedId(element.id)
    const { scaleX, scaleY } = toSlidePx(e.clientX, e.clientY)
    const rect = host.getBoundingClientRect()
    setDraggedElement(element.id)
    setDragOffset({
      x: (e.clientX - rect.left) / scaleX,
      y: (e.clientY - rect.top) / scaleY,
    })
  }

  const handleResizeMouseDown = (e: React.MouseEvent, element: CanvasElement, handle: string) => {
    e.preventDefault()
    e.stopPropagation()
    pushHistory()
    setSelectedId(element.id)
    setResizingElement(element.id)
    setResizeHandle(handle)
    setResizeStart({
      x: e.clientX,
      y: e.clientY,
      width: element.width,
      height: element.height,
      originX: element.x,
      originY: element.y,
    })
  }

  const handleElementMouseDown = (e: React.MouseEvent, element: CanvasElement) => {
    const target = e.target as HTMLElement
    if (target.closest('.element-chrome') || target.closest('.resize-handle')) {
      return
    }
    if (target.closest('textarea, input, button')) {
      setSelectedId(element.id)
      return
    }
    beginMove(e, element, e.currentTarget as HTMLElement)
  }

  const moveSelected = (clientX: number, clientY: number) => {
    const { x, y, scaleX, scaleY } = toSlidePx(clientX, clientY)

    if (draggedElement) {
      setCanvasElements((prev) =>
        prev.map((el) =>
          el.id === draggedElement
            ? { ...el, x: Math.max(0, x - dragOffset.x), y: Math.max(0, y - dragOffset.y) }
            : el
        )
      )
      return
    }

    if (!resizingElement || !resizeHandle) {
      return
    }

    const deltaX = (clientX - resizeStart.x) / scaleX
    const deltaY = (clientY - resizeStart.y) / scaleY

    setCanvasElements((prev) =>
      prev.map((el) => {
        if (el.id !== resizingElement) return el

        if (el.type === 'image') {
          const ratio = el.aspectRatio || resizeStart.width / Math.max(resizeStart.height, 1)
          const next = resizeProportional(resizeHandle, resizeStart, deltaX, deltaY, ratio)
          return { ...el, ...next, autoSize: false }
        }

        let newWidth = el.width
        let newHeight = el.height
        let newX = el.x
        let newY = el.y

        if (resizeHandle.includes('e')) {
          newWidth = Math.max(50, resizeStart.width + deltaX)
        }
        if (resizeHandle.includes('w')) {
          newWidth = Math.max(50, resizeStart.width - deltaX)
          newX = resizeStart.originX + (resizeStart.width - newWidth)
        }
        if (resizeHandle.includes('s')) {
          newHeight = Math.max(30, resizeStart.height + deltaY)
        }
        if (resizeHandle.includes('n')) {
          newHeight = Math.max(30, resizeStart.height - deltaY)
          newY = resizeStart.originY + (resizeStart.height - newHeight)
        }

        return { ...el, width: newWidth, height: newHeight, x: newX, y: newY, autoSize: false }
      })
    )
  }

  const stopMove = () => {
    setDraggedElement(null)
    setResizingElement(null)
    setResizeHandle(null)
  }

  useEffect(() => {
    if (!draggedElement && !resizingElement) return

    const onMove = (event: MouseEvent) => moveSelected(event.clientX, event.clientY)
    const onUp = () => stopMove()
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [draggedElement, resizingElement, resizeHandle, dragOffset, resizeStart, slidePx])

  useEffect(() => {
    const isTextField = (target: EventTarget | null) => {
      const el = target as HTMLElement | null
      if (!el) return false
      const tag = el.tagName
      return tag === 'INPUT' || tag === 'TEXTAREA' || Boolean(el.isContentEditable)
    }

    const onKeyDown = (event: KeyboardEvent) => {
      if (activeTab !== 'editor') return
      const command = event.ctrlKey || event.metaKey
      const key = event.key.toLowerCase()
      const typing = isTextField(event.target)

      if (command && key === 'z' && !event.shiftKey) {
        event.preventDefault()
        undoCanvas()
        setCanvasMenu(null)
        return
      }

      if (command && (key === 'y' || (key === 'z' && event.shiftKey))) {
        event.preventDefault()
        const next = historyRef.current.future.pop()
        if (!next) return
        historyRef.current.past.push(captureNow())
        applySnapshot(next)
        bumpHistory()
        setCanvasMenu(null)
        return
      }

      if (command && key === 'x' && !typing) {
        event.preventDefault()
        cutSelected()
        setCanvasMenu(null)
        return
      }

      if (command && key === 'v' && !typing) {
        event.preventDefault()
        pasteClipboard()
        setCanvasMenu(null)
        return
      }

      if (command && key === 'c' && !typing) {
        event.preventDefault()
        copySelected()
        return
      }

      if (!selectedIdRef.current || typing) return
      if (event.key === 'Delete' || event.key === 'Backspace') {
        event.preventDefault()
        handleElementDelete(selectedIdRef.current)
      }
    }

    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [activeTab])

  useEffect(() => {
    if (!canvasMenu) return
    const onMouseDown = (event: MouseEvent) => {
      const target = event.target as HTMLElement | null
      if (target?.closest('.canvas-menu')) return
      setCanvasMenu(null)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setCanvasMenu(null)
    }
    window.addEventListener('mousedown', onMouseDown)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('mousedown', onMouseDown)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [canvasMenu])

  useEffect(() => {
    if (!propertiesId) return
    const onMouseDown = (event: MouseEvent) => {
      if (event.button === 2) return
      const target = event.target as HTMLElement | null
      if (target?.closest('.properties-panel')) return
      setPropertiesId(null)
    }
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setPropertiesId(null)
    }
    window.addEventListener('mousedown', onMouseDown)
    window.addEventListener('keydown', onKeyDown)
    return () => {
      window.removeEventListener('mousedown', onMouseDown)
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [propertiesId])

  useEffect(() => {
    canvasElements.forEach((el) => {
      document.documentElement.style.setProperty(`--el-${el.id}-x`, `${el.x}px`)
      document.documentElement.style.setProperty(`--el-${el.id}-y`, `${el.y}px`)
      document.documentElement.style.setProperty(`--el-${el.id}-w`, `${el.width}px`)
      document.documentElement.style.setProperty(`--el-${el.id}-h`, `${el.height}px`)
      document.documentElement.style.setProperty(`--el-${el.id}-fs`, el.type === 'heading' ? '24px' : '16px')
      document.documentElement.style.setProperty(`--el-${el.id}-fw`, el.type === 'heading' ? 'bold' : 'normal')
    })
    return () => {
      canvasElements.forEach((el) => {
        document.documentElement.style.removeProperty(`--el-${el.id}-x`)
        document.documentElement.style.removeProperty(`--el-${el.id}-y`)
        document.documentElement.style.removeProperty(`--el-${el.id}-w`)
        document.documentElement.style.removeProperty(`--el-${el.id}-h`)
        document.documentElement.style.removeProperty(`--el-${el.id}-fs`)
        document.documentElement.style.removeProperty(`--el-${el.id}-fw`)
      })
    }
  }, [canvasElements])

  useEffect(() => {
    document.documentElement.style.setProperty('--preview-bg', background)
    return () => {
      document.documentElement.style.removeProperty('--preview-bg')
    }
  }, [background])

  return (
    <div className="app">
      <div className="app-main">
      <header className="app__header">
        <div>
          <p className="eyebrow">HTML to PowerPoint</p>
          <h1>Convert HTML into PPTX slides</h1>
          <p className="subheading">
            Paste HTML, embed images, and export a PowerPoint file with as many slides as you need.
          </p>
        </div>
        <div className="header-actions">
          <button className="icon-btn" onClick={() => setShowSettings(true)} title="Settings">
            ⚙️
          </button>
          <button className="primary" onClick={handleConvert} disabled={isConverting}>
            {isConverting ? 'Converting...' : 'Export PPTX'}
          </button>
        </div>
      </header>

      <div className="grid">
        <section className="panel">
          <div className="panel__header">
            <h2>
              {activeTab === 'preview' ? 'Preview' : activeTab === 'html' ? 'HTML' : 'Editor'}
            </h2>
            <div className="panel__actions">
              <div className="tabs">
                <button
                  className={`tab ${activeTab === 'editor' ? 'tab--active' : ''}`}
                  onClick={() => setActiveTab('editor')}
                >
                  Editor
                </button>
                <button
                  className={`tab ${activeTab === 'preview' ? 'tab--active' : ''}`}
                  onClick={goToPreview}
                >
                  Preview
                </button>
                {(activeTab === 'preview' || activeTab === 'html') && (
                  <button
                    className={`tab ${activeTab === 'html' ? 'tab--active' : ''}`}
                    onClick={() => setActiveTab('html')}
                  >
                    HTML
                  </button>
                )}
              </div>
            </div>
          </div>
          {activeTab === 'html' ? (
            <>
          <label className="field-label" htmlFor="html-input">
            HTML markup
          </label>
            <textarea
              id="html-input"
              value={html}
              onChange={(event) => setHtml(event.target.value)}
              spellCheck={false}
              placeholder="Paste HTML content here"
            />
            </>
          ) : activeTab === 'preview' ? (
            <div className="slide-viewport" ref={previewViewportRef}>
              <div
                className="slide-sizer"
                style={{
                  width: `${slidePx.width * previewScale}px`,
                  height: `${slidePx.height * previewScale}px`,
                }}
              >
                <div
                  className="slide-stage"
                  style={{
                    width: `${slidePx.width}px`,
                    height: `${slidePx.height}px`,
                    transform: `scale(${previewScale})`,
                    background,
                  }}
                  dangerouslySetInnerHTML={{ __html: html }}
                />
              </div>
            </div>
          ) : (
            <div className="builder-container">
              <div className="builder-palette">
                <input
                  ref={htmlFileInputRef}
                  type="file"
                  accept=".html,.htm,text/html"
                  hidden
                  aria-label="Choose HTML file"
                  onChange={(event) => {
                    const file = event.target.files?.[0]
                    if (file) loadHtmlFromFile(file)
                    event.target.value = ''
                  }}
                />
                <button
                  className="builder-action-btn"
                  onClick={() => htmlFileInputRef.current?.click()}
                >
                  Load
                </button>
                <button className="builder-action-btn" onClick={saveHtmlToFile}>
                  Save
                </button>
                <button className="builder-action-btn" onClick={() => parseHtmlToCanvas()}>
                  Load from Preview
                </button>
                <button className="builder-action-btn" onClick={goToPreview}>
                  Preview
                </button>
                <button className="builder-action-btn builder-action-btn--danger" onClick={clearCanvas}>
                  🗑️ Clear Canvas
                </button>
                <h3>Drag elements</h3>
                <div
                  className="palette-item"
                  draggable
                  onDragStart={(e) => e.dataTransfer.setData('elementType', 'heading')}
                >
                  <span>📝</span> Heading
                </div>
                <div
                  className="palette-item"
                  draggable
                  onDragStart={(e) => e.dataTransfer.setData('elementType', 'text')}
                >
                  <span>📄</span> Text
                </div>
                <div
                  className="palette-item"
                  draggable
                  onDragStart={(e) => e.dataTransfer.setData('elementType', 'image')}
                >
                  <span>🖼️</span> Image
                </div>
                <div
                  className="palette-item"
                  draggable
                  onDragStart={(e) => e.dataTransfer.setData('elementType', 'list')}
                >
                  <span>📋</span> List
                </div>
              </div>
              <div className="slide-viewport" ref={builderViewportRef}>
              <div
                className="slide-sizer"
                style={{
                  width: `${slidePx.width * builderScale}px`,
                  height: `${slidePx.height * builderScale}px`,
                }}
              >
              <div
                ref={builderStageRef}
                className="slide-stage builder-canvas"
                style={{
                  width: `${slidePx.width}px`,
                  height: `${slidePx.height}px`,
                  transform: `scale(${builderScale})`,
                  background,
                }}
                onDrop={handleDrop}
                onDragOver={handleDragOver}
                onContextMenu={(event) => openCanvasMenu(event)}
                onMouseDown={(event) => {
                  if (event.target === event.currentTarget) {
                    if (event.button === 2) return
                    setSelectedId(null)
                    setPropertiesId(null)
                    setCanvasMenu(null)
                  }
                }}
              >
                {canvasElements.length === 0 && (
                  <div className="canvas-placeholder">
                    Drag elements here to build your slide
                  </div>
                )}
                {canvasElements.map((element) => (
                  <div
                    key={element.id}
                    className={`canvas-element${draggedElement === element.id ? ' dragging' : ''}${selectedId === element.id ? ' selected' : ''}`}
                    style={{
                      left: `${element.x}px`,
                      top: `${element.y}px`,
                      width: `${element.width}px`,
                      height: `${element.height}px`,
                      cursor: draggedElement === element.id ? 'grabbing' : 'default',
                      fontSize: `${element.fontSize ?? defaultFontSize(element.type)}px`,
                      fontFamily: element.fontFamily ?? 'Arial',
                      color: element.color ?? defaultColor(element.type),
                      fontWeight: (element.bold ?? element.type === 'heading') ? 'bold' : 'normal',
                      fontStyle: element.italic ? 'italic' : 'normal',
                      textAlign: element.align ?? 'left',
                    }}
                    onMouseDown={(e) => handleElementMouseDown(e, element)}
                    onContextMenu={(event) => openCanvasMenu(event, element.id)}
                  >
                    <div className="element-chrome">
                      <button
                        type="button"
                        className="element-grip"
                        title="Drag to move"
                        aria-label="Drag to move"
                        onMouseDown={(event) => {
                          const host = event.currentTarget.closest('.canvas-element')
                          if (host instanceof HTMLElement) {
                            beginMove(event, element, host)
                          }
                        }}
                      >
                        ⋮⋮
                      </button>
                      <button
                        type="button"
                        className="element-btn"
                        title="Delete"
                        aria-label="Delete"
                        onMouseDown={(event) => event.stopPropagation()}
                        onClick={() => handleElementDelete(element.id)}
                      >
                        ✕
                      </button>
                    </div>
                    <div className="resize-handle resize-handle-n" onMouseDown={(e) => handleResizeMouseDown(e, element, 'n')} />
                    <div className="resize-handle resize-handle-s" onMouseDown={(e) => handleResizeMouseDown(e, element, 's')} />
                    <div className="resize-handle resize-handle-e" onMouseDown={(e) => handleResizeMouseDown(e, element, 'e')} />
                    <div className="resize-handle resize-handle-w" onMouseDown={(e) => handleResizeMouseDown(e, element, 'w')} />
                    <div className="resize-handle resize-handle-ne" onMouseDown={(e) => handleResizeMouseDown(e, element, 'ne')} />
                    <div className="resize-handle resize-handle-nw" onMouseDown={(e) => handleResizeMouseDown(e, element, 'nw')} />
                    <div className="resize-handle resize-handle-se" onMouseDown={(e) => handleResizeMouseDown(e, element, 'se')} />
                    <div className="resize-handle resize-handle-sw" onMouseDown={(e) => handleResizeMouseDown(e, element, 'sw')} />
                    <div className="element-body">
                      {element.type === 'image' ? (
                        <img
                          src={element.content}
                          alt="Canvas element"
                          className="element-image"
                          draggable={false}
                          onDrop={(e) => handleImageDrop(e, element.id)}
                          onDragOver={handleImageDragOver}
                        />
                      ) : element.type === 'list' ? (
                        <ListEditor
                          content={element.content}
                          bulletStyle={element.bulletStyle}
                          onChange={(value) => handleElementUpdate(element.id, { content: value })}
                        />
                      ) : (
                        <textarea
                          className="element-input"
                          value={element.content}
                          onChange={(e) =>
                            handleElementUpdate(element.id, { content: e.target.value })
                          }
                          aria-label={`Edit ${element.type}`}
                        />
                      )}
                    </div>
                  </div>
                ))}
              </div>
              </div>
              </div>
            </div>
          )}
        </section>
      </div>
      </div>

      <SlideRail
        slides={slides.map((slide) => ({
          id: slide.id,
          html:
            slide.canvasElements.length > 0
              ? buildHtmlFromCanvas(slide.canvasElements, background)
              : slide.html,
        }))}
        activeId={activeSlide?.id ?? activeSlideId}
        background={background}
        slideWidthPx={slidePx.width}
        slideHeightPx={slidePx.height}
        onSelect={selectSlide}
        onAdd={addSlide}
        onDelete={deleteSlide}
        onMove={moveSlide}
        onReorder={reorderSlides}
      />

      {showSettings && (
        <div className="modal-overlay" onClick={() => setShowSettings(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Export Settings</h2>
              <button className="modal-close" onClick={() => setShowSettings(false)}>
                ✕
              </button>
            </div>
            <div className="modal-body">
              <div className="field">
                <label htmlFor="filename">File name</label>
                <input
                  id="filename"
                  type="text"
                  value={filename}
                  onChange={(event) => setFilename(event.target.value)}
                />
              </div>

              <div className="field-row">
                <div className="field">
                  <label htmlFor="slide-width">Slide width (in)</label>
                  <input
                    id="slide-width"
                    type="number"
                    step="0.1"
                    min="1"
                    value={slideWidth}
                    onChange={(event) => setSlideWidth(Number(event.target.value))}
                  />
                </div>
                <div className="field">
                  <label htmlFor="slide-height">Slide height (in)</label>
                  <input
                    id="slide-height"
                    type="number"
                    step="0.1"
                    min="1"
                    value={slideHeight}
                    onChange={(event) => setSlideHeight(Number(event.target.value))}
                  />
                </div>
              </div>

              <div className="field">
                <label htmlFor="background">Background color</label>
                <input
                  id="background"
                  type="color"
                  value={background}
                  onChange={(event) => setBackground(event.target.value)}
                />
              </div>

              <div className="field">
                <label htmlFor="base-url">Base URL (for relative image paths)</label>
                <input
                  id="base-url"
                  type="text"
                  value={baseUrl}
                  onChange={(event) => setBaseUrl(event.target.value)}
                />
              </div>

              <label className="checkbox">
                <input
                  type="checkbox"
                  checked={embedImages}
                  onChange={(event) => setEmbedImages(event.target.checked)}
                />
                Embed images as data URLs before rendering
              </label>

              <div className="status" role="status">
                {status ?? 'Ready to export.'}
              </div>
            </div>
          </div>
        </div>
      )}
      {canvasMenu && (
        <CanvasContextMenu
          x={canvasMenu.x}
          y={canvasMenu.y}
          canUndo={historyTick >= 0 && historyRef.current.past.length > 0}
          canCut={Boolean(selectedId)}
          canPaste={Boolean(clipboard)}
          showProperties={Boolean(canvasMenu.elementId || selectedId)}
          onUndo={undoCanvas}
          onCut={cutSelected}
          onPaste={() => pasteClipboard(canvasMenu.x, canvasMenu.y)}
          onProperties={() => {
            const id = canvasMenu.elementId ?? selectedId
            if (!id) return
            setPropertiesId(id)
            setPropertiesPos({ x: canvasMenu.x, y: canvasMenu.y })
          }}
          onClose={() => setCanvasMenu(null)}
        />
      )}
      {propertiesId &&
        canvasElements.some((el) => el.id === propertiesId) && (
          <PropertiesPanel
            element={canvasElements.find((el) => el.id === propertiesId)!}
            x={propertiesPos.x}
            y={propertiesPos.y}
            onChange={(updates) => handleElementUpdate(propertiesId, updates)}
            onClose={() => setPropertiesId(null)}
            onDelete={() => handleElementDelete(propertiesId)}
          />
        )}
    </div>
  )
}

export default App
