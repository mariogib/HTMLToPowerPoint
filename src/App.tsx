import { useEffect, useMemo, useState } from 'react'
import { htmlToPptx } from './lib/htmlToPptx'
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

type CanvasElement = {
  id: string
  type: 'text' | 'heading' | 'image' | 'list'
  content: string
  x: number
  y: number
  width: number
  height: number
  style?: Record<string, string>
}

function App() {
  const [html, setHtml] = useState(SAMPLE_HTML)
  const [activeTab, setActiveTab] = useState<'editor' | 'preview' | 'builder'>('editor')
  const [canvasElements, setCanvasElements] = useState<CanvasElement[]>([])
  const [draggedElement, setDraggedElement] = useState<string | null>(null)
  const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 })
  const [resizingElement, setResizingElement] = useState<string | null>(null)
  const [resizeHandle, setResizeHandle] = useState<string | null>(null)
  const [resizeStart, setResizeStart] = useState({ x: 0, y: 0, width: 0, height: 0 })
  const [showSettings, setShowSettings] = useState(false)
  const [filename, setFilename] = useState('html-slide.pptx')
  const [slideWidth, setSlideWidth] = useState(13.333)
  const [slideHeight, setSlideHeight] = useState(7.5)
  const [background, setBackground] = useState('#ffffff')
  const [embedImages, setEmbedImages] = useState(true)
  const [baseUrl, setBaseUrl] = useState(() => window.location.href)
  const [isConverting, setIsConverting] = useState(false)
  const [status, setStatus] = useState<string | null>(null)

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

  const handleConvert = async () => {
    if (!isValidSize) {
      setStatus('Slide dimensions must be greater than 0.')
      return
    }

    setIsConverting(true)
    setStatus('Rendering HTML...')

    try {
      const blob = await htmlToPptx(html, {
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

    const rect = e.currentTarget.getBoundingClientRect()
    const x = e.clientX - rect.left
    const y = e.clientY - rect.top

    const newElement: CanvasElement = {
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
    }

    setCanvasElements([...canvasElements, newElement])
  }

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault()
  }

  const handleElementUpdate = (id: string, updates: Partial<CanvasElement>) => {
    setCanvasElements((prev) =>
      prev.map((el) => (el.id === id ? { ...el, ...updates } : el))
    )
  }

  const handleElementDelete = (id: string) => {
    setCanvasElements((prev) => prev.filter((el) => el.id !== id))
  }

  const handleImageDrop = (e: React.DragEvent, elementId: string) => {
    e.preventDefault()
    e.stopPropagation()

    const file = e.dataTransfer.files[0]
    if (file && file.type.startsWith('image/')) {
      const reader = new FileReader()
      reader.onload = (event) => {
        const dataUrl = event.target?.result as string
        handleElementUpdate(elementId, { content: dataUrl })
      }
      reader.readAsDataURL(file)
    }
  }

  const handleImageDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
  }

  const parseHtmlToCanvas = () => {
    const container = document.createElement('div')
    container.innerHTML = html
    document.body.appendChild(container)

    const elements: CanvasElement[] = []
    let idCounter = 0

    const processElement = (node: Element, offsetX = 0, offsetY = 0) => {
      const tagName = node.tagName.toLowerCase()
      const rect = (node as HTMLElement).getBoundingClientRect()

      if (tagName === 'h1' || tagName === 'h2' || tagName === 'h3') {
        elements.push({
          id: `parsed-${idCounter++}`,
          type: 'heading',
          content: node.textContent?.trim() || '',
          x: offsetX,
          y: offsetY,
          width: rect.width || 300,
          height: rect.height || 50,
        })
        offsetY += rect.height + 10
      } else if (tagName === 'p') {
        elements.push({
          id: `parsed-${idCounter++}`,
          type: 'text',
          content: node.textContent?.trim() || '',
          x: offsetX,
          y: offsetY,
          width: rect.width || 300,
          height: rect.height || 50,
        })
        offsetY += rect.height + 10
      } else if (tagName === 'img') {
        const img = node as HTMLImageElement
        elements.push({
          id: `parsed-${idCounter++}`,
          type: 'image',
          content: img.src,
          x: offsetX,
          y: offsetY,
          width: img.width || 140,
          height: img.height || 140,
        })
        offsetY += (img.height || 140) + 10
      } else if (tagName === 'ul' || tagName === 'ol') {
        const items = Array.from(node.querySelectorAll('li'))
          .map((li) => li.textContent?.trim() || '')
          .join('\n')
        elements.push({
          id: `parsed-${idCounter++}`,
          type: 'list',
          content: items,
          x: offsetX,
          y: offsetY,
          width: rect.width || 300,
          height: rect.height || 100,
        })
        offsetY += rect.height + 10
      } else if (tagName === 'div') {
        Array.from(node.children).forEach((child) => {
          offsetY = processElement(child, offsetX, offsetY)
        })
      }

      return offsetY
    }

    Array.from(container.children).forEach((child) => {
      processElement(child, 50, 50)
    })

    document.body.removeChild(container)
    setCanvasElements(elements)
  }

  const syncCanvasToHtml = () => {
    if (canvasElements.length === 0) return

    const elements = canvasElements.map((el) => {
      const positionStyle = `position: absolute; left: ${el.x}px; top: ${el.y}px; width: ${el.width}px;`
      
      switch (el.type) {
        case 'heading':
          return `<h1 style="${positionStyle} margin: 0; color: #1f2a44;">${el.content}</h1>`
        case 'text':
          return `<p style="${positionStyle} font-size: 18px; margin: 0; color: #3e4c66;">${el.content}</p>`
        case 'image':
          return `<img src="${el.content}" style="${positionStyle} height: ${el.height}px;" alt="Image" />`
        case 'list':
          const items = el.content.split('\n').filter((line) => line.trim())
          return `<ul style="${positionStyle} margin: 0; padding-left: 20px;">${items.map((item) => `<li>${item}</li>`).join('')}</ul>`
        default:
          return ''
      }
    }).join('\n')

    const wrappedHtml = `<div style="font-family: Arial, sans-serif; padding: 28px; height: 100%; box-sizing: border-box; background: ${background}; position: relative;">
${elements}
</div>`
    setHtml(wrappedHtml)
  }

  const clearCanvas = () => {
    setCanvasElements([])
  }

  const handleResizeMouseDown = (e: React.MouseEvent, element: CanvasElement, handle: string) => {
    e.preventDefault()
    e.stopPropagation()
    
    setResizingElement(element.id)
    setResizeHandle(handle)
    setResizeStart({
      x: e.clientX,
      y: e.clientY,
      width: element.width,
      height: element.height,
    })
  }

  const handleElementMouseDown = (e: React.MouseEvent, element: CanvasElement) => {
    if ((e.target as HTMLElement).closest('.element-toolbar')) {
      return
    }
    if ((e.target as HTMLElement).closest('.resize-handle')) {
      return
    }
    if ((e.target as HTMLElement).tagName === 'TEXTAREA' || (e.target as HTMLElement).tagName === 'IMG') {
      return
    }
    e.preventDefault()
    e.stopPropagation()
    
    const rect = e.currentTarget.getBoundingClientRect()
    setDraggedElement(element.id)
    setDragOffset({
      x: e.clientX - rect.left,
      y: e.clientY - rect.top,
    })
  }

  const handleCanvasMouseMove = (e: React.MouseEvent) => {
    if (draggedElement) {
      const rect = e.currentTarget.getBoundingClientRect()
      const newX = e.clientX - rect.left - dragOffset.x
      const newY = e.clientY - rect.top - dragOffset.y

      setCanvasElements((prev) =>
        prev.map((el) =>
          el.id === draggedElement ? { ...el, x: Math.max(0, newX), y: Math.max(0, newY) } : el
        )
      )
    } else if (resizingElement && resizeHandle) {
      const deltaX = e.clientX - resizeStart.x
      const deltaY = e.clientY - resizeStart.y

      setCanvasElements((prev) =>
        prev.map((el) => {
          if (el.id !== resizingElement) return el

          let newWidth = el.width
          let newHeight = el.height
          let newX = el.x
          let newY = el.y

          if (resizeHandle.includes('e')) {
            newWidth = Math.max(50, resizeStart.width + deltaX)
          }
          if (resizeHandle.includes('w')) {
            newWidth = Math.max(50, resizeStart.width - deltaX)
            newX = el.x + (el.width - newWidth)
          }
          if (resizeHandle.includes('s')) {
            newHeight = Math.max(30, resizeStart.height + deltaY)
          }
          if (resizeHandle.includes('n')) {
            newHeight = Math.max(30, resizeStart.height - deltaY)
            newY = el.y + (el.height - newHeight)
          }

          return { ...el, width: newWidth, height: newHeight, x: newX, y: newY }
        })
      )
    }
  }

  const handleCanvasMouseUp = () => {
    setDraggedElement(null)
    setResizingElement(null)
    setResizeHandle(null)
  }

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
      <header className="app__header">
        <div>
          <p className="eyebrow">HTML to PowerPoint</p>
          <h1>Convert HTML into PPTX slides</h1>
          <p className="subheading">
            Paste HTML, embed images, and export a single-slide PowerPoint file.
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
            <h2>HTML input</h2>
            <div className="panel__actions">
              <button className="ghost" onClick={() => setHtml(SAMPLE_HTML)}>
                Load sample
              </button>
              <div className="tabs">
                <button
                  className={`tab ${activeTab === 'editor' ? 'tab--active' : ''}`}
                  onClick={() => setActiveTab('editor')}
                >
                  Editor
                </button>
                <button
                  className={`tab ${activeTab === 'preview' ? 'tab--active' : ''}`}
                  onClick={() => setActiveTab('preview')}
                >
                  Preview
                </button>
                <button
                  className={`tab ${activeTab === 'builder' ? 'tab--active' : ''}`}
                  onClick={() => setActiveTab('builder')}
                >
                  Builder
                </button>
              </div>
            </div>
          </div>
          <label className="field-label" htmlFor="html-input">
            HTML markup
          </label>
          {activeTab === 'editor' ? (
            <textarea
              id="html-input"
              value={html}
              onChange={(event) => setHtml(event.target.value)}
              spellCheck={false}
              placeholder="Paste HTML content here"
            />
          ) : activeTab === 'preview' ? (
            <div
              className="preview"
              dangerouslySetInnerHTML={{ __html: html }}
            />
          ) : (
            <div className="builder-container">
              <div className="builder-palette">
                <button className="builder-action-btn" onClick={parseHtmlToCanvas}>
                  📥 Load from Editor
                </button>
                <button className="builder-action-btn" onClick={syncCanvasToHtml}>
                  💾 Save to Editor
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
              <div
                className="builder-canvas"
                onDrop={handleDrop}
                onDragOver={handleDragOver}
                onMouseMove={handleCanvasMouseMove}
                onMouseUp={handleCanvasMouseUp}
                onMouseLeave={handleCanvasMouseUp}
              >
                {canvasElements.length === 0 && (
                  <div className="canvas-placeholder">
                    Drag elements here to build your slide
                  </div>
                )}
                {canvasElements.map((element) => (
                  <div
                    key={element.id}
                    className={`canvas-element ${draggedElement === element.id ? 'dragging' : ''}`}
                    style={{
                      left: `${element.x}px`,
                      top: `${element.y}px`,
                      width: `${element.width}px`,
                      minHeight: `${element.height}px`,
                      cursor: draggedElement === element.id ? 'grabbing' : 'grab',
                    }}
                    onMouseDown={(e) => handleElementMouseDown(e, element)}
                  >
                    <div className="element-toolbar">
                      <button
                        className="element-btn"
                        onClick={() => handleElementDelete(element.id)}
                        title="Delete"
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
                    {element.type === 'image' ? (
                      <img
                        src={element.content}
                        alt="Canvas element"
                        className="element-image"
                        onDrop={(e) => handleImageDrop(e, element.id)}
                        onDragOver={handleImageDragOver}
                      />
                    ) : (
                      <textarea
                        className={`element-input ${element.type === 'heading' ? 'element-input--heading' : ''}`}
                        value={element.content}
                        onChange={(e) =>
                          handleElementUpdate(element.id, { content: e.target.value })
                        }
                        aria-label={`Edit ${element.type}`}
                      />
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>
      </div>

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
    </div>
  )
}

export default App
