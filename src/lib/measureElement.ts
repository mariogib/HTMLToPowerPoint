import type { CanvasElement } from './canvasTypes'

const CHROME_HEIGHT = 28
const TEXT_PADDING_X = 10
const TEXT_PADDING_Y = 6

const escapeHtml = (value: string) =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

export const defaultFontSize = (type: CanvasElement['type']) =>
  type === 'heading' ? 24 : type === 'text' ? 18 : 16

export const defaultColor = (type: CanvasElement['type']) =>
  type === 'text' ? '#3e4c66' : '#1f2a44'

export const measureElementSize = (element: CanvasElement) => {
  if (element.type === 'image' || element.type === 'html') {
    return { width: element.width, height: element.height }
  }

  const fontSize = element.fontSize ?? defaultFontSize(element.type)
  const fontFamily = element.fontFamily ?? 'Arial, sans-serif'
  const fontWeight = element.bold || element.type === 'heading' ? 'bold' : 'normal'
  const fontStyle = element.italic ? 'italic' : 'normal'

  const probe = document.createElement('div')
  probe.style.cssText = [
    'position:absolute',
    'left:-9999px',
    'top:0',
    'width:max-content',
    'max-width:960px',
    'height:auto',
    `font-size:${fontSize}px`,
    `font-family:${fontFamily}`,
    `font-weight:${fontWeight}`,
    `font-style:${fontStyle}`,
    'line-height:1.4',
    'white-space:pre-wrap',
    `padding:${TEXT_PADDING_Y}px ${TEXT_PADDING_X}px`,
    'box-sizing:border-box',
  ].join(';')

  if (element.type === 'list') {
    const items = element.content.split('\n')
    probe.innerHTML = `<ul style="margin:0;padding-left:1.2em;list-style:disc">${items
      .map((item) => `<li>${escapeHtml(item || ' ')}</li>`)
      .join('')}</ul>`
  } else {
    probe.textContent = element.content || ' '
  }

  document.body.appendChild(probe)
  const width = Math.ceil(probe.offsetWidth) + 4
  const height = Math.ceil(probe.offsetHeight) + CHROME_HEIGHT
  document.body.removeChild(probe)

  return {
    width: Math.max(72, width),
    height: Math.max(CHROME_HEIGHT + 28, height),
  }
}

export const withAutoSize = (element: CanvasElement): CanvasElement => {
  if (element.type === 'image' || element.type === 'html' || element.autoSize === false) {
    return element
  }
  const size = measureElementSize(element)
  return { ...element, width: size.width, height: size.height }
}
