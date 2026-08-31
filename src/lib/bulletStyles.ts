export type BulletStyleId =
  | 'disc'
  | 'circle'
  | 'square'
  | 'dash'
  | 'arrow'
  | 'check'
  | 'number'
  | 'none'

export type BulletStyle = {
  id: BulletStyleId
  label: string
  preview: string
  cssType: string
  pptxChar?: string
}

export const BULLET_STYLES: BulletStyle[] = [
  { id: 'disc', label: 'Filled circle', preview: '•', cssType: 'disc' },
  { id: 'circle', label: 'Hollow circle', preview: '○', cssType: 'circle', pptxChar: '25CB' },
  { id: 'square', label: 'Square', preview: '■', cssType: 'square', pptxChar: '25A0' },
  { id: 'dash', label: 'Dash', preview: '–', cssType: '"– "', pptxChar: '2013' },
  { id: 'arrow', label: 'Arrow', preview: '►', cssType: '"► "', pptxChar: '25BA' },
  { id: 'check', label: 'Check', preview: '✓', cssType: '"✓ "', pptxChar: '2713' },
  { id: 'number', label: 'Numbers', preview: '1.', cssType: 'decimal' },
  { id: 'none', label: 'None', preview: '∅', cssType: 'none' },
]

const STYLE_IDS = new Set<string>(BULLET_STYLES.map((style) => style.id))

export const isBulletStyleId = (value: string | null | undefined): value is BulletStyleId =>
  Boolean(value && STYLE_IDS.has(value))

export const getBulletStyle = (id?: BulletStyleId | string | null): BulletStyle =>
  BULLET_STYLES.find((style) => style.id === id) ?? BULLET_STYLES[0]

export const bulletPreview = (id: BulletStyleId | undefined, index: number) => {
  const style = getBulletStyle(id)
  if (style.id === 'none') return ''
  if (style.id === 'number') return `${index + 1}.`
  return style.preview
}

export const detectBulletStyle = (element: HTMLElement): BulletStyleId => {
  const data = element.getAttribute('data-bullet')
  if (isBulletStyleId(data)) {
    return data
  }
  if (element.tagName.toLowerCase() === 'ol') {
    return 'number'
  }

  const type = window.getComputedStyle(element).listStyleType
  if (type === 'circle') return 'circle'
  if (type === 'square') return 'square'
  if (type === 'none') return 'none'
  if (type === 'decimal' || type === 'decimal-leading-zero') return 'number'
  if (type.includes('–') || type.includes('—') || type.includes('-')) return 'dash'
  if (type.includes('►')) return 'arrow'
  if (type.includes('✓') || type.includes('✔')) return 'check'
  if (type.includes('○')) return 'circle'
  if (type.includes('■')) return 'square'
  return 'disc'
}

export const pptxBulletOption = (id: BulletStyleId | undefined, indent: number) => {
  const style = getBulletStyle(id)
  if (style.id === 'none') {
    return false
  }
  if (style.id === 'number') {
    return { type: 'number' as const, indent }
  }
  if (style.pptxChar) {
    return { characterCode: style.pptxChar, indent }
  }
  return { indent }
}
