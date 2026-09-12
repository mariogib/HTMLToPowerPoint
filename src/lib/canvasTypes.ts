import type { BulletStyleId } from './bulletStyles'

export type CanvasElement = {
  id: string
  type: 'text' | 'heading' | 'image' | 'list' | 'html'
  content: string
  x: number
  y: number
  width: number
  height: number
  aspectRatio?: number
  autoSize?: boolean
  fontSize?: number
  fontFamily?: string
  color?: string
  bold?: boolean
  italic?: boolean
  align?: 'left' | 'center' | 'right'
  bulletStyle?: BulletStyleId
}

export type DeckSlide = {
  id: string
  html: string
  canvasElements: CanvasElement[]
}
