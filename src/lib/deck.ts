import type { CanvasElement, DeckSlide } from './canvasTypes'

export const emptySlideHtml = (background: string) =>
  `<div style="font-family: Arial, sans-serif; width: 100%; height: 100%; box-sizing: border-box; background: ${background}; position: relative;"></div>`

export const createDeckSlide = (
  html: string,
  canvasElements: CanvasElement[] = []
): DeckSlide => ({
  id: `slide-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`,
  html,
  canvasElements,
})

export const reorderById = <T extends { id: string }>(items: T[], fromId: string, toId: string) => {
  const from = items.findIndex((item) => item.id === fromId)
  const to = items.findIndex((item) => item.id === toId)
  if (from < 0 || to < 0 || from === to) return items
  const next = [...items]
  const [item] = next.splice(from, 1)
  next.splice(to, 0, item)
  return next
}

export const moveById = <T extends { id: string }>(items: T[], id: string, direction: -1 | 1) => {
  const index = items.findIndex((item) => item.id === id)
  const nextIndex = index + direction
  if (index < 0 || nextIndex < 0 || nextIndex >= items.length) return items
  const next = [...items]
  const current = next[index]
  const swap = next[nextIndex]
  if (!current || !swap) return items
  next[index] = swap
  next[nextIndex] = current
  return next
}
