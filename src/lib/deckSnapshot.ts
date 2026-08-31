import type { CanvasElement, DeckSlide } from './canvasTypes'

export type DeckSnapshot = {
  slides: DeckSlide[]
  activeSlideId: string
  selectedId: string | null
}

export const cloneCanvasElement = (element: CanvasElement): CanvasElement => ({ ...element })

export const cloneDeck = (slides: DeckSlide[]): DeckSlide[] =>
  slides.map((slide) => ({
    ...slide,
    canvasElements: slide.canvasElements.map(cloneCanvasElement),
  }))

export const captureDeck = (
  slides: DeckSlide[],
  activeSlideId: string,
  selectedId: string | null
): DeckSnapshot => ({
  slides: cloneDeck(slides),
  activeSlideId,
  selectedId,
})
