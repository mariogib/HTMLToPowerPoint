export const CSS_PX_PER_INCH = 96

export const slideSizePx = (widthIn: number, heightIn: number) => ({
  width: widthIn * CSS_PX_PER_INCH,
  height: heightIn * CSS_PX_PER_INCH,
})

export const pxToInches = (px: number) => px / CSS_PX_PER_INCH

export const clientToSlidePx = (
  clientX: number,
  clientY: number,
  el: HTMLElement,
  contentWidthPx: number,
  contentHeightPx: number
) => {
  const rect = el.getBoundingClientRect()
  const scaleX = rect.width / contentWidthPx
  const scaleY = rect.height / contentHeightPx
  return {
    x: (clientX - rect.left) / scaleX,
    y: (clientY - rect.top) / scaleY,
    scaleX,
    scaleY,
  }
}
