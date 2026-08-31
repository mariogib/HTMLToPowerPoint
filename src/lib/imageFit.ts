export type FittedBox = {
  x: number
  y: number
  width: number
  height: number
}

export const containRect = (
  boxX: number,
  boxY: number,
  boxW: number,
  boxH: number,
  natW: number,
  natH: number
): FittedBox => {
  if (natW <= 0 || natH <= 0 || boxW <= 0 || boxH <= 0) {
    return { x: boxX, y: boxY, width: boxW, height: boxH }
  }

  const scale = Math.min(boxW / natW, boxH / natH)
  const width = natW * scale
  const height = natH * scale

  return {
    x: boxX + (boxW - width) / 2,
    y: boxY + (boxH - height) / 2,
    width,
    height,
  }
}

export const loadImageNaturalSize = (src: string) =>
  new Promise<{ width: number; height: number }>((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      const width = img.naturalWidth || img.width
      const height = img.naturalHeight || img.height
      if (!width || !height) {
        reject(new Error('Image has no dimensions'))
        return
      }
      resolve({ width, height })
    }
    img.onerror = () => reject(new Error('Failed to load image'))
    img.src = src
  })

type ResizeStart = {
  width: number
  height: number
  originX: number
  originY: number
}

export const resizeProportional = (
  handle: string,
  start: ResizeStart,
  deltaX: number,
  deltaY: number,
  ratio: number,
  minWidth = 50,
  minHeight = 30
): FittedBox => {
  const safeRatio = ratio > 0 ? ratio : start.width / Math.max(start.height, 1)
  const fromX = handle.includes('e') || handle.includes('w')
  const fromY = handle.includes('n') || handle.includes('s')

  let width = start.width
  let height = start.height

  if (fromX && fromY) {
    const widthFromX = handle.includes('e') ? start.width + deltaX : start.width - deltaX
    const heightFromY = handle.includes('s') ? start.height + deltaY : start.height - deltaY
    const relW = Math.abs(widthFromX / Math.max(start.width, 1))
    const relH = Math.abs(heightFromY / Math.max(start.height, 1))
    if (relW >= relH) {
      width = Math.max(minWidth, widthFromX)
      height = width / safeRatio
    } else {
      height = Math.max(minHeight, heightFromY)
      width = height * safeRatio
    }
  } else if (fromX) {
    width = Math.max(minWidth, handle.includes('e') ? start.width + deltaX : start.width - deltaX)
    height = width / safeRatio
  } else if (fromY) {
    height = Math.max(minHeight, handle.includes('s') ? start.height + deltaY : start.height - deltaY)
    width = height * safeRatio
  }

  if (height < minHeight) {
    height = minHeight
    width = height * safeRatio
  }
  if (width < minWidth) {
    width = minWidth
    height = width / safeRatio
  }

  return {
    x: handle.includes('w') ? start.originX + start.width - width : start.originX,
    y: handle.includes('n') ? start.originY + start.height - height : start.originY,
    width,
    height,
  }
}
