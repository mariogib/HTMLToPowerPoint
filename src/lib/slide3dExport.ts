import JSZip from 'jszip'
import { EMU_PER_PX } from '../pptxToHtml/utils/pptxSlideSize'

export type ExportableSlideModel = {
  name: string
  glbBytes: Uint8Array
  posterDataUrl?: string
  leftPx: number
  topPx: number
  widthPx: number
  heightPx: number
  cameraOrbit?: string
  cameraTarget?: string
  fieldOfView?: string
  orientation?: string
  exposure?: string
  environmentImage?: string
  shadowIntensity?: string
  shadowSoftness?: string
  toneMapping?: string
  bounds?: ModelBounds
}

type ModelBounds = {
  minX: number
  minY: number
  minZ: number
  maxX: number
  maxY: number
  maxZ: number
}

const MODEL3D_REL =
  'http://schemas.microsoft.com/office/2017/06/relationships/model3d'
const IMAGE_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const AM3D_NS = 'http://schemas.microsoft.com/office/drawing/2017/model3d'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const EMU_PER_METER = 36000000
const GLB_MAGIC = 0x46546c67
const JSON_CHUNK = 0x4e4f534a

const TRANSPARENT_PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (char) => char.charCodeAt(0)
)

const PNG_MAGIC = [0x89, 0x50, 0x4e, 0x47]
const JPEG_MAGIC = [0xff, 0xd8]

const readPx = (style: string, property: string) => {
  const match = style.match(new RegExp(`${property}\\s*:\\s*(-?[\\d.]+)px`, 'i'))
  return match ? Number(match[1]) : 0
}

const dataUrlToBytes = (dataUrl: string): Uint8Array | null => {
  const match = dataUrl.match(/^data:[^;,]*(;base64)?,([\s\S]*)$/i)
  if (!match) return null
  const payload = (match[2] || '').replace(/\s+/g, '')
  if (match[1]) {
    try {
      const binary = atob(payload)
      const bytes = new Uint8Array(binary.length)
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
      return bytes
    } catch {
      return null
    }
  }
  return null
}

const toBinary = (bytes: Uint8Array) => {
  const copy = new Uint8Array(bytes.byteLength)
  copy.set(bytes)
  return copy.buffer
}

const startsWith = (bytes: Uint8Array, magic: number[]) =>
  magic.every((value, index) => bytes[index] === value)

const xmlEscape = (value: string) =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')

type GlbFile = {
  json: Record<string, unknown>
  bin: Uint8Array | null
}

const align4 = (value: number) => value + ((4 - (value % 4)) % 4)

const readGlbFile = (bytes: Uint8Array): GlbFile | null => {
  if (bytes.byteLength < 20) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (view.getUint32(0, true) !== GLB_MAGIC) return null
  const jsonLength = view.getUint32(12, true)
  if (view.getUint32(16, true) !== JSON_CHUNK) return null
  if (20 + jsonLength > bytes.byteLength) return null
  let json: Record<string, unknown>
  try {
    json = JSON.parse(new TextDecoder().decode(bytes.subarray(20, 20 + jsonLength)))
  } catch {
    return null
  }
  const binHeader = align4(20 + jsonLength)
  let bin: Uint8Array | null = null
  if (binHeader + 8 <= bytes.byteLength) {
    const binLength = view.getUint32(binHeader, true)
    const binStart = binHeader + 8
    if (binStart + binLength <= bytes.byteLength) {
      bin = bytes.subarray(binStart, binStart + binLength)
    }
  }
  return { json, bin }
}

const readFloat3 = (data: DataView, offset: number, componentType: number, normalized: boolean) => {
  const values = [0, 0, 0]
  for (let i = 0; i < 3; i++) {
    const at = offset + i * (componentType === 5126 ? 4 : componentType === 5123 || componentType === 5122 ? 2 : 1)
    if (componentType === 5126) values[i] = data.getFloat32(at, true)
    else if (componentType === 5123) values[i] = data.getUint16(at, true) / (normalized ? 65535 : 1)
    else if (componentType === 5122) values[i] = data.getInt16(at, true) / (normalized ? 32767 : 1)
    else if (componentType === 5121) values[i] = data.getUint8(at) / (normalized ? 255 : 1)
    else if (componentType === 5120) values[i] = data.getInt8(at) / (normalized ? 127 : 1)
    else return null
  }
  return values
}

const boundsFromAccessor = (
  accessor: {
    min?: number[]
    max?: number[]
    count?: number
    type?: string
    componentType?: number
    normalized?: boolean
    byteOffset?: number
    bufferView?: number
  },
  bufferViews: Array<{ buffer?: number; byteOffset?: number; byteStride?: number }>,
  bin: Uint8Array | null
): { min: number[]; max: number[] } | null => {
  if (accessor.min?.length && accessor.max?.length && accessor.min.length >= 3 && accessor.max.length >= 3) {
    return { min: accessor.min, max: accessor.max }
  }
  if (!bin || accessor.type !== 'VEC3' || typeof accessor.count !== 'number' || accessor.bufferView == null) {
    return null
  }
  const view = bufferViews[accessor.bufferView]
  if (!view) return null
  const start = (view.byteOffset || 0) + (accessor.byteOffset || 0)
  const componentType = accessor.componentType || 5126
  const stride = view.byteStride || (componentType === 5126 ? 12 : componentType === 5123 || componentType === 5122 ? 6 : 3)
  const data = new DataView(bin.buffer, bin.byteOffset, bin.byteLength)
  const min = [Infinity, Infinity, Infinity]
  const max = [-Infinity, -Infinity, -Infinity]
  for (let i = 0; i < accessor.count; i++) {
    const offset = start + i * stride
    if (offset + 12 > bin.byteLength && componentType === 5126) break
    const values = readFloat3(data, offset, componentType, Boolean(accessor.normalized))
    if (!values) break
    for (let axis = 0; axis < 3; axis++) {
      min[axis] = Math.min(min[axis], values[axis])
      max[axis] = Math.max(max[axis], values[axis])
    }
  }
  if (!Number.isFinite(min[0])) return null
  return { min, max }
}

type GlbNode = {
  mesh?: number
  children?: number[]
  matrix?: number[]
  translation?: number[]
  rotation?: number[]
  scale?: number[]
}

const identityMatrix = () => [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1]

const multiplyMatrix = (a: number[], b: number[]) => {
  const out = new Array(16).fill(0)
  for (let col = 0; col < 4; col++) {
    for (let row = 0; row < 4; row++) {
      out[col * 4 + row] =
        a[0 * 4 + row] * b[col * 4 + 0] +
        a[1 * 4 + row] * b[col * 4 + 1] +
        a[2 * 4 + row] * b[col * 4 + 2] +
        a[3 * 4 + row] * b[col * 4 + 3]
    }
  }
  return out
}

const translationMatrix = (x: number, y: number, z: number) => [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, x, y, z, 1]

const scaleMatrix = (x: number, y: number, z: number) => [x, 0, 0, 0, 0, y, 0, 0, 0, 0, z, 0, 0, 0, 0, 1]

const rotationMatrix = (x: number, y: number, z: number, w: number) => {
  const xx = x * x
  const yy = y * y
  const zz = z * z
  const xy = x * y
  const xz = x * z
  const yz = y * z
  const wx = w * x
  const wy = w * y
  const wz = w * z
  return [
    1 - 2 * (yy + zz),
    2 * (xy + wz),
    2 * (xz - wy),
    0,
    2 * (xy - wz),
    1 - 2 * (xx + zz),
    2 * (yz + wx),
    0,
    2 * (xz + wy),
    2 * (yz - wx),
    1 - 2 * (xx + yy),
    0,
    0,
    0,
    0,
    1,
  ]
}

const nodeMatrix = (node: GlbNode) => {
  if (node.matrix?.length === 16) return node.matrix.slice()
  const translation = node.translation ?? [0, 0, 0]
  const rotation = node.rotation ?? [0, 0, 0, 1]
  const scale = node.scale ?? [1, 1, 1]
  return multiplyMatrix(
    multiplyMatrix(translationMatrix(translation[0], translation[1], translation[2]), rotationMatrix(rotation[0], rotation[1], rotation[2], rotation[3])),
    scaleMatrix(scale[0], scale[1], scale[2])
  )
}

const transformPoint = (matrix: number[], x: number, y: number, z: number) => [
  matrix[0] * x + matrix[4] * y + matrix[8] * z + matrix[12],
  matrix[1] * x + matrix[5] * y + matrix[9] * z + matrix[13],
  matrix[2] * x + matrix[6] * y + matrix[10] * z + matrix[14],
]

const expandWorldBox = (box: { min: number[]; max: number[] }, matrix: number[], bounds: ModelBounds) => {
  for (const x of [box.min[0], box.max[0]]) {
    for (const y of [box.min[1], box.max[1]]) {
      for (const z of [box.min[2], box.max[2]]) {
        const [wx, wy, wz] = transformPoint(matrix, x, y, z)
        bounds.minX = Math.min(bounds.minX, wx)
        bounds.minY = Math.min(bounds.minY, wy)
        bounds.minZ = Math.min(bounds.minZ, wz)
        bounds.maxX = Math.max(bounds.maxX, wx)
        bounds.maxY = Math.max(bounds.maxY, wy)
        bounds.maxZ = Math.max(bounds.maxZ, wz)
      }
    }
  }
}

const readGlbBounds = (bytes: Uint8Array): ModelBounds | undefined => {
  const file = readGlbFile(bytes)
  if (!file) return undefined
  const accessors =
    (file.json.accessors as Array<{
      min?: number[]
      max?: number[]
      count?: number
      type?: string
      componentType?: number
      normalized?: boolean
      byteOffset?: number
      bufferView?: number
    }> | undefined) ?? []
  const bufferViews =
    (file.json.bufferViews as Array<{ buffer?: number; byteOffset?: number; byteStride?: number }> | undefined) ?? []
  const meshes =
    (file.json.meshes as Array<{ primitives?: Array<{ attributes?: { POSITION?: number } }> }> | undefined) ?? []
  const nodes = (file.json.nodes as GlbNode[] | undefined) ?? []
  const scenes = (file.json.scenes as Array<{ nodes?: number[] }> | undefined) ?? []
  const sceneIndex = typeof file.json.scene === 'number' ? file.json.scene : 0
  const roots = scenes[sceneIndex]?.nodes ?? nodes.map((_, index) => index)
  const bounds: ModelBounds = {
    minX: Infinity,
    minY: Infinity,
    minZ: Infinity,
    maxX: -Infinity,
    maxY: -Infinity,
    maxZ: -Infinity,
  }
  let found = false

  const addMesh = (meshIndex: number, matrix: number[]) => {
    const mesh = meshes[meshIndex]
    if (!mesh) return
    for (const primitive of mesh.primitives ?? []) {
      const index = primitive.attributes?.POSITION
      const accessor = typeof index === 'number' ? accessors[index] : undefined
      if (!accessor) continue
      const box = boundsFromAccessor(accessor, bufferViews, file.bin)
      if (!box) continue
      expandWorldBox(box, matrix, bounds)
      found = true
    }
  }

  const walk = (nodeIndex: number, parent: number[]) => {
    const node = nodes[nodeIndex]
    if (!node) return
    const world = multiplyMatrix(parent, nodeMatrix(node))
    if (typeof node.mesh === 'number') addMesh(node.mesh, world)
    for (const child of node.children ?? []) walk(child, world)
  }

  for (const root of roots) walk(root, identityMatrix())
  if (!found) {
    meshes.forEach((_, index) => addMesh(index, identityMatrix()))
  }
  if (!found) return undefined

  const worldExtent = Math.max(
    bounds.maxX - bounds.minX,
    bounds.maxY - bounds.minY,
    bounds.maxZ - bounds.minZ
  )
  const scaledLocal = localMeshExtent(meshes, accessors) * maxNodeScale(nodes)
  if (scaledLocal > worldExtent * 1.05) {
    const extra = scaledLocal / 2
    const cx = (bounds.minX + bounds.maxX) / 2
    const cy = (bounds.minY + bounds.maxY) / 2
    const cz = (bounds.minZ + bounds.maxZ) / 2
    bounds.minX = Math.min(bounds.minX, cx - extra)
    bounds.minY = Math.min(bounds.minY, cy - extra)
    bounds.minZ = Math.min(bounds.minZ, cz - extra)
    bounds.maxX = Math.max(bounds.maxX, cx + extra)
    bounds.maxY = Math.max(bounds.maxY, cy + extra)
    bounds.maxZ = Math.max(bounds.maxZ, cz + extra)
  }
  return bounds
}

const maxNodeScale = (nodes: GlbNode[]) => {
  let scale = 1
  for (const node of nodes) {
    if (node.scale?.length) {
      scale = Math.max(scale, ...node.scale.map((value) => Math.abs(value)))
    }
    if (node.matrix?.length === 16) {
      scale = Math.max(
        scale,
        Math.hypot(node.matrix[0], node.matrix[1], node.matrix[2]),
        Math.hypot(node.matrix[4], node.matrix[5], node.matrix[6]),
        Math.hypot(node.matrix[8], node.matrix[9], node.matrix[10])
      )
    }
  }
  return scale
}

const localMeshExtent = (
  meshes: Array<{ primitives?: Array<{ attributes?: { POSITION?: number } }> }>,
  accessors: Array<{ min?: number[]; max?: number[] }>
) => {
  let extent = 0
  for (const mesh of meshes) {
    for (const primitive of mesh.primitives ?? []) {
      const index = primitive.attributes?.POSITION
      const accessor = typeof index === 'number' ? accessors[index] : undefined
      if (!accessor?.min || !accessor.max || accessor.min.length < 3 || accessor.max.length < 3) continue
      extent = Math.max(
        extent,
        accessor.max[0] - accessor.min[0],
        accessor.max[1] - accessor.min[1],
        accessor.max[2] - accessor.min[2]
      )
    }
  }
  return extent
}

const fitTransform = (bounds?: ModelBounds) => {
  if (!bounds) {
    return { unitN: 1000000, unitD: 1000000, dx: 0, dy: 0, dz: 0 }
  }
  const extent = Math.max(
    bounds.maxX - bounds.minX,
    bounds.maxY - bounds.minY,
    bounds.maxZ - bounds.minZ,
    1e-6
  )
  // PowerPoint uses 1 / maxAxisExtent, not 1 / (2 * extent).
  const metersPerUnit = 1 / extent
  const unitD = 1000000
  const unitN = Math.max(1, Math.round(metersPerUnit * unitD))
  const toEmu = (value: number) => Math.round(-value * (unitN / unitD) * EMU_PER_METER)
  return {
    unitN,
    unitD,
    dx: toEmu((bounds.minX + bounds.maxX) / 2),
    dy: toEmu((bounds.minY + bounds.maxY) / 2),
    dz: toEmu((bounds.minZ + bounds.maxZ) / 2),
  }
}

export const extractExportableModels = (html: string): ExportableSlideModel[] => {
  if (typeof DOMParser === 'undefined' || !html.includes('model-viewer')) {
    return []
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const models: ExportableSlideModel[] = []

  doc.querySelectorAll('model-viewer').forEach((viewer, index) => {
    const src = viewer.getAttribute('src') || ''
    const glbBytes = src.startsWith('data:') ? dataUrlToBytes(src) : null
    if (!glbBytes || glbBytes.byteLength < 12) return

    const host =
      (viewer.closest('.slide-3d-model') as HTMLElement | null) ?? (viewer as HTMLElement)
    const box =
      (viewer.closest('.slide-html-object') as HTMLElement | null) ?? host
    const style = box.getAttribute('style') || host.getAttribute('style') || viewer.getAttribute('style') || ''
    const poster =
      viewer.getAttribute('poster') ||
      viewer.querySelector('img')?.getAttribute('src') ||
      undefined

    models.push({
      name: viewer.getAttribute('alt') || host.getAttribute('title') || `3D Model ${index + 1}`,
      glbBytes,
      posterDataUrl: poster,
      bounds: readGlbBounds(glbBytes),
      leftPx: readPx(style, 'left'),
      topPx: readPx(style, 'top'),
      widthPx: readPx(style, 'width') || 200,
      heightPx: readPx(style, 'height') || 200,
      cameraOrbit: viewer.getAttribute('camera-orbit') || undefined,
      cameraTarget: viewer.getAttribute('camera-target') || undefined,
      fieldOfView: viewer.getAttribute('field-of-view') || undefined,
      orientation: viewer.getAttribute('orientation') || undefined,
      exposure: viewer.getAttribute('exposure') || undefined,
      environmentImage: viewer.getAttribute('environment-image') || undefined,
      shadowIntensity: viewer.getAttribute('shadow-intensity') || undefined,
      shadowSoftness: viewer.getAttribute('shadow-softness') || undefined,
      toneMapping: viewer.getAttribute('tone-mapping') || undefined,
    })
  })

  return models
}

const isUsableImageDataUrl = (dataUrl?: string) => {
  if (!dataUrl) return false
  const bytes = dataUrlToBytes(dataUrl)
  if (!bytes || bytes.byteLength < 24) return false
  return startsWith(bytes, PNG_MAGIC) || startsWith(bytes, JPEG_MAGIC)
}

export const prepareExportableModels = async (html: string): Promise<ExportableSlideModel[]> =>
  extractExportableModels(html)

const posterToImageBytes = async (dataUrl?: string) => {
  if (!isUsableImageDataUrl(dataUrl)) return { bytes: TRANSPARENT_PNG, ext: 'png' as const }
  const bytes = dataUrlToBytes(dataUrl!)!
  if (startsWith(bytes, JPEG_MAGIC)) return { bytes, ext: 'jpeg' as const }
  if (startsWith(bytes, PNG_MAGIC)) return { bytes, ext: 'png' as const }
  return { bytes: TRANSPARENT_PNG, ext: 'png' as const }
}

const nextRelationshipId = (relsXml: string) => {
  const ids = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((match) => Number(match[1]))
  return Math.max(0, ...ids) + 1
}

const nextShapeId = (slideXml: string) => {
  const ids = Array.from(slideXml.matchAll(/\bid="(\d+)"/g)).map((match) => Number(match[1]))
  return Math.max(2, ...ids) + 1
}

const modelMarkup = (
  model: ExportableSlideModel,
  glbRid: string,
  rasterRid: string,
  shapeId: number
) => {
  const x = Math.round(model.leftPx * EMU_PER_PX)
  const y = Math.round(model.topPx * EMU_PER_PX)
  const cx = Math.max(1, Math.round(model.widthPx * EMU_PER_PX))
  const cy = Math.max(1, Math.round(model.heightPx * EMU_PER_PX))
  const name = xmlEscape(model.name)
  const fit = fitTransform(model.bounds ?? readGlbBounds(model.glbBytes))

  // Camera and lights stay PowerPoint-safe. Scale/center come from the GLB so
  // the model fills its box the way model-viewer does in the HTML.
  return `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice xmlns:am3d="${AM3D_NS}" Requires="am3d"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${shapeId}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="${AM3D_NS}"><am3d:model3d r:embed="${glbRid}"><am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr><am3d:camera><am3d:pos x="0" y="0" z="81455949"/><am3d:up dx="0" dy="${EMU_PER_METER}" dz="0"/><am3d:lookAt x="0" y="0" z="0"/><am3d:perspective fov="2700000"/></am3d:camera><am3d:trans><am3d:meterPerModelUnit n="${fit.unitN}" d="${fit.unitD}"/><am3d:preTrans dx="${fit.dx}" dy="${fit.dy}" dz="${fit.dz}"/><am3d:scale><am3d:sx n="1000000" d="1000000"/><am3d:sy n="1000000" d="1000000"/><am3d:sz n="1000000" d="1000000"/></am3d:scale><am3d:rot ax="0" ay="0" az="0"/><am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans><am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed="${rasterRid}"/></am3d:raster><am3d:objViewport viewportSz="5299912"/><am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr><am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight><am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="100000" g="75000" b="50000"/></am3d:clr><am3d:intensity n="9765625" d="1000000"/><am3d:pos x="21959998" y="70920001" z="16344003"/></am3d:ptLight><am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="40000" g="60000" b="95000"/></am3d:clr><am3d:intensity n="12250000" d="1000000"/><am3d:pos x="-37964106" y="51130435" z="57631972"/></am3d:ptLight><am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="86837" g="72700" b="100000"/></am3d:clr><am3d:intensity n="3125000" d="1000000"/><am3d:pos x="-37739122" y="58056624" z="-34769649"/></am3d:ptLight></am3d:model3d></a:graphicData></a:graphic></p:graphicFrame></mc:Choice><mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="${name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rasterRid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></mc:Fallback></mc:AlternateContent>`
}

const ensureGlbContentType = (typesXml: string) => {
  if (/Extension="glb"/i.test(typesXml)) return typesXml
  return typesXml.replace(
    '</Types>',
    '<Default Extension="glb" ContentType="model/gltf.binary"/></Types>'
  )
}

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

const asPptxBlob = (value: Blob | ArrayBuffer | Uint8Array) =>
  value instanceof Blob
    ? value
    : new Blob([value instanceof Uint8Array ? new Uint8Array(value) : value], { type: PPTX_MIME })

export const embedModelsInPptx = async (
  pptxBlob: Blob | ArrayBuffer | Uint8Array,
  modelsBySlide: ExportableSlideModel[][]
): Promise<Blob> => {
  if (!modelsBySlide.some((models) => models.length > 0)) {
    return asPptxBlob(pptxBlob)
  }

  try {
    const zip = await JSZip.loadAsync(pptxBlob)
    const typesFile = zip.file('[Content_Types].xml')
    if (typesFile) {
      zip.file('[Content_Types].xml', ensureGlbContentType(await typesFile.async('string')))
    }

    let modelIndex = 0
    for (let slideIndex = 0; slideIndex < modelsBySlide.length; slideIndex++) {
      const models = modelsBySlide[slideIndex]
      if (!models?.length) continue

      const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`
      const relsPath = `ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`
      const slideFile = zip.file(slidePath)
      const relsFile = zip.file(relsPath)
      if (!slideFile || !relsFile) continue

      let slideXml = await slideFile.async('string')
      let relsXml = await relsFile.async('string')
      let relId = nextRelationshipId(relsXml)
      let shapeId = nextShapeId(slideXml)
      const fragments: string[] = []

      for (const model of models) {
        modelIndex += 1
        const glbName = `model3d${modelIndex}.glb`
        const raster = await posterToImageBytes(model.posterDataUrl)
        const rasterName = `model3d${modelIndex}-raster.${raster.ext === 'jpeg' ? 'jpeg' : 'png'}`
        zip.file(`ppt/media/${glbName}`, toBinary(model.glbBytes), { binary: true })
        zip.file(`ppt/media/${rasterName}`, toBinary(raster.bytes), { binary: true })

        const glbRid = `rId${relId++}`
        const rasterRid = `rId${relId++}`
        relsXml = relsXml.replace(
          '</Relationships>',
          `<Relationship Id="${glbRid}" Type="${MODEL3D_REL}" Target="../media/${glbName}"/>` +
            `<Relationship Id="${rasterRid}" Type="${IMAGE_REL}" Target="../media/${rasterName}"/>` +
            '</Relationships>'
        )
        fragments.push(modelMarkup(model, glbRid, rasterRid, shapeId++))
      }

      if (!slideXml.includes('</p:spTree>') || fragments.length === 0) continue
      slideXml = slideXml.replace('</p:spTree>', `${fragments.join('')}</p:spTree>`)
      zip.file(slidePath, slideXml)
      zip.file(relsPath, relsXml)
    }

    return zip.generateAsync({
      type: 'blob',
      mimeType: PPTX_MIME,
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    })
  } catch (error) {
    console.warn('Failed to embed 3D models, returning snapshot PPTX:', error)
    return asPptxBlob(pptxBlob)
  }
}
