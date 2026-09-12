import { Buffer } from 'buffer'
import { decodeImage, parseDDSHeader } from 'dds-ktx-parser'
import UPNG from 'upng-js'

// dds-ktx-parser uses global Buffer.alloc internally
if (typeof globalThis.Buffer === 'undefined') {
  ;(globalThis as unknown as { Buffer: typeof Buffer }).Buffer = Buffer
}

const GLB_MAGIC = 0x46546c67 // 'glTF'
const CHUNK_JSON = 0x4e4f534a // 'JSON'
const CHUNK_BIN = 0x004e4942 // 'BIN\0'
const DDS_MAGIC = 0x20534444 // 'DDS '
const DDS_MIME = 'image/vnd.ms-dds'

type GltfImage = {
  mimeType?: string
  bufferView?: number
  uri?: string
  name?: string
}

type GltfTexture = {
  source?: number
  sampler?: number
  name?: string
  extensions?: {
    MSFT_texture_dds?: { source?: number }
    [key: string]: unknown
  }
}

type GltfBufferView = {
  buffer: number
  byteOffset?: number
  byteLength: number
  byteStride?: number
  target?: number
}

type GltfMaterial = {
  pbrMetallicRoughness?: {
    baseColorTexture?: { index?: number; texCoord?: number }
    baseColorFactor?: number[]
    metallicRoughnessTexture?: { index?: number }
    metallicFactor?: number
    roughnessFactor?: number
  }
  normalTexture?: { index?: number }
  occlusionTexture?: { index?: number }
  emissiveTexture?: { index?: number }
  doubleSided?: boolean
  extensions?: Record<string, unknown>
}

type GltfJson = {
  asset?: { version?: string; [key: string]: unknown }
  buffers?: Array<{ byteLength: number; uri?: string }>
  bufferViews?: GltfBufferView[]
  images?: GltfImage[]
  textures?: GltfTexture[]
  materials?: GltfMaterial[]
  meshes?: Array<{ primitives?: Array<{ attributes?: Record<string, number> }> }>
  extensionsUsed?: string[]
  extensionsRequired?: string[]
  [key: string]: unknown
}

export type GlbSanitizeResult = {
  bytes: Uint8Array
  convertedDds: number
  failedDds: number
  convertedSpecularGlossiness: number
  browserSafeTextures: number
  embeddedExternalImages: number
  hasVertexColors: boolean
  hasUvs: boolean
  /** True when interactive GLB would likely render as white clay */
  texturesUnreliable: boolean
}

export type ExternalImageResolver = (uri: string) => Uint8Array | null | Promise<Uint8Array | null>

const textDecoder = new TextDecoder()
const textEncoder = new TextEncoder()

const readHeader = (data: DataView) => {
  const magic = data.getUint32(0, true)
  const version = data.getUint32(4, true)
  const length = data.getUint32(8, true)
  return { magic, version, length }
}

const parseGlb = (
  bytes: Uint8Array
): { json: GltfJson; bin: Uint8Array } | null => {
  if (bytes.byteLength < 12) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const { magic, version } = readHeader(view)
  if (magic !== GLB_MAGIC || version !== 2) return null

  let offset = 12
  let json: GltfJson | null = null
  let bin = new Uint8Array(0)

  while (offset + 8 <= bytes.byteLength) {
    const chunkLength = view.getUint32(offset, true)
    const chunkType = view.getUint32(offset + 4, true)
    offset += 8
    const chunkData = bytes.subarray(offset, offset + chunkLength)
    offset += chunkLength

    if (chunkType === CHUNK_JSON) {
      const jsonText = textDecoder.decode(chunkData).replace(/\0+$/, '')
      json = JSON.parse(jsonText) as GltfJson
    } else if (chunkType === CHUNK_BIN) {
      bin = Uint8Array.from(chunkData)
    }
  }

  if (!json) return null
  return { json, bin }
}

const align4 = (n: number) => (n + 3) & ~3

const buildGlb = (json: GltfJson, bin: Uint8Array): Uint8Array => {
  const jsonText = JSON.stringify(json)
  const jsonBytes = textEncoder.encode(jsonText)
  const jsonPadding = align4(jsonBytes.length) - jsonBytes.length
  const binPadding = align4(bin.length) - bin.length

  const total =
    12 + 8 + jsonBytes.length + jsonPadding + 8 + bin.length + binPadding
  const out = new Uint8Array(total)
  const view = new DataView(out.buffer)

  view.setUint32(0, GLB_MAGIC, true)
  view.setUint32(4, 2, true)
  view.setUint32(8, total, true)

  let offset = 12
  view.setUint32(offset, jsonBytes.length + jsonPadding, true)
  view.setUint32(offset + 4, CHUNK_JSON, true)
  offset += 8
  out.set(jsonBytes, offset)
  offset += jsonBytes.length
  for (let i = 0; i < jsonPadding; i++) out[offset++] = 0x20

  view.setUint32(offset, bin.length + binPadding, true)
  view.setUint32(offset + 4, CHUNK_BIN, true)
  offset += 8
  out.set(bin, offset)
  offset += bin.length
  for (let i = 0; i < binPadding; i++) out[offset++] = 0

  return out
}

const getBufferViewBytes = (
  bin: Uint8Array,
  bufferViews: GltfBufferView[],
  index: number
): Uint8Array | null => {
  const bv = bufferViews[index]
  if (!bv || bv.buffer !== 0) return null
  const start = bv.byteOffset || 0
  const end = start + bv.byteLength
  if (end > bin.byteLength) return null
  return bin.subarray(start, end)
}

const looksLikeDds = (bytes: Uint8Array | null): boolean => {
  if (!bytes || bytes.byteLength < 4) return false
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return view.getUint32(0, true) === DDS_MAGIC
}

const sniffImageMime = (bytes: Uint8Array | null): string | null => {
  if (!bytes || bytes.byteLength < 4) return null
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return 'image/png'
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  if (
    bytes[0] === 0x52 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x46 &&
    bytes.byteLength > 11 &&
    bytes[8] === 0x57 &&
    bytes[9] === 0x45 &&
    bytes[10] === 0x42 &&
    bytes[11] === 0x50
  ) {
    return 'image/webp'
  }
  if (looksLikeDds(bytes)) return DDS_MIME
  return null
}

const meshHasAttribute = (json: GltfJson, attr: string): boolean => {
  for (const mesh of json.meshes || []) {
    for (const prim of mesh.primitives || []) {
      if (prim.attributes && attr in prim.attributes) return true
    }
  }
  return false
}

const isDdsMime = (mime?: string): boolean =>
  !!mime && (mime.toLowerCase().includes('dds') || mime.toLowerCase().includes('vnd.ms-dds'))

const isBrowserImageMime = (mime?: string): boolean => {
  const m = (mime || '').toLowerCase()
  return m === 'image/png' || m === 'image/jpeg' || m === 'image/jpg' || m === 'image/webp'
}

const isUnsupportedCompressedMime = (mime?: string): boolean => {
  const m = (mime || '').toLowerCase()
  return m.includes('ktx') || m.includes('basis') || m.includes('dds')
}

const stripDdsExtensions = (json: GltfJson) => {
  for (const tex of json.textures || []) {
    if (tex.extensions?.MSFT_texture_dds) {
      delete tex.extensions.MSFT_texture_dds
      if (tex.extensions && Object.keys(tex.extensions).length === 0) {
        delete tex.extensions
      }
    }
  }
  if (json.extensionsUsed) {
    json.extensionsUsed = json.extensionsUsed.filter((e) => e !== 'MSFT_texture_dds')
    if (json.extensionsUsed.length === 0) delete json.extensionsUsed
  }
  if (json.extensionsRequired) {
    json.extensionsRequired = json.extensionsRequired.filter((e) => e !== 'MSFT_texture_dds')
    if (json.extensionsRequired.length === 0) delete json.extensionsRequired
  }
}

const countBrowserSafeMaterialTextures = (json: GltfJson): number => {
  const images = json.images || []
  const textures = json.textures || []
  let count = 0

  const textureIsSafe = (texIndex: number | undefined): boolean => {
    if (texIndex === undefined || !textures[texIndex]) return false
    const src = textures[texIndex].source
    if (src === undefined || !images[src]) return false
    return isBrowserImageMime(images[src].mimeType) && !isDdsMime(images[src].mimeType)
  }

  for (const mat of json.materials || []) {
    const sg = mat.extensions?.KHR_materials_pbrSpecularGlossiness as
      | { diffuseTexture?: { index?: number }; specularGlossinessTexture?: { index?: number } }
      | undefined
    const refs = [
      mat.pbrMetallicRoughness?.baseColorTexture?.index,
      mat.pbrMetallicRoughness?.metallicRoughnessTexture?.index,
      mat.normalTexture?.index,
      mat.occlusionTexture?.index,
      mat.emissiveTexture?.index,
      sg?.diffuseTexture?.index,
      sg?.specularGlossinessTexture?.index,
    ]
    for (const ref of refs) {
      if (textureIsSafe(ref)) count++
    }
  }
  return count
}

/**
 * model-viewer / modern three.js dropped KHR_materials_pbrSpecularGlossiness.
 * Office/Sketchfab assets still use it as extensionsRequired → white clay mesh.
 * Convert to metallic-roughness using the diffuse map as baseColor (good visual restore).
 */
const convertSpecularGlossinessToMetallicRoughness = (json: GltfJson): number => {
  let converted = 0
  for (const mat of json.materials || []) {
    const sg = mat.extensions?.KHR_materials_pbrSpecularGlossiness as
      | {
          diffuseFactor?: number[]
          diffuseTexture?: { index: number; texCoord?: number }
          specularFactor?: number[]
          glossinessFactor?: number
          specularGlossinessTexture?: { index: number; texCoord?: number }
        }
      | undefined
    if (!sg) continue

    const hasSgTexture = !!sg.specularGlossinessTexture
    const gloss =
      typeof sg.glossinessFactor === 'number'
        ? sg.glossinessFactor
        : hasSgTexture
          ? 0.5
          : 1
    mat.pbrMetallicRoughness = {
      ...(mat.pbrMetallicRoughness || {}),
      baseColorFactor: sg.diffuseFactor || mat.pbrMetallicRoughness?.baseColorFactor || [1, 1, 1, 1],
      baseColorTexture: sg.diffuseTexture || mat.pbrMetallicRoughness?.baseColorTexture,
      // Characters/props from Office are usually dielectrics; keep metal low
      metallicFactor: 0,
      roughnessFactor: Math.max(0, Math.min(1, 1 - gloss)),
    }

    if (mat.extensions) {
      delete mat.extensions.KHR_materials_pbrSpecularGlossiness
      if (Object.keys(mat.extensions).length === 0) delete mat.extensions
    }
    converted++
  }

  if (converted > 0) {
    if (json.extensionsUsed) {
      json.extensionsUsed = json.extensionsUsed.filter(
        (e) => e !== 'KHR_materials_pbrSpecularGlossiness'
      )
      if (json.extensionsUsed.length === 0) delete json.extensionsUsed
    }
    if (json.extensionsRequired) {
      json.extensionsRequired = json.extensionsRequired.filter(
        (e) => e !== 'KHR_materials_pbrSpecularGlossiness'
      )
      if (json.extensionsRequired.length === 0) delete json.extensionsRequired
    }
    console.log(`Converted ${converted} specular-glossiness material(s) → metallic-roughness`)
  }

  return converted
}

const ddsToPng = (ddsBytes: Uint8Array): Uint8Array | null => {
  try {
    const input = Buffer.from(ddsBytes)
    const info = parseDDSHeader(input)
    if (!info || !info.layers?.length) {
      console.warn('Unsupported or invalid DDS header in GLB texture')
      return null
    }

    const layer = info.layers[0]
    const rgba = decodeImage(input, info.format, layer)
    const rgbaBytes = rgba instanceof Uint8Array ? rgba : new Uint8Array(rgba as ArrayBuffer)
    const rgbaBuffer = rgbaBytes.buffer.slice(
      rgbaBytes.byteOffset,
      rgbaBytes.byteOffset + rgbaBytes.byteLength
    ) as ArrayBuffer
    const png = UPNG.encode(
      [rgbaBuffer],
      layer.shape.width,
      layer.shape.height,
      0
    ) as ArrayBuffer
    return new Uint8Array(png)
  } catch (error) {
    console.warn('Failed converting DDS texture to PNG:', error)
    return null
  }
}

/**
 * Office / Remix 3D GLBs often embed MSFT_texture_dds (DXT/BC7) images,
 * or reference external images beside the GLB in the PPTX package.
 * model-viewer / three.js cannot decode DDS, so the model becomes white clay
 * once the poster is dismissed. Convert DDS → PNG, embed externals, and strip
 * unsupported extensions.
 */
export const sanitizeGlbForBrowser = async (
  input: Uint8Array,
  resolveExternal?: ExternalImageResolver
): Promise<GlbSanitizeResult> => {
  const parsed = parseGlb(input)
  if (!parsed) {
    return {
      bytes: input,
      convertedDds: 0,
      failedDds: 0,
      convertedSpecularGlossiness: 0,
      browserSafeTextures: 0,
      embeddedExternalImages: 0,
      hasVertexColors: false,
      hasUvs: false,
      texturesUnreliable: true,
    }
  }

  const json = JSON.parse(JSON.stringify(parsed.json)) as GltfJson
  const images = json.images || (json.images = [])
  const bufferViews = json.bufferViews || (json.bufferViews = [])
  let bin = parsed.bin
  let cursor = bin.byteLength
  const chunks: Uint8Array[] = [bin]
  const newBufferViews = bufferViews.map((bv) => ({ ...bv }))
  let embeddedExternalImages = 0

  const appendBytes = (bytes: Uint8Array): number => {
    const paddedLen = align4(bytes.length)
    const padded = new Uint8Array(paddedLen)
    padded.set(bytes)
    chunks.push(padded)
    const viewIndex = newBufferViews.length
    newBufferViews.push({
      buffer: 0,
      byteOffset: cursor,
      byteLength: bytes.length,
    })
    cursor += paddedLen
    return viewIndex
  }

  // Resolve external image URIs into bufferViews (common when GLB sits in udata/)
  if (resolveExternal) {
    for (let i = 0; i < images.length; i++) {
      const img = images[i]
      if (!img?.uri || img.bufferView !== undefined) continue
      if (/^data:/i.test(img.uri)) continue
      try {
        const external = await resolveExternal(img.uri)
        if (!external || external.byteLength === 0) continue
        const mime = sniffImageMime(external) || img.mimeType || 'image/png'
        const viewIndex = appendBytes(external)
        images[i] = {
          ...img,
          mimeType: mime,
          bufferView: viewIndex,
          uri: undefined,
        }
        embeddedExternalImages++
      } catch (error) {
        console.warn('Failed embedding external GLB image', img.uri, error)
      }
    }
  }

  // Fill missing mime types from magic bytes
  for (const img of images) {
    if (img.mimeType || img.bufferView === undefined) continue
    const bytes = getBufferViewBytes(
      // provisional bin merge for newly appended views
      (() => {
        const merged = new Uint8Array(cursor)
        let at = 0
        for (const c of chunks) {
          merged.set(c, at)
          at += c.byteLength
        }
        return merged
      })(),
      newBufferViews,
      img.bufferView
    )
    const sniffed = sniffImageMime(bytes)
    if (sniffed) img.mimeType = sniffed
  }

  const provisionalBin = (() => {
    const merged = new Uint8Array(cursor)
    let at = 0
    for (const c of chunks) {
      merged.set(c, at)
      at += c.byteLength
    }
    return merged
  })()

  const ddsImageIndices = new Set<number>()
  images.forEach((img, i) => {
    if (isDdsMime(img.mimeType)) {
      ddsImageIndices.add(i)
      return
    }
    if (img.bufferView !== undefined) {
      const bytes = getBufferViewBytes(provisionalBin, newBufferViews, img.bufferView)
      if (looksLikeDds(bytes)) ddsImageIndices.add(i)
    }
  })
  for (const tex of json.textures || []) {
    const ddsSource = tex.extensions?.MSFT_texture_dds?.source
    if (typeof ddsSource === 'number') ddsImageIndices.add(ddsSource)
  }

  const replacements: Array<{ imageIndex: number; png: Uint8Array }> = []
  let failedDds = 0

  for (const imageIndex of ddsImageIndices) {
    const img = images[imageIndex]
    if (!img || img.bufferView === undefined) {
      failedDds++
      continue
    }
    const ddsBytes = getBufferViewBytes(provisionalBin, newBufferViews, img.bufferView)
    if (!ddsBytes || !looksLikeDds(ddsBytes)) {
      if (isBrowserImageMime(img.mimeType)) continue
      failedDds++
      continue
    }
    const png = ddsToPng(ddsBytes)
    if (!png) {
      failedDds++
      continue
    }
    replacements.push({ imageIndex, png })
  }

  for (const { imageIndex, png } of replacements) {
    const viewIndex = appendBytes(png)
    images[imageIndex] = {
      ...images[imageIndex],
      mimeType: 'image/png',
      bufferView: viewIndex,
      uri: undefined,
    }
  }

  for (const tex of json.textures || []) {
    const ddsSource = tex.extensions?.MSFT_texture_dds?.source
    const primary = tex.source
    const primarySafe =
      primary !== undefined &&
      images[primary] &&
      isBrowserImageMime(images[primary].mimeType)

    if (!primarySafe && typeof ddsSource === 'number') {
      if (replacements.some((r) => r.imageIndex === ddsSource)) {
        tex.source = ddsSource
      } else if (images[ddsSource] && isBrowserImageMime(images[ddsSource].mimeType)) {
        tex.source = ddsSource
      }
    } else if (
      !primarySafe &&
      primary !== undefined &&
      replacements.some((r) => r.imageIndex === primary)
    ) {
      tex.source = primary
    }
  }

  stripDdsExtensions(json)
  const convertedSpecularGlossiness = convertSpecularGlossinessToMetallicRoughness(json)
  json.bufferViews = newBufferViews
  json.images = images

  const outBin = new Uint8Array(cursor)
  let writeAt = 0
  for (const chunk of chunks) {
    outBin.set(chunk, writeAt)
    writeAt += chunk.byteLength
  }
  if (json.buffers && json.buffers[0]) {
    json.buffers[0].byteLength = cursor
    delete json.buffers[0].uri
  } else {
    json.buffers = [{ byteLength: cursor }]
  }

  const finalBytes = buildGlb(json, outBin)
  const browserSafeTextures = countBrowserSafeMaterialTextures(json)
  const hasVertexColors = meshHasAttribute(json, 'COLOR_0')
  const hasUvs = meshHasAttribute(json, 'TEXCOORD_0') || meshHasAttribute(json, 'TEXCOORD_1')

  const stillHasDds = images.some((img) => {
    if (isUnsupportedCompressedMime(img.mimeType) && !isBrowserImageMime(img.mimeType)) return true
    if (img.bufferView === undefined) return false
    return looksLikeDds(getBufferViewBytes(outBin, json.bufferViews || [], img.bufferView))
  })

  const materialsNeedTextures = (json.materials || []).some((mat) =>
    Boolean(
      mat.pbrMetallicRoughness?.baseColorTexture ||
        mat.pbrMetallicRoughness?.metallicRoughnessTexture ||
        mat.normalTexture ||
        mat.occlusionTexture ||
        mat.emissiveTexture
    )
  )

  const materialsHaveNonWhiteFactor = (json.materials || []).some((mat) => {
    const f = mat.pbrMetallicRoughness?.baseColorFactor
    if (!f || f.length < 3) return false
    // Treat near-white as "no colour information"
    return f[0] < 0.92 || f[1] < 0.92 || f[2] < 0.92
  })

  // PowerPoint often stores geometry-only GLBs; colour lives in the poster PNG.
  // UVs without raster images ⇒ textures were stripped or never embedded.
  const missingExpectedTextures =
    (materialsNeedTextures || hasUvs) && browserSafeTextures === 0 && !hasVertexColors

  const hasMeaningfulColor =
    browserSafeTextures > 0 || hasVertexColors || materialsHaveNonWhiteFactor

  const texturesUnreliable =
    stillHasDds ||
    missingExpectedTextures ||
    !hasMeaningfulColor ||
    (materialsNeedTextures && browserSafeTextures === 0) ||
    (failedDds > 0 && replacements.length === 0 && ddsImageIndices.size > 0)

  if (replacements.length > 0 || embeddedExternalImages > 0 || convertedSpecularGlossiness > 0) {
    console.log(
      `Sanitized GLB: DDS→PNG=${replacements.length}, SG→MR=${convertedSpecularGlossiness}, externalImages=${embeddedExternalImages}` +
        (failedDds ? `, ddsFailed=${failedDds}` : '') +
        `, safeMaterialTextures=${browserSafeTextures}`
    )
  } else if (ddsImageIndices.size > 0) {
    console.warn('GLB has DDS textures but none could be converted', { failedDds })
  } else if (missingExpectedTextures) {
    console.warn('GLB has UVs/material texture refs but no browser-safe images — will keep PowerPoint poster')
  }

  return {
    bytes: finalBytes,
    convertedDds: replacements.length,
    failedDds,
    convertedSpecularGlossiness,
    browserSafeTextures,
    embeddedExternalImages,
    hasVertexColors,
    hasUvs,
    texturesUnreliable,
  }
}

export const uint8ToBase64 = (bytes: Uint8Array): string => {
  const chunkSize = 0x8000
  let binary = ''
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length))
    binary += String.fromCharCode.apply(null, Array.from(chunk) as unknown as number[])
  }
  return btoa(binary)
}
