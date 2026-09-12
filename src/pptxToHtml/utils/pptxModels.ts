import type JSZip from 'jszip'
import { loadRelationships, loadXml } from './pptxBackground'
import { sanitizeGlbForBrowser, uint8ToBase64 } from './glbSanitize'
import { DEFAULT_SLIDE_SIZE, emuToPx, type SlideSize } from './pptxSlideSize'

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface Slide3DModel {
  id: string
  name: string
  /** In-app / fallback URL (blob: or data:) */
  dataUrl: string
  /** Raw sanitized GLB bytes for multi-file zip export */
  glbBytes?: Uint8Array
  /** Relative path inside export zip, e.g. media/model-1.glb */
  exportPath?: string
  posterDataUrl?: string
  /** When true, GLB textures could not be made browser-safe — prefer poster */
  posterOnly?: boolean
  leftPx: number
  topPx: number
  widthPx: number
  heightPx: number
}

const localName = (el: Element | null | undefined): string => {
  if (!el) return ''
  return el.localName || el.tagName.replace(/^.*:/, '')
}

const elementsByLocal = (root: ParentNode | null | undefined, name: string): Element[] => {
  if (!root) return []
  const scope = root instanceof Document ? root.documentElement : (root as Element)
  if (!scope?.getElementsByTagName) return []
  return Array.from(scope.getElementsByTagName('*')).filter((el) => localName(el) === name) as Element[]
}

const firstByLocal = (parent: ParentNode | null | undefined, name: string): Element | null =>
  elementsByLocal(parent, name)[0] || null

const getZipEntry = (zip: JSZip, path: string) => {
  const candidates = [path, path.replace(/^\//, ''), `/${path.replace(/^\//, '')}`]
  for (const candidate of candidates) {
    const entry = zip.file(candidate) || zip.files[candidate]
    if (entry && !entry.dir) return entry
  }
  const needle = path.replace(/^\//, '').toLowerCase()
  const match = Object.keys(zip.files).find((key) => key.replace(/^\//, '').toLowerCase() === needle)
  return match ? zip.files[match] : null
}

const readEmu = (el: Element | null, attr: string): number => {
  if (!el) return 0
  return parseInt(el.getAttribute(attr) || '0', 10) || 0
}

const readTransform = (
  container: Element | null,
  slideSize: SlideSize
): {
  leftPx: number
  topPx: number
  widthPx: number
  heightPx: number
} => {
  const xfrm =
    firstByLocal(container, 'xfrm') ||
    firstByLocal(firstByLocal(container, 'spPr'), 'xfrm') ||
    firstByLocal(firstByLocal(container, 'pic'), 'xfrm')

  const off = firstByLocal(xfrm, 'off')
  const ext = firstByLocal(xfrm, 'ext')
  const x = readEmu(off, 'x')
  const y = readEmu(off, 'y')
  const cx = readEmu(ext, 'cx') || slideSize.widthEmu * 0.4
  const cy = readEmu(ext, 'cy') || slideSize.heightEmu * 0.4

  return {
    leftPx: Math.max(0, emuToPx(x)),
    topPx: Math.max(0, emuToPx(y)),
    widthPx: Math.max(80, emuToPx(cx)),
    heightPx: Math.max(80, emuToPx(cy)),
  }
}

const getEmbedId = (el: Element | null): string | null => {
  if (!el) return null
  return (
    el.getAttributeNS(R_NS, 'embed') ||
    el.getAttribute('r:embed') ||
    el.getAttribute('embed') ||
    el.getAttribute('r:link') ||
    el.getAttributeNS(R_NS, 'link')
  )
}

const isGlbPath = (path: string): boolean => {
  const lower = path.toLowerCase()
  return lower.endsWith('.glb') || lower.endsWith('.gltf')
}

const isImagePath = (path: string): boolean => {
  const lower = path.toLowerCase()
  return (
    lower.endsWith('.png') ||
    lower.endsWith('.jpg') ||
    lower.endsWith('.jpeg') ||
    lower.endsWith('.gif') ||
    lower.endsWith('.webp') ||
    lower.endsWith('.emf') ||
    lower.endsWith('.wmf')
  )
}

const mimeForMedia = (path: string): string => {
  const ext = path.split('.').pop()?.toLowerCase()
  switch (ext) {
    case 'glb':
      return 'model/gltf-binary'
    case 'gltf':
      return 'model/gltf+json'
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg'
    case 'gif':
      return 'image/gif'
    case 'webp':
      return 'image/webp'
    case 'png':
    default:
      return 'image/png'
  }
}

const toDataUrl = async (zip: JSZip, mediaPath: string): Promise<string | null> => {
  const entry = getZipEntry(zip, mediaPath)
  if (!entry) return null
  try {
    const base64 = await entry.async('base64')
    return `data:${mimeForMedia(mediaPath)};base64,${base64}`
  } catch (error) {
    console.error('Failed reading media part:', mediaPath, error)
    return null
  }
}

const toSanitizedGlbDataUrl = async (
  zip: JSZip,
  mediaPath: string,
  modelId: string
): Promise<{
  dataUrl: string
  glbBytes: Uint8Array
  exportPath: string
  posterOnly: boolean
  convertedDds: number
  safeTextures: number
  convertedSg: number
} | null> => {
  const entry = getZipEntry(zip, mediaPath)
  if (!entry) return null
  try {
    const raw = await entry.async('uint8array')
    const glbDir = mediaPath.replace(/\\/g, '/').split('/').slice(0, -1).join('/')

    const resolveExternal = async (uri: string): Promise<Uint8Array | null> => {
      const cleaned = uri.replace(/^\.\//, '').replace(/\\/g, '/')
      const candidates = [
        cleaned,
        `${glbDir}/${cleaned}`,
        `ppt/media/${cleaned.split('/').pop()}`,
        `ppt/slides/udata/${cleaned.split('/').pop()}`,
      ]
      for (const candidate of candidates) {
        const file = getZipEntry(zip, candidate)
        if (file) {
          console.log(`Embedding external GLB image ${uri} from ${candidate}`)
          return file.async('uint8array')
        }
      }
      console.warn(`External GLB image not found in PPTX: ${uri}`)
      return null
    }

    const sanitized = await sanitizeGlbForBrowser(raw, resolveExternal)
    // Prefer blob URLs in-app — huge data: URLs break texture decoding for 20–40MB models
    const blob = new Blob([Uint8Array.from(sanitized.bytes)], { type: 'model/gltf-binary' })
    const dataUrl = URL.createObjectURL(blob)
    const exportPath = `media/${modelId}.glb`

    return {
      dataUrl,
      glbBytes: sanitized.bytes,
      exportPath,
      posterOnly: sanitized.texturesUnreliable,
      convertedDds: sanitized.convertedDds,
      safeTextures: sanitized.browserSafeTextures,
      convertedSg: sanitized.convertedSpecularGlossiness,
    }
  } catch (error) {
    console.error('Failed sanitizing GLB:', mediaPath, error)
    const raw = await entry.async('uint8array')
    const blob = new Blob([Uint8Array.from(raw)], { type: 'model/gltf-binary' })
    return {
      dataUrl: URL.createObjectURL(blob),
      glbBytes: raw,
      exportPath: `media/${modelId}.glb`,
      posterOnly: true,
      convertedDds: 0,
      safeTextures: 0,
      convertedSg: 0,
    }
  }
}

const findNearestTransformContainer = (modelEl: Element): Element => {
  let current: Element | null = modelEl
  while (current) {
    const name = localName(current)
    if (
      name === 'graphicFrame' ||
      name === 'sp' ||
      name === 'pic' ||
      name === 'AlternateContent' ||
      name === 'Choice' ||
      name === 'Fallback'
    ) {
      if (firstByLocal(current, 'xfrm') || firstByLocal(current, 'spPr')) {
        return current
      }
    }
    current = current.parentElement
  }
  return modelEl
}

const findPosterNearModel = (
  modelEl: Element,
  rels: Map<string, { type: string; target: string }>
): string | null => {
  // PowerPoint stores the coloured render on am3d:raster/am3d:blip and Fallback
  const blipCandidates = [
    ...elementsByLocal(modelEl, 'blip'),
    ...(() => {
      let current: Element | null = modelEl
      const found: Element[] = []
      while (current) {
        if (localName(current) === 'AlternateContent') {
          found.push(...elementsByLocal(current, 'blip'))
        }
        current = current.parentElement
      }
      return found
    })(),
  ]

  for (const blip of blipCandidates) {
    const embed = getEmbedId(blip)
    if (embed && rels.get(embed) && isImagePath(rels.get(embed)!.target)) {
      return rels.get(embed)!.target
    }
  }

  const frame = findNearestTransformContainer(modelEl)
  const blip = firstByLocal(frame, 'blip')
  const embed = getEmbedId(blip)
  if (embed && rels.get(embed) && isImagePath(rels.get(embed)!.target)) {
    return rels.get(embed)!.target
  }
  return null
}

/** Only real model3d nodes — never Choice wrappers or other AlternateContent. */
const collectModel3dElements = (doc: Document): Element[] => {
  const found = [
    ...elementsByLocal(doc, 'model3d'),
    ...elementsByLocal(doc, 'model3D'),
    ...elementsByLocal(doc, 'Model3D'),
  ]
  return Array.from(new Set(found))
}

const resolveGlbEmbed = (
  modelEl: Element,
  rels: Map<string, { type: string; target: string }>
): string | null => {
  const embed = getEmbedId(modelEl)
  if (!embed || !rels.has(embed)) return null
  const target = rels.get(embed)!.target
  return isGlbPath(target) ? target : null
}

export const extractSlide3DModels = async (
  zip: JSZip,
  slidePath: string,
  slideDoc: Document,
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<Slide3DModel[]> => {
  const rels = await loadRelationships(zip, slidePath)
  const models: Slide3DModel[] = []
  const modelElements = collectModel3dElements(slideDoc)

  for (let i = 0; i < modelElements.length; i++) {
    const modelEl = modelElements[i]
    const glbPath = resolveGlbEmbed(modelEl, rels)
    if (!glbPath) {
      console.warn(`Skipping model3d on ${slidePath}: no explicit .glb r:embed`)
      continue
    }

    const modelId = `model-${models.length + 1}`
    const glb = await toSanitizedGlbDataUrl(zip, glbPath, modelId)
    if (!glb) continue

    const posterPath = findPosterNearModel(modelEl, rels)
    const posterDataUrl = posterPath ? (await toDataUrl(zip, posterPath)) || undefined : undefined
    const transform = readTransform(findNearestTransformContainer(modelEl), slideSize)
    const name = glbPath.split('/').pop() || modelId

    const usePosterOnly = Boolean(posterDataUrl && glb.posterOnly)

    models.push({
      id: modelId,
      name,
      dataUrl: glb.dataUrl,
      glbBytes: glb.glbBytes,
      exportPath: glb.exportPath,
      posterDataUrl,
      posterOnly: usePosterOnly,
      ...transform,
    })

    console.log(`Extracted 3D model on ${slidePath}:`, name, {
      poster: !!posterDataUrl,
      posterOnly: usePosterOnly,
      convertedDds: glb.convertedDds,
      convertedSg: glb.convertedSg,
      safeTextures: glb.safeTextures,
      bytes: glb.glbBytes.byteLength,
      size: `${Math.round(transform.widthPx)}x${Math.round(transform.heightPx)}`,
      at: `${Math.round(transform.leftPx)},${Math.round(transform.topPx)}`,
    })
  }

  return models
}

export const render3DModelsHtml = (models: Slide3DModel[]): string => {
  if (models.length === 0) return ''

  return models
    .map((model) => {
      const box = `position:absolute;left:${model.leftPx.toFixed(1)}px;top:${model.topPx.toFixed(1)}px;width:${model.widthPx.toFixed(1)}px;height:${model.heightPx.toFixed(1)}px;z-index:5;`

      if (model.posterOnly && model.posterDataUrl) {
        const orbitId = `poster-orbit-${model.id}`
        return `
    <div class="slide-3d-model slide-3d-poster-orbit" data-model="${model.id}" title="${model.name}" style="${box}">
      <div class="poster-orbit-scene" id="${orbitId}">
        <img class="slide-3d-poster poster-orbit-img" src="${model.posterDataUrl}" alt="${model.name}" draggable="false" />
      </div>
      <script>
        (function () {
          const root = document.getElementById(${JSON.stringify(orbitId)});
          if (!root) return;
          const img = root.querySelector('img');
          if (!img) return;
          let ry = 0, rx = 8, dragging = false, lx = 0, ly = 0;
          const apply = function () {
            img.style.transform = 'rotateX(' + rx + 'deg) rotateY(' + ry + 'deg)';
          };
          apply();
          setInterval(function () { if (!dragging) { ry += 0.4; apply(); } }, 32);
          root.addEventListener('pointerdown', function (e) {
            dragging = true; lx = e.clientX; ly = e.clientY; root.setPointerCapture(e.pointerId);
          });
          root.addEventListener('pointermove', function (e) {
            if (!dragging) return;
            ry += (e.clientX - lx) * 0.4;
            rx = Math.max(-30, Math.min(30, rx + (e.clientY - ly) * 0.3));
            lx = e.clientX; ly = e.clientY; apply();
          });
          const stop = function () { dragging = false; };
          root.addEventListener('pointerup', stop);
          root.addEventListener('pointercancel', stop);
        })();
      </script>
    </div>`
      }

      // Preview uses short-lived blob: URLs from conversion; export rewrites these.
      const posterAttr = model.posterDataUrl ? ` poster="${model.posterDataUrl}"` : ''
      const mvId = `mv-${model.id}`

      return `
    <div class="slide-3d-model" data-model="${model.id}" title="${model.name}" style="${box}">
      <model-viewer
        id="${mvId}"
        src="${model.dataUrl}"
        alt="${model.name}"
        ${posterAttr}
        camera-controls
        touch-action="pan-y"
        shadow-intensity="0.85"
        shadow-softness="0.8"
        exposure="1.05"
        environment-image="neutral"
        tone-mapping="commerce"
        style="width:100%;height:100%;background:transparent;"
      >
        ${
          model.posterDataUrl
            ? `<img slot="poster" src="${model.posterDataUrl}" alt="${model.name} preview" class="slide-3d-poster" />`
            : ''
        }
      </model-viewer>
    </div>`
    })
    .join('\n')
}

/**
 * file:// pages cannot fetch sibling .glb files (CORS). Embed each GLB as base64 and
 * assign a blob: URL at runtime — works when the HTML is opened directly.
 */
export const buildInlineGlbLoaderScripts = (models: Slide3DModel[]): string => {
  const loaders = models
    .filter((m) => !m.posterOnly && m.glbBytes && m.glbBytes.byteLength > 0)
    .map((model) => {
      const mvId = `mv-${model.id}`
      const b64 = uint8ToBase64(model.glbBytes!)
      return `<script>
(function () {
  var el = document.getElementById(${JSON.stringify(mvId)});
  if (!el) return;
  var b64 = ${JSON.stringify(b64)};
  var bin = atob(b64);
  var bytes = new Uint8Array(bin.length);
  for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  el.src = URL.createObjectURL(new Blob([bytes], { type: 'model/gltf-binary' }));
})();
</script>`
    })

  return loaders.join('\n')
}

/** Strip blob:/media src so inline loaders own model-viewer.src on file:// opens. */
export const stripModelViewerExternalSrc = (html: string): string =>
  html.replace(
    /(<model-viewer\b[^>]*?)\s+src="(?:blob:[^"]+|media\/[^"]+|data:model\/[^"]+)"/gi,
    '$1'
  )

export const collectExportableGlbFiles = (
  models: Slide3DModel[]
): Array<{ path: string; bytes: Uint8Array }> =>
  models
    .filter((m) => !m.posterOnly && m.glbBytes && m.exportPath)
    .map((m) => ({ path: m.exportPath!, bytes: m.glbBytes! }))

export const modelViewerHeadSnippet = `
    <script type="module" src="https://ajax.googleapis.com/ajax/libs/model-viewer/3.5.0/model-viewer.min.js"></script>
`

export const modelViewerCss = `
      .slide-3d-model {
        overflow: hidden;
        border-radius: 4px;
      }
      .slide-3d-model model-viewer {
        display: block;
        width: 100%;
        height: 100%;
        --poster-color: transparent;
      }
      .slide-3d-poster {
        width: 100%;
        height: 100%;
        object-fit: contain;
      }
      .slide-3d-poster-orbit {
        background: transparent;
      }
      .poster-orbit-scene {
        width: 100%;
        height: 100%;
        perspective: 900px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: grab;
        touch-action: none;
        user-select: none;
      }
      .poster-orbit-scene:active {
        cursor: grabbing;
      }
      .poster-orbit-img {
        max-width: 100%;
        max-height: 100%;
        object-fit: contain;
        transform-style: preserve-3d;
        will-change: transform;
        filter: drop-shadow(0 8px 18px rgba(0,0,0,0.25));
      }
`

export const extractSlideImagesHtml = async (
  zip: JSZip,
  slidePath: string,
  slideDoc: Document,
  skipEmbedIds: Set<string> = new Set(),
  slideSize: SlideSize = DEFAULT_SLIDE_SIZE
): Promise<string> => {
  const rels = await loadRelationships(zip, slidePath)
  const parts: string[] = []
  const used = new Set<string>()

  const isUnder = (el: Element | null, names: string[]): boolean => {
    let cur: Element | null = el
    while (cur) {
      if (names.includes(localName(cur))) return true
      cur = cur.parentElement
    }
    return false
  }

  // Only real picture shapes — never re-export slide/master background blips as <img>
  for (const pic of elementsByLocal(slideDoc, 'pic')) {
    if (isUnder(pic, ['bg', 'bgPr', 'bgRef'])) continue

    if (isUnder(pic, ['AlternateContent'])) {
      let el: Element | null = pic
      let skipModel = false
      while (el) {
        if (localName(el) === 'AlternateContent') {
          skipModel =
            elementsByLocal(el, 'model3d').length > 0 ||
            elementsByLocal(el, 'model3D').length > 0 ||
            elementsByLocal(el, 'Choice').some((c) =>
              (c.getAttribute('Requires') || '').toLowerCase().includes('am3d')
            )
          break
        }
        el = el.parentElement
      }
      if (skipModel) continue
    }

    if (isUnder(pic, ['graphicFrame'])) {
      const frame = (() => {
        let el: Element | null = pic
        while (el) {
          if (localName(el) === 'graphicFrame') return el
          el = el.parentElement
        }
        return null
      })()
      if (frame) {
        const uri = (firstByLocal(frame, 'graphicData')?.getAttribute('uri') || '').toLowerCase()
        if (uri.includes('diagram') || elementsByLocal(frame, 'relIds').length > 0) continue
      }
    }

    const blip = firstByLocal(pic, 'blip')
    const embed = getEmbedId(blip)
    if (!embed || skipEmbedIds.has(embed) || !rels.has(embed)) continue
    const target = rels.get(embed)!.target
    if (!isImagePath(target) || used.has(`${target}@${parts.length}`)) continue

    const dataUrl = await toDataUrl(zip, target)
    if (!dataUrl) continue
    used.add(target)

    const box = readTransform(pic, slideSize)
    const name = firstByLocal(pic, 'cNvPr')?.getAttribute('name') || `Slide image ${parts.length + 1}`
    parts.push(
      `    <img src="${dataUrl}" alt="${name.replace(/"/g, '')}" class="slide-image" style="position:absolute;left:${box.leftPx.toFixed(1)}px;top:${box.topPx.toFixed(1)}px;width:${box.widthPx.toFixed(1)}px;height:${box.heightPx.toFixed(1)}px;object-fit:fill;margin:0;max-width:none;max-height:none;z-index:2;" />\n`
    )
  }

  return parts.join('')
}

export { loadXml }
