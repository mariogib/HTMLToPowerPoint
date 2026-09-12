import React, { useState, useCallback } from 'react'
import { useDropzone, type DropEvent, type FileRejection } from 'react-dropzone'
import JSZip from 'jszip'
import { saveAs } from 'file-saver'
import ProgressBar from './ProgressBar'
import FolderSelector from './FolderSelector'
import SlideRenderer from './SlideRenderer'
import {
  backgroundCssForSlide,
  backgroundInlineStyle,
  resolveSlideBackground,
  type SlideBackground,
} from '../utils/pptxBackground'
import {
  buildInlineGlbLoaderScripts,
  extractSlide3DModels,
  extractSlideImagesHtml,
  modelViewerCss,
  modelViewerHeadSnippet,
  render3DModelsHtml,
  stripModelViewerExternalSrc,
  type Slide3DModel,
} from '../utils/pptxModels'
import {
  collectSmartArtImageEmbedIds,
  extractSlideSmartArt,
  renderSmartArtHtml,
  smartArtCss,
} from '../utils/pptxSmartArt'
import {
  extractSlideTextShapes,
  renderTextShapesHtml,
  textShapesCss,
} from '../utils/pptxText'
import {
  DEFAULT_SLIDE_SIZE,
  readPresentationSlideSize,
  type SlideSize,
} from '../utils/pptxSlideSize'
import './PowerPointConverter.css'

interface Slide {
  id: string
  content: string
  styles: string
  notes?: string
  models3d?: Slide3DModel[]
}

interface ConversionResult {
  slides: Slide[]
  metadata: {
    title: string
    slideCount: number
    convertedAt: string
    slideWidthPx: number
    slideHeightPx: number
  }
}

const isPowerPointFile = (file: File) => {
  const name = file.name.toLowerCase()
  return name.endsWith('.pptx') || name.endsWith('.ppt')
}

const PowerPointConverter: React.FC = () => {
  const [files, setFiles] = useState<File[]>([])
  const [isConverting, setIsConverting] = useState(false)
  const [progress, setProgress] = useState(0)
  const [convertedData, setConvertedData] = useState<ConversionResult | null>(null)
  const [downloadAcknowledged, setDownloadAcknowledged] = useState(false)
  const [error, setError] = useState<string>('')

  const selectPowerPointFiles = useCallback((incoming: File[]) => {
    const pptxFiles = incoming.filter(isPowerPointFile)

    if (pptxFiles.length === 0) {
      setError('Please select a valid PowerPoint file (.pptx or .ppt)')
      return
    }

    setFiles([pptxFiles[0]])
    setError('')
    setConvertedData(null)
  }, [])

  const onDrop = useCallback((acceptedFiles: File[], fileRejections: FileRejection[]) => {
    // Prefer accepted files; if MIME mismatch rejected a valid .pptx/.ppt (common on Windows),
    // still accept by extension.
    const rejectedButValid = fileRejections
      .map((rejection) => rejection.file)
      .filter(isPowerPointFile)

    selectPowerPointFiles([...acceptedFiles, ...rejectedButValid])
  }, [selectPowerPointFiles])

  const onDropRejected = useCallback((fileRejections: FileRejection[]) => {
    if (fileRejections.some((rejection) => isPowerPointFile(rejection.file))) {
      return
    }
    setError('Please select a valid PowerPoint file (.pptx or .ppt)')
  }, [])

  // Read files synchronously from the drop event. Avoids Chromium/Windows failures where
  // awaiting getAsFileSystemHandle() clears DataTransferItem before getAsFile() can run.
  const getFilesFromEvent = useCallback(async (event: DropEvent | FileSystemFileHandle[]): Promise<Array<File | DataTransferItem>> => {
    if (Array.isArray(event)) {
      return Promise.all(event.map((handle) => handle.getFile()))
    }

    if ('dataTransfer' in event && event.dataTransfer) {
      const { files, items } = event.dataTransfer
      if (files?.length) {
        return Array.from(files)
      }
      if (items?.length) {
        return Array.from(items)
          .filter((item): item is DataTransferItem => item.kind === 'file')
          .map((item) => item.getAsFile())
          .filter((file): file is File => file !== null)
      }
    }

    if ('target' in event && event.target instanceof HTMLInputElement && event.target.files) {
      return Array.from(event.target.files)
    }

    return []
  }, [])

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    onDropRejected,
    getFilesFromEvent,
    accept: {
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
      'application/vnd.ms-powerpoint': ['.ppt'],
      'application/octet-stream': ['.pptx', '.ppt'],
      'application/zip': ['.pptx'],
      'application/x-zip-compressed': ['.pptx']
    },
    multiple: false,
    noClick: false,
    useFsAccessApi: false
  })

  const extractPowerPointContent = async (file: File): Promise<ConversionResult> => {
    const zip = new JSZip()
    const content = await file.arrayBuffer()
    const zipFile = await zip.loadAsync(content)
    const slideSize = await readPresentationSlideSize(zipFile)
    
    console.log('PowerPoint file structure:', Object.keys(zipFile.files))
    
    // Extract slides
    const slides: Slide[] = []
    const slideFiles = Object.keys(zipFile.files).filter(name => 
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    )
    
    console.log('Found slide files:', slideFiles)
    
    // Sort slide files numerically
    slideFiles.sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || '0')
      const numB = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || '0')
      return numA - numB
    })

    let processedSlides = 0
    const totalSlides = slideFiles.length

    for (const slideFile of slideFiles) {
      try {
        const slideXml = await zipFile.files[slideFile].async('text')
        const slideNumber = slideFile.match(/slide(\d+)\.xml/)?.[1] || '1'
        
        console.log(`Processing slide ${slideNumber}:`)
        console.log('Raw XML length:', slideXml.length)
        console.log('XML preview:', slideXml.substring(0, 500))
        
        // Additional debugging: dump XML structure
        const parser = new DOMParser()
        const slideDoc = parser.parseFromString(slideXml, 'text/xml')
        
        console.log('=== SLIDE XML STRUCTURE DEBUG ===')
        console.log('All element types found:', Array.from(new Set(Array.from(slideDoc.getElementsByTagName('*')).map(el => el.tagName))))
        
        // Check for common PowerPoint elements
        const commonElements = ['p:sp', 'a:t', 'a:p', 'p:txBody', 'p:ph', 'a:r', 'a:rPr']
        commonElements.forEach(tagName => {
          const elements = slideDoc.getElementsByTagName(tagName)
          console.log(`${tagName}: ${elements.length} found`)
          if (elements.length > 0 && elements.length <= 5) {
            for (let i = 0; i < elements.length; i++) {
              console.log(`  ${tagName}[${i}] text:`, elements[i].textContent?.substring(0, 100))
            }
          }
        })
        console.log('=== END DEBUG ===')
        
        // Parse slide content and convert to HTML
        const background = await resolveSlideBackground(zipFile, slideFile, slideDoc, slideSize)
        const { html: htmlContent, models3d } = await convertSlideToHTML(
          slideXml,
          zipFile,
          slideNumber,
          background,
          slideFile,
          slideSize
        )
        const styles = await extractSlideStyles(slideNumber, background, slideSize)
        
        slides.push({
          id: `slide-${slideNumber}`,
          content: htmlContent,
          styles: styles,
          models3d,
        })
        
        processedSlides++
        setProgress((processedSlides / totalSlides) * 100)
        
        // Add small delay to show progress
        await new Promise(resolve => setTimeout(resolve, 100))
      } catch (error) {
        console.error(`Error processing slide ${slideFile}:`, error)
        
        // Add a fallback slide with error information
        const slideNumber = slideFile.match(/slide(\d+)\.xml/)?.[1] || '1'
        slides.push({
          id: `slide-${slideNumber}`,
          content: `
            <div class="slide slide-${slideNumber}" id="slide-${slideNumber}" data-slide="${slideNumber}">
              <div class="slide-content">
                <div class="error-content">
                  <h3>Error Processing Slide ${slideNumber}</h3>
                  <p>Unable to extract content from this slide.</p>
                  <p>Error: ${error instanceof Error ? error.message : 'Unknown error'}</p>
                </div>
              </div>
            </div>
          `,
          styles: await extractSlideStyles(slideNumber, {
            type: 'solid',
            color: '#ffffff',
            css: 'background-color: #ffffff; background-image: none;',
          }, slideSize)
        })
        
        processedSlides++
        setProgress((processedSlides / totalSlides) * 100)
      }
    }

    return {
      slides,
      metadata: {
        title: file.name.replace(/\.(pptx?|ppt)$/i, ''),
        slideCount: slides.length,
        convertedAt: new Date().toISOString(),
        slideWidthPx: slideSize.widthPx,
        slideHeightPx: slideSize.heightPx,
      }
    }
  }

  const convertSlideToHTML = async (
    slideXml: string,
    zip: JSZip,
    slideNumber: string,
    background: SlideBackground,
    slidePath: string,
    slideSize: SlideSize
  ): Promise<{ html: string; models3d: Slide3DModel[] }> => {
    // Parse XML and extract text and shape content
    const parser = new DOMParser()
    const xmlDoc = parser.parseFromString(slideXml, 'text/xml')
    const bgStyle = backgroundInlineStyle(background)
    
    let htmlContent = `<div class="slide slide-${slideNumber}" id="slide-${slideNumber}" data-slide="${slideNumber}" style="${bgStyle}">\n`
    let hasContent = false
    
    console.log(`Processing slide ${slideNumber}`)

    // Positioned text shapes with inherited layout/master typography
    try {
      const textShapes = await extractSlideTextShapes(zip, slidePath, xmlDoc, slideSize)
      if (textShapes.length > 0) {
        htmlContent += '  <div class="slide-text-layer">\n'
        htmlContent += renderTextShapesHtml(textShapes)
        htmlContent += '  </div>\n'
        hasContent = true
      }
    } catch (error) {
      console.error('Error extracting slide text:', error)
    }

    // Extract SmartArt diagrams (image fallback and/or node text layout)
    try {
      const smartArts = await extractSlideSmartArt(zip, slidePath, xmlDoc, slideSize)
      if (smartArts.length > 0) {
        htmlContent += '  <div class="slide-media-layer">\n'
        htmlContent += renderSmartArtHtml(smartArts)
        htmlContent += '  </div>\n'
        hasContent = true
      }
    } catch (error) {
      console.error('Error extracting SmartArt:', error)
    }

    // Extract embedded 3D models (.glb) with interactive model-viewer + poster fallback
    let models3d: Slide3DModel[] = []
    try {
      models3d = await extractSlide3DModels(zip, slidePath, xmlDoc, slideSize)
      if (models3d.length > 0) {
        htmlContent += '  <div class="slide-media-layer">\n'
        htmlContent += render3DModelsHtml(models3d)
        htmlContent += '  </div>\n'
        hasContent = true
      }
    } catch (error) {
      console.error('Error extracting 3D models:', error)
    }

    // Extract images via relationship ids (skip SmartArt/3D posters handled above)
    try {
      const skipEmbeds = collectSmartArtImageEmbedIds(xmlDoc)
      const imagesHtml = await extractSlideImagesHtml(zip, slidePath, xmlDoc, skipEmbeds, slideSize)
      if (imagesHtml) {
        htmlContent += imagesHtml
        hasContent = true
      }
    } catch (error) {
      console.error('Error extracting images:', error)
    }
    
    htmlContent += '</div>\n'
    console.log(`Generated HTML for slide ${slideNumber} (hasContent: ${hasContent}):`, htmlContent.slice(0, 500))
    return { html: htmlContent, models3d }
  }

  const extractSlideStyles = async (
    slideNumber: string,
    background: SlideBackground,
    slideSize: SlideSize = DEFAULT_SLIDE_SIZE
  ): Promise<string> => {
    const styles = `
      .slide {
        width: ${slideSize.widthPx}px;
        height: ${slideSize.heightPx}px;
        position: relative;
        border: 1px solid #ddd;
        margin: 20px auto;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        overflow: hidden;
        font-family: 'Calibri', sans-serif;
        padding: 0;
        box-sizing: border-box;
      }

      ${backgroundCssForSlide(slideNumber, background)}

      ${modelViewerCss}
      ${smartArtCss}
      ${textShapesCss}

      .slide-media-layer {
        position: absolute;
        inset: 0;
        pointer-events: none;
      }

      .slide-text-layer {
        position: absolute;
        inset: 0;
        pointer-events: none;
        z-index: 3;
      }

      .slide-media-layer .slide-3d-model,
      .slide-media-layer .smartart {
        pointer-events: auto;
      }
      
      .slide-content {
        width: 100%;
        height: 100%;
        position: relative;
        z-index: 1;
        padding: 20px;
        box-sizing: border-box;
      }
      
      .text-element {
        margin: 10px 0;
        line-height: 1.4;
      }
      
      .text-element p {
        margin: 5px 0;
        font-size: 16px;
        color: #333;
      }
      
      .shape-element {
        margin: 15px 0;
        padding: 8px;
        background-color: rgba(240, 240, 240, 0.3);
        border-radius: 4px;
      }
      
      .shape-element p {
        margin: 0;
        font-size: 16px;
        color: #333;
      }
      
      .slide-title {
        font-size: 24px;
        font-weight: bold;
        color: #1a1a1a;
        margin-bottom: 20px;
        text-align: center;
      }
      
      .slide-title p {
        margin: 0;
        font-size: 24px;
      }
      
      .slide-content-ph {
        margin: 15px 0;
      }
      
      .slide-content-ph p {
        font-size: 18px;
        line-height: 1.5;
      }
      
      .no-content {
        text-align: center;
        color: #666;
        font-style: italic;
        margin-top: 50px;
      }
      
      .no-content p {
        margin: 10px 0;
      }
      
      .slide-image {
        display: block;
        border: 0;
        box-sizing: border-box;
      }
      
      .error-content {
        text-align: center;
        color: #dc3545;
        padding: 20px;
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 8px;
        margin: 20px;
      }
      
      .error-content h3 {
        margin: 0 0 15px 0;
        color: #721c24;
      }
      
      .error-content p {
        margin: 5px 0;
        color: #721c24;
      }
      
      .fallback-content {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        padding: 15px;
        border-radius: 6px;
        margin: 20px;
      }
      
      .fallback-content p {
        margin: 8px 0;
        color: #856404;
      }
      
      .fallback-content strong {
        color: #533f03;
      }
    `

    return styles
  }

  const handleConvert = async () => {
    if (files.length === 0) {
      setError('Please select a PowerPoint file first')
      return
    }

    if (!downloadAcknowledged) {
      setError('Please acknowledge the download location')
      return
    }

    setIsConverting(true)
    setProgress(0)
    setError('')

    try {
      const result = await extractPowerPointContent(files[0])
      setConvertedData(result)
      
      // Generate HTML file
      await generateHTMLFile(result)
      
    } catch (error) {
      console.error('Conversion error:', error)
      setError('Failed to convert PowerPoint file. Please ensure it\'s a valid .pptx file.')
    } finally {
      setIsConverting(false)
      setProgress(0)
    }
  }

  const generateHTMLFile = async (data: ConversionResult) => {
    const htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${data.metadata.title}</title>
    ${modelViewerHeadSnippet}
    <style>
        body {
            margin: 0;
            padding: 20px;
            font-family: 'Calibri', 'Arial', sans-serif;
            background-color: #f5f5f5;
        }
        
        .presentation-container {
            max-width: ${Math.max(data.metadata.slideWidthPx + 40, 1000)}px;
            margin: 0 auto;
        }
        
        .presentation-header {
            text-align: center;
            margin-bottom: 30px;
            padding: 20px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .navigation {
            text-align: center;
            margin: 20px 0;
        }
        
        .nav-button {
            background: #007acc;
            color: white;
            border: none;
            padding: 10px 20px;
            margin: 0 5px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
        }
        
        .nav-button:hover {
            background: #005a9e;
        }
        
        .nav-button:disabled {
            background: #ccc;
            cursor: not-allowed;
        }
        
        .slide-counter {
            margin: 10px 0;
            font-weight: bold;
        }
        
        ${data.slides.map(slide => slide.styles).join('\n')}
        
        /* Override any display settings from individual slide styles */
        .slide {
            display: none !important;
        }
        
        .slide.active {
            display: block !important;
        }
    </style>
</head>
<body>
    <div class="presentation-container">
        <div class="presentation-header">
            <h1>${data.metadata.title}</h1>
            <p>Converted from PowerPoint \u2022 ${data.metadata.slideCount} slides</p>
            <p class="slide-counter">Slide <span id="current-slide">1</span> of ${data.metadata.slideCount}</p>
        </div>
        
        <div class="navigation">
            <button class="nav-button" id="prev-btn" onclick="previousSlide()">\u2190 Previous</button>
            <button class="nav-button" id="next-btn" onclick="nextSlide()">Next \u2192</button>
        </div>
        
        <div class="slides-container">
            ${data.slides.map((slide, index) => {
              // slide.content already includes the .slide wrapper with data-slide
              const content = slide.content.replace(
                /class="slide /,
                `class="slide ${index === 0 ? 'active ' : ''}`
              )
              return content
            }).join('')}
        </div>
        
        <div class="navigation">
            <button class="nav-button" id="prev-btn-bottom" onclick="previousSlide()">\u2190 Previous</button>
            <button class="nav-button" id="next-btn-bottom" onclick="nextSlide()">Next \u2192</button>
        </div>
    </div>
    
    <script>
        let currentSlide = 1;
        const totalSlides = ${data.metadata.slideCount};
        
        function showSlide(slideNumber) {
            document.querySelectorAll('.slide').forEach(slide => {
                slide.classList.remove('active');
            });
            
            const slide = document.querySelector('[data-slide="' + slideNumber + '"]');
            if (slide) {
                slide.classList.add('active');
            }
            
            document.getElementById('current-slide').textContent = slideNumber;
            
            // Update button states
            const prevButtons = document.querySelectorAll('#prev-btn, #prev-btn-bottom');
            const nextButtons = document.querySelectorAll('#next-btn, #next-btn-bottom');
            
            prevButtons.forEach(btn => btn.disabled = slideNumber === 1);
            nextButtons.forEach(btn => btn.disabled = slideNumber === totalSlides);
        }
        
        function nextSlide() {
            if (currentSlide < totalSlides) {
                currentSlide++;
                showSlide(currentSlide);
            }
        }
        
        function previousSlide() {
            if (currentSlide > 1) {
                currentSlide--;
                showSlide(currentSlide);
            }
        }
        
        // Keyboard navigation
        document.addEventListener('keydown', function(event) {
            if (event.key === 'ArrowRight' || event.key === ' ') {
                nextSlide();
            } else if (event.key === 'ArrowLeft') {
                previousSlide();
            }
        });
        
        // Initialize
        showSlide(1);
    </script>
</body>
</html>`

    // file:// cannot fetch sibling .glb files (browser CORS). Embed GLBs as
    // base64 -> blob: URLs so double-clicking the HTML works offline.
    const allModels = data.slides.flatMap((s) => s.models3d || [])
    let exportHtml = stripModelViewerExternalSrc(htmlContent)
    const inlineLoaders = buildInlineGlbLoaderScripts(allModels)
    if (inlineLoaders) {
      exportHtml = exportHtml.replace('</body>', `${inlineLoaders}\n</body>`)
    }

    const blob = new Blob([exportHtml], { type: 'text/html;charset=utf-8' })
    saveAs(blob, `${data.metadata.title}.html`)
  }

  const createTestSlides = () => {
    const testSlides: Slide[] = [
      {
        id: 'test-slide-1',
        content: `
          <div class="slide slide-1" id="slide-1">
            <div class="slide-content">
              <div class="slide-title">
                <p>Test Slide 1</p>
              </div>
              <div class="text-element">
                <p>This is a test slide to verify the conversion process.</p>
                <p>If you can see this, the basic slide rendering is working.</p>
              </div>
            </div>
          </div>
        `,
        styles: `
          .slide { width: ${DEFAULT_SLIDE_SIZE.widthPx}px; height: ${DEFAULT_SLIDE_SIZE.heightPx}px; background: white; border: 1px solid #ddd; margin: 20px auto; padding: 0; box-sizing: border-box; }
          .slide-content { padding: 20px; box-sizing: border-box; }
          .slide-title { font-size: 24px; font-weight: bold; margin-bottom: 20px; text-align: center; }
          .text-element { margin: 15px 0; }
          .text-element p { font-size: 16px; margin: 8px 0; }
        `
      },
      {
        id: 'test-slide-2',
        content: `
          <div class="slide slide-2" id="slide-2">
            <div class="slide-content">
              <div class="slide-title">
                <p>Test Slide 2</p>
              </div>
              <div class="text-element">
                <p>This is the second test slide.</p>
                <p>It demonstrates multiple text elements and formatting.</p>
              </div>
              <div class="shape-element">
                <p>This is a shape element with different styling.</p>
              </div>
            </div>
          </div>
        `,
        styles: `
          .slide { width: ${DEFAULT_SLIDE_SIZE.widthPx}px; height: ${DEFAULT_SLIDE_SIZE.heightPx}px; background: white; border: 1px solid #ddd; margin: 20px auto; padding: 0; box-sizing: border-box; }
          .slide-content { padding: 20px; box-sizing: border-box; }
          .slide-title { font-size: 24px; font-weight: bold; margin-bottom: 20px; text-align: center; }
          .text-element { margin: 15px 0; }
          .text-element p { font-size: 16px; margin: 8px 0; }
          .shape-element { background: #f0f0f0; padding: 10px; border-radius: 4px; margin: 15px 0; }
        `
      }
    ]

    const testData: ConversionResult = {
      slides: testSlides,
      metadata: {
        title: 'Test-Presentation',
        slideCount: testSlides.length,
        convertedAt: new Date().toISOString(),
        slideWidthPx: DEFAULT_SLIDE_SIZE.widthPx,
        slideHeightPx: DEFAULT_SLIDE_SIZE.heightPx,
      }
    }

    setConvertedData(testData)
    generateHTMLFile(testData)
  }

  return (
    <div className="powerpoint-converter">
      <div className="upload-section">
        <div
          {...getRootProps()}
          className={`dropzone ${isDragActive ? 'active' : ''}`}
        >
          <input {...getInputProps()} />
          {isDragActive ? (
            <div className="drop-message">
              <h3>Drop your PowerPoint file here...</h3>
            </div>
          ) : (
            <div className="drop-message">
              <h3>Drag & drop a PowerPoint file here</h3>
              <p>or click to select a file</p>
              <small>Supported formats: .pptx, .ppt</small>
            </div>
          )}
        </div>
        
        {files.length > 0 && (
          <div className="file-info">
            <h4>Selected file:</h4>
            <p>{files[0].name} ({(files[0].size / 1024 / 1024).toFixed(2)} MB)</p>
          </div>
        )}
      </div>

      <FolderSelector onFolderSelect={() => setDownloadAcknowledged(true)} selectedFolder="" />

      {downloadAcknowledged && (
        <div className="download-acknowledged">
          <span className="check-icon">âœ…</span>
          <p>Download location acknowledged. Ready to convert!</p>
        </div>
      )}

      {error && (
        <div className="error-message">
          <p>{error}</p>
        </div>
      )}

      <div className="convert-section">
        <button
          onClick={handleConvert}
          disabled={files.length === 0 || !downloadAcknowledged || isConverting}
          className="convert-button"
        >
          {isConverting ? 'Converting...' : 'Convert to HTML'}
        </button>
        
        <button
          onClick={() => createTestSlides()}
          className="test-button"
          style={{ marginLeft: '10px' }}
        >
          Create Test Slides
        </button>
      </div>

      {isConverting && (
        <ProgressBar progress={progress} />
      )}

      {convertedData && (
        <div className="conversion-result">
          <h3>Conversion Complete!</h3>
          <p>Successfully converted {convertedData.metadata.slideCount} slides</p>
          <p><strong>The HTML file has been downloaded to your Downloads folder.</strong></p>
          <p>Look for: <code>{convertedData.metadata.title}.html</code></p>
          
          <div className="debug-info">
            <details>
              <summary>Debug Information</summary>
              <p>Open browser console (F12) to see detailed extraction logs</p>
              <p>Slides processed: {convertedData.slides.length}</p>
            </details>
          </div>
          
          <div className="preview-section">
            <h4>Preview:</h4>
            <SlideRenderer slides={convertedData.slides} />
          </div>
        </div>
      )}
    </div>
  )
}

export default PowerPointConverter
