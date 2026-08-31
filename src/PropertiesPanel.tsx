import type { CanvasElement } from './lib/canvasTypes'
import { BULLET_STYLES } from './lib/bulletStyles'
import { containRect } from './lib/imageFit'
import { defaultColor, defaultFontSize } from './lib/measureElement'

const FONT_FACES = ['Arial', 'Calibri', 'Georgia', 'Times New Roman', 'Verdana', 'Inter']

type PropertiesPanelProps = {
  element: CanvasElement
  x: number
  y: number
  onChange: (updates: Partial<CanvasElement>) => void
  onClose: () => void
  onDelete: () => void
}

export const PropertiesPanel = ({
  element,
  x,
  y,
  onChange,
  onClose,
  onDelete,
}: PropertiesPanelProps) => {
  const fontSize = element.fontSize ?? defaultFontSize(element.type)
  const color = element.color ?? defaultColor(element.type)
  const fontFamily = element.fontFamily ?? 'Arial'
  const left = Math.min(x, Math.max(12, window.innerWidth - 300))
  const top = Math.min(y, Math.max(12, window.innerHeight - 480))
  const isText = element.type === 'text' || element.type === 'heading' || element.type === 'list'
  const imageRatio = element.aspectRatio || element.width / Math.max(element.height, 1)
  const title =
    element.type === 'heading'
      ? 'Heading properties'
      : element.type === 'text'
        ? 'Text box properties'
        : element.type === 'list'
          ? 'List properties'
          : 'Image properties'

  return (
    <div
      className="properties-panel"
      style={{ left, top }}
      role="dialog"
      aria-label={title}
      onMouseDown={(event) => event.stopPropagation()}
    >
      <div className="properties-panel__header">
        <h3>{title}</h3>
        <button type="button" className="modal-close" onClick={onClose} aria-label="Close properties">
          ✕
        </button>
      </div>

      <div className="properties-panel__body">
        {isText && (
          <div className="field">
            <label htmlFor={`prop-content-${element.id}`}>Text</label>
            <textarea
              id={`prop-content-${element.id}`}
              className="properties-textarea"
              rows={element.type === 'list' ? 5 : 3}
              value={element.content}
              onChange={(event) => onChange({ content: event.target.value })}
            />
          </div>
        )}

        {element.type === 'list' && (
          <div className="field">
            <span className="field-label-inline" id={`prop-bullet-${element.id}`}>
              Bullet style
            </span>
            <div className="bullet-choices" role="group" aria-labelledby={`prop-bullet-${element.id}`}>
              {BULLET_STYLES.map((style) => (
                <button
                  key={style.id}
                  type="button"
                  className={`bullet-choice${(element.bulletStyle ?? 'disc') === style.id ? ' bullet-choice--active' : ''}`}
                  title={style.label}
                  aria-label={style.label}
                  aria-pressed={(element.bulletStyle ?? 'disc') === style.id}
                  onClick={() => onChange({ bulletStyle: style.id })}
                >
                  {style.preview}
                </button>
              ))}
            </div>
          </div>
        )}

        {element.type === 'image' && (
          <div className="field">
            <label htmlFor={`prop-image-${element.id}`}>Replace image</label>
            <input
              id={`prop-image-${element.id}`}
              type="file"
              accept="image/*"
              onChange={(event) => {
                const file = event.target.files?.[0]
                if (!file) return
                const reader = new FileReader()
                reader.onload = () => {
                  if (typeof reader.result !== 'string') return
                  const dataUrl = reader.result
                  const img = new Image()
                  img.onload = () => {
                    const natW = img.naturalWidth || img.width
                    const natH = img.naturalHeight || img.height
                    const fitted = containRect(
                      element.x,
                      element.y,
                      element.width,
                      element.height,
                      natW,
                      natH
                    )
                    onChange({
                      content: dataUrl,
                      x: fitted.x,
                      y: fitted.y,
                      width: fitted.width,
                      height: fitted.height,
                      aspectRatio: natW / natH,
                    })
                  }
                  img.src = dataUrl
                }
                reader.readAsDataURL(file)
              }}
            />
          </div>
        )}

        {isText && (
          <>
            <div className="field-row">
              <div className="field">
                <label htmlFor={`prop-font-${element.id}`}>Font</label>
                <select
                  id={`prop-font-${element.id}`}
                  value={fontFamily}
                  onChange={(event) => onChange({ fontFamily: event.target.value })}
                >
                  {FONT_FACES.map((face) => (
                    <option key={face} value={face}>
                      {face}
                    </option>
                  ))}
                </select>
              </div>
              <div className="field">
                <label htmlFor={`prop-size-${element.id}`}>Size</label>
                <input
                  id={`prop-size-${element.id}`}
                  type="number"
                  min={8}
                  max={96}
                  value={fontSize}
                  onChange={(event) => onChange({ fontSize: Number(event.target.value) })}
                />
              </div>
            </div>

            <div className="field">
              <label htmlFor={`prop-color-${element.id}`}>Color</label>
              <input
                id={`prop-color-${element.id}`}
                type="color"
                value={color.startsWith('#') ? color : `#${color}`}
                onChange={(event) => onChange({ color: event.target.value })}
              />
            </div>

            <div className="properties-toggles">
              <label className="checkbox">
                <input
                  type="checkbox"
                  checked={element.bold ?? element.type === 'heading'}
                  onChange={(event) => onChange({ bold: event.target.checked })}
                />
                Bold
              </label>
              <label className="checkbox">
                <input
                  type="checkbox"
                  checked={Boolean(element.italic)}
                  onChange={(event) => onChange({ italic: event.target.checked })}
                />
                Italic
              </label>
            </div>

            <div className="field">
              <span className="field-label-inline">Align</span>
              <div className="align-buttons">
                {(['left', 'center', 'right'] as const).map((align) => (
                  <button
                    key={align}
                    type="button"
                    className={`align-btn${(element.align ?? 'left') === align ? ' align-btn--active' : ''}`}
                    onClick={() => onChange({ align })}
                  >
                    {align}
                  </button>
                ))}
              </div>
            </div>

            <label className="checkbox">
              <input
                type="checkbox"
                checked={element.autoSize !== false}
                onChange={(event) => onChange({ autoSize: event.target.checked })}
              />
              Auto-fit to content
            </label>
          </>
        )}

        <div className="field-row">
          <div className="field">
            <label htmlFor={`prop-x-${element.id}`}>X</label>
            <input
              id={`prop-x-${element.id}`}
              type="number"
              value={Math.round(element.x)}
              onChange={(event) => onChange({ x: Number(event.target.value) })}
            />
          </div>
          <div className="field">
            <label htmlFor={`prop-y-${element.id}`}>Y</label>
            <input
              id={`prop-y-${element.id}`}
              type="number"
              value={Math.round(element.y)}
              onChange={(event) => onChange({ y: Number(event.target.value) })}
            />
          </div>
        </div>

        <div className="field-row">
          <div className="field">
            <label htmlFor={`prop-w-${element.id}`}>Width</label>
            <input
              id={`prop-w-${element.id}`}
              type="number"
              min={40}
              value={Math.round(element.width)}
              onChange={(event) => {
                const width = Number(event.target.value)
                if (element.type === 'image') {
                  onChange({
                    width,
                    height: Math.max(28, width / imageRatio),
                    autoSize: false,
                  })
                  return
                }
                onChange({ width, autoSize: false })
              }}
            />
          </div>
          <div className="field">
            <label htmlFor={`prop-h-${element.id}`}>Height</label>
            <input
              id={`prop-h-${element.id}`}
              type="number"
              min={28}
              value={Math.round(element.height)}
              onChange={(event) => {
                const height = Number(event.target.value)
                if (element.type === 'image') {
                  onChange({
                    height,
                    width: Math.max(40, height * imageRatio),
                    autoSize: false,
                  })
                  return
                }
                onChange({ height, autoSize: false })
              }}
            />
          </div>
        </div>

        <button type="button" className="properties-delete" onClick={onDelete}>
          Delete control
        </button>
      </div>
    </div>
  )
}
