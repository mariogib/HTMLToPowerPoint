import { useState } from 'react'

type SlidePreview = {
  id: string
  html: string
}

type SlideRailProps = {
  slides: SlidePreview[]
  activeId: string
  background: string
  slideWidthPx: number
  slideHeightPx: number
  onSelect: (id: string) => void
  onAdd: () => void
  onDelete: (id: string) => void
  onMove: (id: string, direction: -1 | 1) => void
  onReorder: (fromId: string, toId: string) => void
}

const THUMB_WIDTH = 164

export const SlideRail = ({
  slides,
  activeId,
  background,
  slideWidthPx,
  slideHeightPx,
  onSelect,
  onAdd,
  onDelete,
  onMove,
  onReorder,
}: SlideRailProps) => {
  const scale = THUMB_WIDTH / Math.max(slideWidthPx, 1)
  const thumbHeight = THUMB_WIDTH * (slideHeightPx / Math.max(slideWidthPx, 1))
  const canDelete = slides.length > 1
  const [overId, setOverId] = useState<string | null>(null)

  return (
    <aside className="slide-rail" aria-label="Slides">
      <div className="slide-rail__header">
        <h2>Slides</h2>
        <button type="button" className="slide-rail__add" onClick={onAdd} aria-label="Add slide">
          +
        </button>
      </div>
      <div className="slide-rail__list" role="list">
        {slides.map((slide, index) => {
          const isActive = slide.id === activeId
          return (
            <div
              key={slide.id}
              className={`slide-rail-item${isActive ? ' slide-rail-item--active' : ''}${overId === slide.id ? ' slide-rail-item--over' : ''}`}
              role="listitem"
              draggable
              onDragStart={(event) => {
                if ((event.target as HTMLElement).closest('button:not(.slide-thumb)')) {
                  event.preventDefault()
                  return
                }
                event.dataTransfer.setData('slideId', slide.id)
                event.dataTransfer.effectAllowed = 'move'
              }}
              onDragOver={(event) => {
                event.preventDefault()
                event.dataTransfer.dropEffect = 'move'
                setOverId(slide.id)
              }}
              onDragLeave={() => {
                setOverId((current) => (current === slide.id ? null : current))
              }}
              onDrop={(event) => {
                event.preventDefault()
                setOverId(null)
                const fromId = event.dataTransfer.getData('slideId')
                if (fromId) onReorder(fromId, slide.id)
              }}
              onDragEnd={() => setOverId(null)}
            >
              <div className="slide-rail-item__meta">
                <span className="slide-rail-item__index">{index + 1}</span>
                <div className="slide-rail-item__shuffle">
                  <button
                    type="button"
                    aria-label={`Move slide ${index + 1} up`}
                    disabled={index === 0}
                    onClick={() => onMove(slide.id, -1)}
                  >
                    ▲
                  </button>
                  <button
                    type="button"
                    aria-label={`Move slide ${index + 1} down`}
                    disabled={index === slides.length - 1}
                    onClick={() => onMove(slide.id, 1)}
                  >
                    ▼
                  </button>
                </div>
              </div>
              <button
                type="button"
                className={`slide-thumb${isActive ? ' slide-thumb--active' : ''}`}
                aria-current={isActive ? 'true' : undefined}
                aria-label={`Slide ${index + 1}`}
                onClick={() => onSelect(slide.id)}
              >
                <div
                  className="slide-thumb__frame"
                  style={{ width: THUMB_WIDTH, height: thumbHeight, background }}
                >
                  <div
                    className="slide-thumb__stage"
                    style={{
                      width: slideWidthPx,
                      height: slideHeightPx,
                      transform: `scale(${scale})`,
                      background,
                    }}
                    dangerouslySetInnerHTML={{ __html: slide.html }}
                  />
                </div>
              </button>
              <button
                type="button"
                className="slide-rail-item__delete"
                aria-label={`Delete slide ${index + 1}`}
                disabled={!canDelete}
                onClick={() => onDelete(slide.id)}
              >
                Delete
              </button>
            </div>
          )
        })}
      </div>
    </aside>
  )
}
