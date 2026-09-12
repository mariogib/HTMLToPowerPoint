import React, { useState } from 'react'
import './SlideRenderer.css'

interface Slide {
  id: string
  content: string
  styles: string
  notes?: string
}

interface SlideRendererProps {
  slides: Slide[]
}

const SlideRenderer: React.FC<SlideRendererProps> = ({ slides }) => {
  const [currentSlide, setCurrentSlide] = useState(0)

  const goToNextSlide = () => {
    if (currentSlide < slides.length - 1) {
      setCurrentSlide(currentSlide + 1)
    }
  }

  const goToPreviousSlide = () => {
    if (currentSlide > 0) {
      setCurrentSlide(currentSlide - 1)
    }
  }

  const goToSlide = (index: number) => {
    if (index >= 0 && index < slides.length) {
      setCurrentSlide(index)
    }
  }

  if (slides.length === 0) {
    return (
      <div className="slide-renderer">
        <p>No slides to display</p>
      </div>
    )
  }

  return (
    <div className="slide-renderer">
      <div className="slide-navigation">
        <button 
          onClick={goToPreviousSlide} 
          disabled={currentSlide === 0}
          className="nav-button"
        >
          ← Previous
        </button>
        
        <span className="slide-counter">
          Slide {currentSlide + 1} of {slides.length}
        </span>
        
        <button 
          onClick={goToNextSlide} 
          disabled={currentSlide === slides.length - 1}
          className="nav-button"
        >
          Next →
        </button>
      </div>

      <div className="slide-container">
        <style dangerouslySetInnerHTML={{ __html: slides[currentSlide].styles }} />
        <div 
          className="slide-content"
          dangerouslySetInnerHTML={{ __html: slides[currentSlide].content }}
        />
      </div>

      <div className="slide-thumbnails">
        {slides.map((_, index) => (
          <button
            key={index}
            onClick={() => goToSlide(index)}
            className={`thumbnail ${index === currentSlide ? 'active' : ''}`}
          >
            {index + 1}
          </button>
        ))}
      </div>
    </div>
  )
}

export default SlideRenderer
