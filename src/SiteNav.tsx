import { useEffect, useState } from 'react'

export const useHashRoute = () => {
  const [hash, setHash] = useState(() => window.location.hash)

  useEffect(() => {
    const onChange = () => setHash(window.location.hash)
    window.addEventListener('hashchange', onChange)
    return () => window.removeEventListener('hashchange', onChange)
  }, [])

  return hash.replace(/^#/, '') || '/'
}

export const SiteNav = () => {
  const route = useHashRoute()
  const isPptxToHtml = route === '/pptx-to-html'

  return (
    <header className="site-chrome">
      <div className="site-nav">
        <a href="#/" className="site-nav__brand" aria-label="SlideBridge home">
          <span className="site-nav__title">SlideBridge</span>
          <span className="site-nav__badge">Prototype</span>
        </a>
        <nav className="site-nav__links" aria-label="App pages">
          <a
            href="#/"
            className={`site-nav__link${!isPptxToHtml ? ' site-nav__link--active' : ''}`}
          >
            HTML to PowerPoint
          </a>
          <a
            href="#/pptx-to-html"
            className={`site-nav__link${isPptxToHtml ? ' site-nav__link--active' : ''}`}
          >
            PowerPoint to HTML
          </a>
        </nav>
      </div>
      <p className="site-banner" role="note">
        <strong>Prototype limitations.</strong> SlideBridge converts HTML and PowerPoint in
        the browser. It is not Microsoft PowerPoint. Charts, tables, animations, transitions,
        videos, audio, comments, speaker notes, slide masters, and many effects are missing or
        flattened. SmartArt and 3D models are approximations and may not match Office. Fonts
        and text wrapping depend on the browser. Large decks can be slow or incomplete. Keep
        your original files.
      </p>
    </header>
  )
}
