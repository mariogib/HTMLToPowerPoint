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
    <nav className="site-nav" aria-label="App pages">
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
  )
}
