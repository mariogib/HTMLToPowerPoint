import App from './App'
import { SiteNav, useHashRoute } from './SiteNav'
import PptxToHtmlPage from './pptxToHtml/PptxToHtmlPage'

function Root() {
  const route = useHashRoute()

  return (
    <>
      <SiteNav />
      {route === '/pptx-to-html' ? <PptxToHtmlPage /> : <App />}
    </>
  )
}

export default Root
