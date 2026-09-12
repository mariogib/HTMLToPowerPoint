import PowerPointConverter from './components/PowerPointConverter'
import './PptxToHtmlPage.css'

function PptxToHtmlPage() {
  return (
    <div className="pptx-to-html">
      <header className="pptx-to-html__header">
        <h1>PowerPoint to HTML Converter</h1>
        <p>Convert your PowerPoint presentations to HTML with pixel-perfect accuracy</p>
      </header>
      <main className="pptx-to-html__main">
        <PowerPointConverter />
      </main>
    </div>
  )
}

export default PptxToHtmlPage
