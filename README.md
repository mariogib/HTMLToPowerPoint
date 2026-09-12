# HTML to PowerPoint

A browser-based slide editor that turns a visual canvas and HTML into a multi-slide PowerPoint file. Layout is measured from the DOM (CSS pixels at 96 DPI) so positions, line breaks, list bullets, and image aspect ratios match the exported PPTX. A second page converts PowerPoint files back to standalone HTML.

Live demo: [https://mariogib.github.io/HTMLToPowerPoint/](https://mariogib.github.io/HTMLToPowerPoint/)

## Features

### Editor

- Drag headings, text boxes, lists, and images onto a widescreen slide canvas.
- Move, resize, and delete controls (grip, handles, Delete/Backspace).
- Text boxes auto-size to their content; newlines in the editor appear in Preview, HTML, and PowerPoint.
- Right-click a control for properties: text, font, size, color, bold, italic, alignment, and position.
- List properties include bullet styles: filled circle, hollow circle, square, dash, arrow, check, numbers, or none.
- Images keep their natural aspect ratio in the editor and in PowerPoint (`object-fit: contain` / fitted export). Drop or pick a file to replace an image.
- Undo, cut, and paste (`Ctrl+Z`, `Ctrl+X`, `Ctrl+V`), plus a canvas context menu with the same commands (and Properties when a control is selected).
- **Load** an `.html` file into the current slide and place it on the canvas.
- **Save** writes the canvas to Preview, then opens a Save As dialog so you can choose folder and file name (falls back to a download if the File System Access API is unavailable).
- **Load from Preview** rebuilds the canvas from the current slide HTML.
- **Preview** (button or tab) saves the canvas to HTML and opens Preview. Toggling between Preview and HTML does not overwrite markup from the editor.

### Preview and HTML

- Tabs: **Editor** · **Preview** · **HTML** (HTML is available from Preview).
- Preview shows the slide HTML at the configured slide size.
- HTML is the raw markup for the current slide (editable).

### Slides

- Right-hand slide strip: add, delete, select, and reorder slides (▲ / ▼ or drag).
- At least one slide is always kept.
- **Export PPTX** writes the full deck in strip order.

### PowerPoint to HTML

- Open **PowerPoint to HTML** in the top nav (`#/pptx-to-html`).
- Drag and drop a `.pptx` or `.ppt` file, convert slides to HTML, preview them, and download a self-contained HTML file.
- Preserves slide size, backgrounds, text, images, SmartArt (where supported), and 3D models when possible.

### Export

- Client-side PPTX via PptxGenJS (no server).
- Settings: file name, slide width/height (inches), background color, base URL for relative images, embed images as data URLs.
- Headings, paragraphs, lists (with chosen bullets), and images are placed using measured coordinates.
- SVG images are rasterized to PNG for PowerPoint.

## Tech stack

| Layer | Choice |
| --- | --- |
| UI | React 19, TypeScript |
| Bundler | Vite 7 (Rolldown) |
| PPTX out | [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) 4 |
| PPTX in | JSZip, react-dropzone, file-saver |
| 3D / textures | buffer, dds-ktx-parser, upng-js |
| Layout | Off-screen DOM measurement, 96 CSS px = 1 inch |
| File save | File System Access API (`showSaveFilePicker`) |
| Hosting | GitHub Pages (`gh-pages` branch) |

**Dev tooling:** ESLint, typescript-eslint, React Compiler (Babel), `@vitejs/plugin-react`.

## Getting started

Install dependencies:

```bash
npm install
```

Start the dev server (use a free port if 5173 is blocked; this project often runs on 3000/3001 on Windows):

```bash
npm run dev
```

Production build:

```bash
npm run build
```

Preview the production build:

```bash
npm run preview
```

## Usage

1. Open **Editor** and drag controls onto the slide, or **Load** / **Load from Preview**.
2. Use the strip on the right to add and reorder slides.
3. Click **Preview** to review the slide; open **HTML** from Preview to inspect or edit markup.
4. Open Settings (gear) for slide size, background, file name, and image options.
5. Click **Export PPTX**.

## Notes

- Everything runs in the browser; HTML and images are not uploaded to a server.
- Remote images need CORS to embed. Relative URLs are resolved with **Base URL**.
- GitHub Pages is published from the `gh-pages` branch (`base: './'` in Vite). After changes on `main`, rebuild and publish `dist` to update the live site.
