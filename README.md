# HTML to PowerPoint

Convert HTML into a PPTX slide and embed HTML images directly into the exported file. The app renders HTML to an off-screen canvas, embeds images as data URLs, and exports a single-slide PowerPoint file.

## Features

- Paste HTML and export a PPTX in one click.
- Embed HTML images before rendering.
- Configure slide width, height, and background color.

## Getting started

Install dependencies:

```
npm install
```

Start the dev server:

```
npm run dev
```

Build for production:

```
npm run build
```

## Usage

1. Paste or edit HTML in the left panel.
2. Set slide size and background color.
3. Click “Export PPTX”.

## Notes

- Remote images require CORS access to embed successfully.
- Relative image URLs use the “Base URL” value for resolution.
