# Cardmaker

Cardmaker is a single-page tool for designing print-ready business cards directly in the browser. It provides canvas-based editing, inline typography controls powered by TinyMCE, and utilities for exporting designs to print layouts.

## Features

- Dual-sided canvas with zoomable preview, bleed/trim overlays, and per-layer stacking controls.
- Built-in text and image layers, plus per-card back-image overrides.
- Inline TinyMCE editor for rich typography (fonts, colors, lists, alignment).
- 3D card preview with configurable paper textures and thickness.
- Export helpers for PDF/imposition and project-saving to `.card` bundles.

## Getting started

### Prerequisites

- [Node.js](https://nodejs.org/) 18+ (for installing TinyMCE locally)

### Setup

```bash
npm install
```

### Run

Use any static HTTP server. Two easy options:

```bash
# Option 1: via npm script/lite-server if you add one
npx serve .

# Option 2: VS Code "Live Server" extension
```

Open `http://localhost:<port>/` in your browser.

## Project structure

```
static/
  fonts/        -> web fonts used by the UI
  img/          -> icons, textures, and brand marks
  demo/         -> sample `.card` project files (ACME.card)
index.html      -> UI layout
styles.css      -> theme + layout styling
app.js          -> editor logic (canvas, TinyMCE, export)
```

Keep any new fonts or images inside `static/` and reference them with `static/...` paths so the app stays self-contained.

## Development notes

- TinyMCE is loaded from `node_modules`. Ensure `npm install` has completed before launching the page.
- Inline editing requires a secure context for some browsers; run via `localhost` or HTTPS to avoid mixed-content/font warnings.
- Sample data lives in `static/demo/ACME.card`. Replace or expand with your own `.card` bundles for demos.

## License

Cardmaker is released under the [MIT License](https://opensource.org/licenses/MIT).  
made © [acme-prototypes.com](https://acme-prototypes.com/)
