# Aurelia Slides

A dependency-free presentation editor written in plain HTML, CSS, and JavaScript, with a PowerPoint-style ribbon, slide thumbnails, property inspector, and speaker notes. The rendering engine uses WebGPU when available and falls back to Canvas 2D.

**Application:** https://wieslawsoltes.github.io/AureliaSlides/

## Run

Open `index.html` directly, or serve the repository locally:

```sh
python3 -m http.server 8000
```

Open http://localhost:8000/ for the standalone application or http://localhost:8000/dev.html for the native ES-module development entry point. There are no runtime packages, remote scripts, accounts, or external assets. Localhost or HTTPS is required for WebGPU. The status bar reports the active rendering backend.

## Develop and test

Node.js 22 or newer is required for the build and kernel tests; no `npm install` is necessary.

```sh
npm test
npm run check
npm run build
npm run build:site
```

`npm run build` deterministically regenerates the single-file `index.html` from `dev.html` and the six source modules. `npm run build:site` also stages only the deployable application in `_site/`.

## Features

- Canvas editing: selection and marquee, Shift-select, moving, resizing, rotation, nudge, alignment guides, snapping, zoom, and pan.
- Text boxes: inline editing, Unicode, wrapping, font selection, size, bold/italic/underline, paragraph alignment, line spacing, and bullets. Formatting applies to the entire text box.
- Objects: common vector shapes, lines, raster pictures, process diagrams, logical grouping, lock/hide, layering, duplicate, copy/paste, and transactional undo/redo.
- Slides: creation, duplication, deletion, renaming, drag ordering, slide sorter, layouts, themes, backgrounds, and dimensions.
- Data: editable single-series bar, line, and donut charts; simple tables with tab-separated data editing.
- Presentation: slideshow controller, notes panel, transitions, click-triggered entrances, laser pointer, blackout, and timer.
- Files: native `.aurelia` project format, browser-local autosave, PPTX export and best-effort import, PNG/SVG export, slide-image archives, and browser printing.
- Example: a seven-slide presentation built from editable objects.

## Architecture

| File | Responsibility |
| --- | --- |
| `src/core.js` | Document model, validation, geometry, selection, transactions, undo/redo. |
| `src/renderer.js` | WebGPU pipelines and WGSL, instanced analytic shapes, cached text/image tiles, Canvas 2D fallback. |
| `src/demo.js` | Editable example presentation. |
| `src/icons.js` | Inline SVG UI icons. |
| `src/pptx.js` | ZIP reader/writer and PresentationML interchange. |
| `src/app.js` | Ribbon, interactions, inspector, persistence, file commands, slideshow controller. |
| `dev.html` | Editable HTML/CSS shell and module entry point. |
| `scripts/build.mjs` | Zero-dependency standalone bundler and Pages site staging. |
| `tests/core.test.mjs` | Kernel, geometry, history, chart/table expansion, validation, ZIP tests. |

The renderer preserves painter ordering, batches consecutive compatible shapes, grows instance buffers geometrically, and renders on demand. Browser text shaping/rasterization feeds cached GPU textures. Texture caching has a 96 MB soft budget. Moving a text box does not require rerasterizing its contents. Pointer gestures are single undoable transactions.

## GitHub Pages

The deployment workflow tests the source, rebuilds the standalone HTML, and publishes `_site/` using the official GitHub Pages actions. Configure **Settings → Pages → Build and deployment → Source → GitHub Actions** if the repository does not yet have Pages enabled.

Pages provides the secure origin needed by WebGPU, but hardware/browser availability still determines whether the GPU renderer can initialize.

## Validation and limitations

All **34 kernel tests pass** locally. A previous development-session browser run completed 14 editor integration checks using Canvas 2D, but stopped at slideshow-test setup. That browser run is not an automated test suite included here.

WebGPU hardware execution, secure-origin autosave, slideshow playback, and full PPTX round-trips have not been verified end-to-end. No frame-rate or production-readiness claim is made.

This is a working first implementation, not complete Microsoft PowerPoint parity. Missing areas include mixed-style text runs, comprehensive slide-master inheritance, SmartArt, embedded media, collaboration, and native Office chart objects. PPTX charts and tables export as drawing objects; donut charts export as raster images. Generic PPTX import is best-effort. Save a native `.aurelia` project to preserve the editor's own document model.

The project is independently developed and is not affiliated with Microsoft.
