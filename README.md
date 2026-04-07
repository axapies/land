# Land Layout Planner

A desktop-first SVG planning surface for arranging a land boundary, house, garage, garden, driveway, notes, and calibrated background imagery in metric world coordinates.

## Run

From this directory:

```bash
npm install
npm run build
open index.html
```

For iterative work:

```bash
npm run watch
```

## Deploy On Vercel

Publish only this `land-layout-planner/` folder to a dedicated GitHub repository, or set Vercel's project root directory to `land-layout-planner` if you keep it inside a larger repository.

Vercel settings:

```text
Framework Preset: Other
Build Command: npm run build
Output Directory: .
```

The app persists project history in the user's browser with `localStorage`. Use JSON export/import for moving projects between browsers or devices.

## Implemented First Slice

- Multiple local projects with tabs, history, duplicate, delete, import, export, and `localStorage` persistence.
- New project tabs start as blank canvases; duplicate keeps the current option for comparison.
- SVG grid-paper workspace with zoom, pan, scale label, north arrow, and snapping.
- Rectangular preset land plus click-to-draw freeform land boundary, side-length editing, removal, and a front-edge marker.
- Rectangular and freeform polygon site elements with category defaults.
- Selection, dragging, exact X/Y, width/depth, rotation, color, label, and legend controls.
- Freeform land and element vertex editing.
- Notes, generated legend, overlap/out-of-bound warnings, and distance overlays rounded to 0.1 m.
- Background image import with metric calibration controls, opacity, and locking.
- Browser print/PDF export path for A4/A3 with optional summary page.
- Undo/redo for recorded editing actions.

This is intentionally self-contained TypeScript/SVG so it does not disturb the existing GSD package in the repository root.
