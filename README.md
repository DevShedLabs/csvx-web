# CSVX Web

CSVX Web is the React/Vite demo workbench for opening, inspecting, editing, and exporting CSVX
files. It is a consumer of the CSVX specification and engine model, not a second definition of the
format.

## Current scope

The first slice is an accessible, developer-oriented read-only workbench with:

- Workbook and sheet navigation
- CSV-backed spreadsheet data display
- Cell selection with A1 coordinates
- Package and data-layer context
- Responsive layout and keyboard-visible focus states

The current demo data is represented in the UI while the browser-side CSVX package reader is being
added. `example.csvx` is the real CSVX package fixture for that integration.

## Commands

```bash
npm install
npm run dev
npm run build
npm run preview
```

## Direction

The UI will remain separate from CSVX semantics. The browser application should eventually load the
ZIP package client-side, preserve CSV and `.meta.json` resources, expose formulas and diagnostics,
and export a valid `.csvx` package without requiring an upload to a server.