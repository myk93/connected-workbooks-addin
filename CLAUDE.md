# CLAUDE.md

This file provides guidance to Claude Code when working with the `connected-workbooks-addin` Office Task Pane add-in.

## Project Overview

A React 18 + TypeScript Office Add-in that lets users export Excel selections as `.xlsx` workbooks using the `@microsoft/connected-workbooks` library. Three export modes: **Download**, **Open in Web**, and **Power Query** (always downloads).

## Common Commands

```bash
# One-time cert setup (required for HTTPS dev server)
npx office-addin-dev-certs install

# Start dev server at https://localhost:3000
npm start

# Sideload into Excel Desktop and launch it
npm run sideload

# Type-check without emitting
npx tsc --noEmit

# Production bundle
npm run build
```

## Architecture

See `ARCHITECTURE.md` for full diagrams. Short summary:

```
src/
├── types/addin.ts          # All shared types: ExportMode, ExportScope, SelectionData, TableInfo, LinkedQuery
├── services/
│   ├── officeService.ts    # ALL Office.js calls (getSelectedRange + table/PQ detection)
│   ├── workbookService.ts  # Thin wrapper around @microsoft/connected-workbooks API
│   └── mashupExtractor.ts  # Decodes DataMashup binary → M formula text
├── hooks/
│   ├── useSelection.ts     # Selection state + DocumentSelectionChanged listener (debounced 350ms)
│   └── useExport.ts        # Export state machine (loading/success/error)
└── taskpane/
    ├── index.html          # Loads Office.js from CDN (never bundle it)
    ├── index.tsx           # Office.onReady → ReactDOM.createRoot
    └── components/
        ├── App.tsx             # All state; two-step export flow (read → optional dialog → export)
        ├── ModeSelector.tsx    # Tab strip: Download / Open in Web / Power Query
        ├── HeaderToggle.tsx    # "Treat first row as headers" checkbox
        ├── SelectionPreview.tsx# Address + dims + PQ badge; auto-updates on selection change
        ├── PowerQueryPanel.tsx # M textarea, query name (validated), refresh switch
        ├── ExportButton.tsx    # Primary CTA with spinner
        ├── StatusBar.tsx       # Success/error MessageBar
        └── TableScopeDialog.tsx# Dialog: Selected cells / Full table / Re-export as PQ
```

## Critical Webpack Notes

| Setting | Reason |
|---------|--------|
| `resolve.symlinks: false` | `@microsoft/connected-workbooks` is a `file:` symlink; without this webpack resolves transitive deps from the wrong directory |
| `resolve.fallback: { buffer: ... }` | Library uses `Buffer.from()` which doesn't exist in browsers |
| `externals: { "office-js": "Office" }` | Office.js must come from CDN, never bundled |
| `TS_NODE_PROJECT=tsconfig.node.json` | Main tsconfig uses `"module": "ESNext"` which breaks `__dirname` in webpack.config.ts; the node tsconfig overrides to CommonJS |

## Export Mode Routing

| Tab active | Dialog scope | Output |
|---|---|---|
| Download | selection / fullTable | Downloads `.xlsx` |
| Open in Web | selection / fullTable | Opens in Excel for the Web |
| Power Query | — (no dialog shown) | Downloads `.xlsx` |
| Any tab (dialog) | powerQuery | Downloads `.xlsx` |

The table-scope dialog is **skipped entirely** when the Power Query tab is active (`mode !== "powerQuery"` guard in `App.tsx`).

## Power Query M Extraction

Excel's JS API (`workbook.queries`) exposes query *metadata* only — not the M formula. To read the actual M expression the code:

1. `workbook.customXmlParts.getByNamespace(pqNs)` — finds the DataMashup XML part
2. `part.getXml()` — reads raw XML containing a base64 blob
3. Decode base64 → binary; parse layout: `[4B version][4B LE packageSize][packageSize B OPC zip]`
4. `JSZip.loadAsync(packageOPC)` → read `Formulas/Section1.m`
5. Regex to extract `shared #"QueryName" = <expr>;` → returns just `<expr>`

Requires ExcelApi 1.9 (getTables) + 1.14 (workbook.queries). Degrades silently on older versions.

## Key Constraints

- `@types/office-js` — `Excel.Query` has no `.formula` property; M lives in DataMashup custom XML only
- Query names: max 128 chars (library), max 80 chars (UI), no `"` `.` or control chars
- PQ export always downloads (never `openInExcelWeb`) by design
- `connected-workbooks` creates single-query workbooks only
