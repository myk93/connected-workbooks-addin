# Connected Workbooks — Excel Add-in

An Excel Task Pane add-in that lets you export any selection of cells as a standalone `.xlsx` workbook — or as a Power Query–connected workbook — directly from inside Excel, without leaving the app.

Built on top of the [`@microsoft/connected-workbooks`](https://github.com/microsoft/connected-workbooks) library.

---

## What it does

Select a range of cells in Excel, open the task pane, and choose how to export:

| Mode | What happens |
|------|-------------|
| **Download** | Generates a clean `.xlsx` file and downloads it to your computer |
| **Open in Web** | Uploads the workbook and opens it instantly in Excel for the Web |
| **Power Query** | Packages the data with an M query you write — downloads a connected `.xlsx` |

### Smart table detection

When your selection is inside an Excel table, the add-in offers extra options:

- **Selected cells only** — export just the rows/columns you highlighted
- **Entire table** — export the full table even if you only selected part of it
- **Re-export as Power Query** — if the table is backed by a Power Query, the add-in automatically extracts the M formula and pre-fills the editor for you

---

## Requirements

| Requirement | Version |
|-------------|---------|
| Node.js | 18 or later |
| Excel Desktop | Microsoft 365 (for sideloading) |
| Excel for the Web | Any modern browser |
| ExcelApi | 1.2 minimum · 1.9+ for table detection · 1.14+ for Power Query extraction |

---

## Getting started

### 1. Clone and install

```bash
git clone https://github.com/<your-org>/connected-workbooks-addin.git
cd connected-workbooks-addin
npm install
```

### 2. Install dev certificates (one-time)

The dev server must run over HTTPS for Excel to load it. Run this once:

```bash
npx office-addin-dev-certs install
```

You may be prompted for administrator permission — this is expected. It installs a trusted localhost certificate into your system's certificate store.

### 3. Start the dev server

```bash
npm start
```

The task pane is now being served at `https://localhost:3000`.

### 4. Sideload into Excel

**Option A — automatic (Excel Desktop):**
```bash
npm run sideload
```
This launches Excel and registers the add-in automatically.

**Option B — manual (Excel Desktop or Web):**
1. Open Excel
2. Go to **Insert → Office Add-ins → Upload My Add-in**
3. Browse to `manifest.xml` in this folder and click **Upload**
4. The **"Export"** button will appear on the **Home** tab under **Connected Workbooks**

---

## Using the add-in

### Basic export

1. Select any range of cells in Excel (e.g. `A1:D20`)
2. Click **Export** on the Home ribbon to open the task pane
3. The selection preview updates automatically as you change your selection
4. Toggle **"Treat first row as headers"** if your data has a header row
5. Choose a mode tab and click **Export**

### Power Query mode

1. Switch to the **Power Query** tab
2. Write your M expression in the text area — just the body, no `section` declaration needed:
   ```
   let
       Source = Csv.Document(Web.Contents("https://example.com/data.csv")),
       Headers = Table.PromoteHeaders(Source)
   in
       Headers
   ```
3. Set a **Query name** (default: `Query1`)
   - Max 80 characters
   - No quotes `"`, periods `.`, or control characters
4. Leave **Refresh on open** on so the query runs automatically when the workbook opens
5. Click **Export** — the workbook downloads with your query embedded

### Exporting from a Power Query table

If you click a cell inside a table that was loaded by Power Query, the add-in automatically:
- Detects the linked query
- Extracts the M formula
- Switches to Power Query mode and pre-fills the editor

You can edit the formula before exporting to create a modified version of the query.

---

## Project structure

```
connected-workbooks-addin/
├── manifest.xml                  # Office Add-in manifest
├── package.json
├── tsconfig.json                 # TypeScript config for src/
├── tsconfig.node.json            # TypeScript config for webpack.config.ts
├── webpack.config.ts
├── assets/                       # Add-in icons (16/32/64/80px)
└── src/
    ├── types/
    │   └── addin.ts              # Shared types: ExportMode, SelectionData, TableInfo …
    ├── services/
    │   ├── officeService.ts      # All Office.js API calls
    │   ├── workbookService.ts    # Wrapper around @microsoft/connected-workbooks
    │   └── mashupExtractor.ts    # Decodes Power Query M formula from DataMashup XML
    ├── hooks/
    │   ├── useSelection.ts       # Selection state + auto-refresh on cell change
    │   └── useExport.ts          # Export loading / success / error state
    └── taskpane/
        ├── index.html            # HTML shell (loads Office.js from CDN)
        ├── index.tsx             # Entry point: Office.onReady → React mount
        └── components/
            ├── App.tsx           # Root component — all state and export flow
            ├── ModeSelector.tsx
            ├── HeaderToggle.tsx
            ├── SelectionPreview.tsx
            ├── PowerQueryPanel.tsx
            ├── ExportButton.tsx
            ├── StatusBar.tsx
            └── TableScopeDialog.tsx
```

For a detailed diagram of how data flows through the app, see [ARCHITECTURE.md](./ARCHITECTURE.md).

---

## Available scripts

| Script | Description |
|--------|-------------|
| `npm start` | Start HTTPS dev server at `https://localhost:3000` |
| `npm run build` | Production bundle → `dist/` |
| `npm run build:dev` | Development bundle → `dist/` |
| `npm run sideload` | Launch Excel Desktop with the add-in sideloaded |
| `npm run validate` | Validate `manifest.xml` against the Office schema |

---

## Production deployment

1. Host the contents of `dist/` on any HTTPS web server or CDN
2. Update the `<SourceLocation>` and all `<bt:Url>` entries in `manifest.xml` to point to your hosted URL
3. Distribute `manifest.xml` to users via a SharePoint App Catalog or Microsoft 365 Admin Center

---

## Tech stack

- **React 18** + **TypeScript**
- **Fluent UI v9** (Microsoft's design system)
- **Webpack 5** with HTTPS dev server
- **Office.js** (loaded from CDN, never bundled)
- **[@microsoft/connected-workbooks](https://www.npmjs.com/package/@microsoft/connected-workbooks)** for workbook generation

---

## Contributing

1. Fork the repository and create a branch from `main`
2. Make your changes — run `npx tsc --noEmit` to type-check before committing
3. Open a pull request with a clear description of what changed and why

---

## License

MIT
