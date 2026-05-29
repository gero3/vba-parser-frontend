# VBA Parser Frontend

Frontend-only Office/VBA extraction tool. It runs locally in the browser and can inspect macro-enabled Office documents, VBA projects, forms, FRX resources, ActiveX package parts, and recovered media without uploading documents to a server.

## Features

- Extracts VBA modules, classes, document modules, and UserForms.
- Parses `PROJECT` and `dir` metadata, including VBA references.
- Analyzes `.frm` and `.frx` resources.
- Recovers common embedded media formats where possible.
- Parses MSForms / MS-OFORMS control metadata.
- Parses Office ActiveX XML and `.bin` persistence storages.
- Builds a migration ZIP with JSON, Markdown reports, procedure chunks, designer summaries, and visual previews.
- Supports direct `.bas`, `.cls`, `.frm`, and `.frx` files.

## Privacy

Document processing happens in the browser. There is no backend upload path in this app. A local static server may be used to serve the HTML/CSS/JS files, but parsing and export are client-side.

## Development

```bash
npm install
npm run dev
```

Vite serves the app locally and loads `src/main.ts` as the entrypoint. The source is split into focused ES modules under `src/`, and Vite bundles/minifies them for production.

No generated JavaScript is committed.

## Production Build

```bash
npm run build:prod
```

This creates a minified static site in `dist/` using Vite.

## Repository Scope

This repository intentionally excludes downloaded demo Office files, example workbooks/documents, temporary parser outputs, and generated `dist/` files.
