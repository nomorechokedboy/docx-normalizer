# Margin Lab — DOCX margin normalization prototype

A small browser-first proof of concept for normalizing page margins in `.docx` files while retaining the original OOXML structures and resources as much as possible. Designed so the UI can later be wrapped in Tauri.

## What it demonstrates

- Select a `.docx` file and inspect section count/current margin values.
- Choose a preset or enter custom top/right/bottom/left margins in millimeters.
- Rebuild the DOCX ZIP package, retaining all original entries and changing only `word/document.xml` margin attributes.
- Preview the original and candidate side-by-side using `docx-preview`.
- Download the candidate DOCX.
- Process files locally in the browser; no server upload is implemented.

## Run

Requires Node.js 20+ and pnpm/npm.

```bash
pnpm install
pnpm dev
```

Or use `npm install && npm run dev`.

## Current limitations

- This is a demo, not a production document conversion engine.
- The preset called `Administrative example` is illustrative; verify the applicable regulation/profile before using it for official documents.
- It writes four margin attributes on each `w:sectPr` in `word/document.xml`. It does not implement advanced WordprocessingML inheritance, Strict OOXML, or every edge case in malformed/atypical documents.
- The package is re-created from original ZIP entries, but only `word/document.xml` is edited. It is a preservation-oriented package reconstruction, not a semantic rebuild of every paragraph/table/image.
- The integrity check verifies that the resulting ZIP can be reopened and contains core parts. It does not detect visual clipping, broken table geometry, or cross-renderer differences automatically.
- `docx-preview` is a browser preview, not Microsoft Word. Final production validation should compare PDFs rendered with a controlled LibreOffice installation and include structural checks.
- The app uses system fonts and does not require external network access at runtime.

## Next milestone

Move the ZIP/XML manipulation into a separately tested engine; add OOXML regression fixtures, LibreOffice PDF rendering, page-image comparison, and conservative warnings for suspected clipping/table overflow before shipping output.
