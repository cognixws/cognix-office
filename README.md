# Office files — spreadsheets, Word documents, and presentations in COGNIX[WS]

A COGNIX[WS] extension providing a single **Office files** app and three Office
Open XML artifact editors: spreadsheets (`.xlsx`), Word documents (`.docx`),
and presentations (`.pptx`).

## Install

1. In COGNIX[WS], open **Extensions → Install** and give this repository's
   address (or install **Office files** from **Extensions → Browse**).
2. Turn it on in the workspace where you and your agents publish artifacts.

## Two ways to use Office files

### 1. From **Apps → Office files** (standalone)
Open **Office files** in the **Apps** grid to create a blank spreadsheet, Word
document, or presentation, or open an `.xlsx`, `.docx`, or `.pptx` file from
your computer. Working files save automatically in the extension's state so you
can return to them from **Office files**.

### 2. From **Artifacts** (versioned collaboration with agents)

```sh
picode artifact publish q3.xlsx --title "Q3 plan" --kind ext:office/xlsx
picode artifact publish spec.docx --title "Product spec" --kind ext:office/docx
picode artifact publish pitch.pptx --title "Q4 review" --kind ext:office/pptx
```

Each version opens on its single file of the declared type under any name
(give `--entry` only when a version carries two files of that type). Importing
an `.xlsx`, `.docx`, or `.pptx` from **Artifacts → Import** creates version 1
with `publish` enabled so you can edit immediately.

## Architecture

- **Single unified entry (`ui/office.html` + `ui/icon.svg`):** `picode-extension.json`
  declares one `ui.pages` entry (`office`) so **Apps** shows a single **Office
  files** tile with its icon. When opened inside an artifact frame, `office.html`
  routes directly to `ui/sheet.html`, `ui/doc.html`, or `ui/deck.html` so each
  artifact downloads only its own engine.
- **Offline sandboxed bundles (`connect-src 'none'`, `script-src 'self'`, no WASM):**
  - `ui/sheet.html` (`office.js` / `office.css`): [Univer](https://univer.ai)
    sheets engine + SheetJS OOXML reader/writer — formulas, sheet tabs, freeze
    panes, number formats, column widths, merges, and `Ctrl+F` find & replace.
  - `ui/doc.html` (`doc.js` / `doc.css`): Tiptap / ProseMirror rich document
    editor + lossless-preserving OOXML `.docx` reader/writer (`fflate` + XML
    DOM) — headings, character marks, colors, highlights, alignments, bullet
    and numbered lists, tables, inline images, links, word count, and `Ctrl+F`
    find & replace, keeping untouched parts (`header*`, `footer*`, `styles.xml`,
    `docProps`, `w:sectPr`, media) intact.
  - `ui/deck.html` (`deck.js` / `deck.css`): interactive slide stage + slide
    rail + speaker notes + lossless-preserving PresentationML `.pptx`
    reader/writer — master/layout/theme color & placeholder inheritance,
    shapes, text boxes, images, tables, connectors, groups, drag/resize,
    add/duplicate/reorder/delete slides, speaker notes, and `Ctrl+F` search
    across slides. Untouched slides keep their exact XML bytes.
- **Autosave & conflicts:** artifact versions with `publish` save 1.5 s after
  the last edit (`Saved · vN`) and surface a conflict bar if an agent publishes
  concurrently. Standalone working files in **Office files** autosave to the
  extension's key-value state (`Saved`).
- **Store-time container rules (`entryContains`):**
  - `xlsx`: `["[Content_Types].xml"]`
  - `docx`: `["[Content_Types].xml", "word/document.xml"]`
  - `pptx`: `["[Content_Types].xml", "ppt/presentation.xml"]`

## Development

```sh
cd web
npm install
npm test          # runs i18n, hub, docx, and pptx unit tests
npm run build     # writes ui/hub.{js,css}, ui/office.{js,css}, ui/doc.{js,css}, ui/deck.{js,css}
```

The built files under `ui/` are committed so installing from a git tag works
without a build step.

## License

Apache 2.0. See [LICENSE](LICENSE).
