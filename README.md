# Office files — spreadsheets, Word documents, and presentations as COGNIX[WS] artifacts

A COGNIX[WS] extension that draws and edits Office Open XML artifact kinds:
spreadsheets (`.xlsx`), Word documents (`.docx`), and presentations (`.pptx`).
An agent publishes the file — or you import one — and this extension's pages
render and edit it in the artifact pane, on the Artifacts page, and on the phone.

## Install

1. In COGNIX[WS], open **Extensions → Install** and give this repository's
   address (or a local checkout path).
2. Turn it on in the workspace where the agent publishes.

## Publish or import Office artifacts

```sh
picode artifact publish q3.xlsx --title "Q3 plan" --kind ext:office/xlsx
picode artifact publish spec.docx --title "Product spec" --kind ext:office/docx
picode artifact publish pitch.pptx --title "Q4 review" --kind ext:office/pptx
```

Each version opens on its single file of the declared type under any name
(give `--entry` only when a version carries two files of that type). Importing
an `.xlsx`, `.docx`, or `.pptx` from **Artifacts → Import** creates version 1
with `publish` enabled so you can edit immediately.

## What the pages do

- **Offline sandboxed bundles:** all three viewers/editors run inside the
  strict artifact sandbox (`connect-src 'none'`, `script-src 'self'`, no CDN
  calls, no WebAssembly):
  - `ui/sheet.html` (`office.js` / `office.css`): [Univer](https://univer.ai)
    sheets engine + SheetJS OOXML reader/writer — formulas, sheet tabs, freeze
    panes, number formats, column widths, merges, and `Ctrl+F` find & replace.
  - `ui/doc.html` (`doc.js` / `doc.css`): Tiptap / ProseMirror rich document
    editor + lossless-preserving OOXML `.docx` reader/writer (`fflate` + XML
    DOM) — headings, character marks, colors, highlights, alignments, bullet
    and numbered lists, tables, inline images, links, word count, and `Ctrl+F`
    find & replace, while keeping untouched package parts (`header*`, `footer*`,
    `styles.xml`, `docProps`, `w:sectPr`, media) intact.
  - `ui/deck.html` (`deck.js` / `deck.css`): interactive slide stage + slide
    rail + speaker notes + lossless-preserving PresentationML `.pptx`
    reader/writer — master/layout/theme color & placeholder inheritance,
    shapes, text boxes, images, tables, connectors, groups, drag/resize,
    add/duplicate/reorder/delete slides, speaker notes, and `Ctrl+F` search
    across slides. Untouched slides keep their exact XML bytes.
- **Autosave & conflicts:** when a version declares `publish` (on desktop),
  edits save 1.5 s after the last change as a new immutable version authored
  by you (`Saved · vN`). If an agent publishes a newer version meanwhile, the
  conflict bar offers **Open the new version** or **Keep mine** rather than
  overwriting silently.
- **Store-time container rules (`entryContains`):**
  - `xlsx`: `["[Content_Types].xml"]`
  - `docx`: `["[Content_Types].xml", "word/document.xml"]`
  - `pptx`: `["[Content_Types].xml", "ppt/presentation.xml"]`
  Renamed or non-OOXML files are refused with HTTP 400 before a version is stored.
- **Languages & themes:** **Português** and **English** (defaulting to the
  browser's language); dark and light themes follow COGNIX[WS] live; narrow
  phone viewports (`≤ 560 px`) present a clean read-only layout.

## Development

```sh
cd web
npm install
npm test          # runs i18n, docx, and pptx round-trip unit tests
npm run build     # writes ui/office.{js,css}, ui/doc.{js,css}, ui/deck.{js,css}
```

The built files under `ui/` are committed so installing from a git tag works
without a build step.

## License

Apache 2.0. See [LICENSE](LICENSE).
