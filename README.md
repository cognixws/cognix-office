# Office files — spreadsheets as COGNIX[WS] artifacts

A COGNIX[WS] extension that draws an artifact kind COGNIX[WS] has no viewer
for: a spreadsheet. The agent publishes the workbook; this extension's own
page renders it in the artifact pane, on the Artifacts page and on the phone.

## Install

1. In COGNIX[WS], open **Extensions → Install** and give this repository's
   address (or a local checkout path).
2. Turn it on in the workspace where the agent publishes.

## Publish a workbook

```sh
picode artifact publish q3.xlsx --title "Q3 plan" --kind ext:office/xlsx
```

The version opens on its only `.xlsx` file, under any name. Two workbooks in
one version need `--entry` to say which one it opens on.

## What the page is

- **Sheet engine:** [Univer](https://univer.ai) (Apache-2.0 core), bundled in
  `ui/office.js` at build time — no CDN at runtime, so it works offline and
  under the artifact CSP (`script-src 'self'`), which a sandboxed page allows
  for its own files.
- **Reading the workbook:** SheetJS (Apache-2.0) parses the version's base64
  bytes and converts them into a Univer workbook — values, formulas, number
  formats, column widths and merges. The page reads that one version through
  the artifact door.
- **Editing (v2.1):** when the version declared `publish`, the grid is
  editable and saves itself 1.5 s after the last change as a new version,
  authored by the person, note "Saved by the page". The bar reads
  `Salvo · vN`; a version that lands meanwhile raises a conflict line
  ("Abra a vN" / "Manter a minha" / baixar) instead of overwriting it.
  Read-only versions keep `setEditable(false)`.
- **Find and replace:** `Ctrl+F` (or `Cmd+F`) inside the grid opens
  "Localizar" with the match count; **Pesquisa e Substituição Avançadas**
  brings the replace field and the search options. The dialog is capped to the
  pane and scrolls inside it, so its actions stay reachable in a small pane.
- **Language:** the bar's selector offers **Português** and **English**, and
  the browser's language is the default on every load (the artifact frame is
  sandboxed without storage, so a choice is not kept across loads). The choice
  moves Univer's own UI and the sheet's number formats: pt-BR reads
  `1.234,57`, en-US `1,234.57`. The patterns saved in the `.xlsx` are the
  workbook's own (`#,##0.00`) — the display never rewrites the file.
- **A file that is not a workbook** is refused before a version is stored:
  the manifest declares `entryContains: ["[Content_Types].xml"]`, so PiCode
  refuses a version whose entry is not a ZIP holding an OOXML workbook's
  content types. The page also refuses one with a sentence ("This workbook
  could not be read…") rather than drawing an empty grid.
- **Freeze panes** from the bar's **Freeze…** menu: freeze the rows above or
  the columns left of the active cell, or both; unfreeze. It is a view
  setting of the current pane — the `.xlsx` is not rewritten, so it is not
  kept across loads.

## Development

```sh
cd web
npm install
npm run build     # writes ui/office.js and ui/office.css
```

The built files are committed: an install clones the repository at a tag and
must work without a build step. Rebuild them when a dependency changes, and
bump the manifest's `version` with the release tag (`v<version>`); 0.3.6 is
freeze panes and the entry container rule (0.3.5 was the language selector).

## Limits

- One workbook per version, up to 16 MiB (the artifact cap); the page loads
  the whole file before drawing it.
- Number formats are the workbook's format codes; their separators follow the
  page's language (pt-BR `1.234,57`, en-US `1,234.57`), not a stored sheet
  locale — the artifact is the same file in either language.
- The phone shows the same page; the grid scrolls and pinches. Editing on a
  phone is untested — the pane is small and touch selection is the engine's.
- Find and replace is inside the sheet's own UI; there is no PiCode-side
  search over a workbook.

## License

Apache 2.0. See [LICENSE](LICENSE).
