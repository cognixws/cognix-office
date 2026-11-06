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
picode artifact publish q3.xlsx --title "Q3 plan" --kind ext:picode-office/xlsx
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
- **A file that is not a workbook** is refused on the page with one sentence
  ("This workbook could not be read…"), not drawn as an empty grid.

## Development

```sh
cd web
npm install
npm run build     # writes ui/office.js and ui/office.css
```

The built files are committed: an install clones the repository at a tag and
must work without a build step. Rebuild them when a dependency changes, and
bump the manifest's `version` with the release tag (`v<version>`); 0.3.4 is
find and replace plus the dialog cap.

## Limits

- One workbook per version, up to 16 MiB (the artifact cap); the page loads
  the whole file before drawing it.
- Number formats follow the workbook's format codes; separators are the
  sheet engine's (a pt-BR formatter is polish, not a promise).
- The phone shows the same page; the grid scrolls and pinches. Editing on a
  phone is untested — the pane is small and touch selection is the engine's.
- Find and replace is inside the sheet's own UI; there is no PiCode-side
  search over a workbook.

## License

Apache 2.0. See [LICENSE](LICENSE).
