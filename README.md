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
  the artifact door and writes nothing.
- **Read-only (v2.0):** `setEditable(false)`. Editing, autosave and conflict
  handling are v2.1, which needs the write doors in the core (ADR-0230
  amendment); see `docs/plans/artifact-xlsx-v2.md` in the COGNIX[WS]
  repository.

## Development

```sh
cd web
npm install
npm run build     # writes ui/office.js and ui/office.css
```

The built files are committed: an install clones the repository at a tag and
must work without a build step. Rebuild them when a dependency changes, and
bump the manifest's `version` with the release tag (`v<version>`).

## Limits

- One workbook per version, up to 16 MiB (the artifact cap); the page loads
  the whole file before drawing it.
- Number formats follow the workbook's format codes; separators are the
  sheet engine's (a pt-BR formatter is polish, not a promise).
- The phone shows the same read-only sheet; editing there is out of scope.

## License

Apache 2.0. See [LICENSE](LICENSE).
