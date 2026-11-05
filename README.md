# Office files — spreadsheets as COGNIX[WS] artifacts

A COGNIX[WS] extension that draws an artifact kind COGNIX[WS] has no viewer
for: a spreadsheet. An agent publishes the workbook, and this extension's own
page renders it in the artifact pane, beside the agent, on the Artifacts page
and on the phone. The file never enters COGNIX[WS]'s closed preview list.

## Install

1. In COGNIX[WS], open **Extensions → Install** and give this repository's
   address (or a local checkout path).
2. Turn it on in the workspace where the agent publishes.

## Publish a workbook

Ask the agent for a spreadsheet, or publish one yourself through the artifacts
tool with the extension's kind:

```sh
picode artifact publish q3.xlsx --title "Q3 plan" --kind ext:picode-office/xlsx
```

The version opens on its only `.xlsx` file, under any name. Two workbooks in
one version need `--entry` to say which one it opens on.

## How it works

- The manifest declares
  `contributes.artifactKinds: [{kind: "xlsx", title: "Spreadsheet", types: ["xlsx"], page: "sheet"}]`
  (ADR-0230, amendment of 2026-11-05). Declaring the type is what tells
  COGNIX[WS] to carry the file; versions, comments, Send, restore and backups
  stay COGNIX[WS]'s.
- The page reads that one version through the artifact frame door
  (`picode.use("artifact")`) — a workbook is a ZIP, so it arrives as base64 —
  and renders every sheet with SheetJS from the CDN the artifact policy
  already allows.
- It writes nothing and reaches nothing else.

## Development

No build step: `ui/sheet.html` is the page. To try a change, copy this folder,
open **Extensions → Install**, pick the folder, and publish a `.xlsx` to the
kind. The manifest's `version` is bumped per release, and the release tag is
`v<version>` (the catalog pins a tag).

## License

Apache 2.0. See [LICENSE](LICENSE).
