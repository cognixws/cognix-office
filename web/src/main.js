import * as XLSX from "xlsx";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCorePtBR from "@univerjs/preset-sheets-core/locales/pt-BR";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import "@univerjs/preset-sheets-core/lib/index.css";

// The Office files page: one spreadsheet artifact, drawn by Univer from the
// workbook itself. The version arrives through the artifact door as base64
// (a workbook is a ZIP); this page reads nothing else and writes nothing.
// Read-only until the owner-edit doors land (ADR-0230 amendment).

const app = document.getElementById("app");
const state = document.getElementById("state");
const stateText = document.getElementById("state-text");

function fail(text) {
  if (state) { state.hidden = false; stateText.textContent = text; }
  if (app) app.replaceChildren();
}

document.getElementById("retry")?.addEventListener("click", () => location.reload());

function columnWidth(col) {
  if (!col) return 0;
  const w = col.wpx || (col.wch ? Math.round(col.wch * 7 + 5) : 0);
  return w ? Math.round(Math.max(64, w * 1.15)) : 0;
}

// toSnapshot turns a SheetJS workbook into a Univer workbook: cells (value,
// formula, number format), column widths, row heights and merges. Styles are
// workbook-level; each cell points at one by id.
function toSnapshot(book, title) {
  const sheets = {};
  const order = [];
  const styles = {};
  const styleIds = new Map();
  const styleOf = (format, bold) => {
    if (!format && !bold) return undefined;
    const key = (format || "") + "|" + (bold ? "b" : "");
    if (styleIds.has(key)) return styleIds.get(key);
    const id = "st" + styleIds.size;
    const style = {};
    if (format) style.n = { pattern: format };
    if (bold) style.bl = 1;
    styles[id] = style;
    styleIds.set(key, id);
    return id;
  };

  for (const name of book.SheetNames) {
    const ws = book.Sheets[name];
    const sheet = { id: name, name, cellData: {}, rowCount: 100, columnCount: 26 };
    if (ws && ws["!ref"]) {
      const range = XLSX.utils.decode_range(ws["!ref"]);
      for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (!cell) continue;
          const out = {};
          if (typeof cell.f === "string" && cell.f) out.f = "=" + cell.f;
          if (cell.v !== undefined && cell.v !== null) out.v = cell.v;
          const bold = cell.s && cell.s.font && cell.s.font.bold;
          const style = styleOf(cell.z, bold);
          if (style) out.s = style;
          (sheet.cellData[r] ||= {})[c] = out;
        }
      }
      sheet.rowCount = Math.max(range.e.r + 1, 24);
      sheet.columnCount = Math.max(range.e.c + 1, 12);
      if (ws["!cols"]) {
        const columnData = {};
        ws["!cols"].forEach((col, i) => {
          const w = columnWidth(col);
          if (w) columnData[i] = { w };
        });
        if (Object.keys(columnData).length) sheet.columnData = columnData;
      }
      if (ws["!rows"]) {
        const rowData = {};
        ws["!rows"].forEach((row, i) => {
          if (row && row.hpx) rowData[i] = { h: row.hpx };
        });
        if (Object.keys(rowData).length) sheet.rowData = rowData;
      }
      if (ws["!merges"] && ws["!merges"].length) {
        sheet.mergeData = ws["!merges"].map((m) => ({
          startRow: m.s.r,
          endRow: m.e.r,
          startColumn: m.s.c,
          endColumn: m.e.c,
        }));
      }
    }
    sheets[name] = sheet;
    order.push(name);
  }

  return {
    id: "workbook-" + Date.now(),
    name: title || "Spreadsheet",
    sheetOrder: order,
    styles,
    sheets,
  };
}

async function main() {
  let door;
  try {
    door = await picode.use("artifact");
  } catch {
    fail("This page draws an artifact version. Open it from the artifact, not on its own.");
    return;
  }
  let info, file;
  try {
    info = await door.get();
    file = await door.read(info.entry);
  } catch (e) {
    fail("This workbook could not be read. Its file is still in the artifact's versions.");
    return;
  }
  const base64 = file && file.base64;
  if (!base64) {
    fail("This file is not a workbook COGNIX[WS] can draw.");
    return;
  }

  let book;
  try {
    book = XLSX.read(base64, { type: "base64", cellNF: true, cellStyles: true });
  } catch {
    fail("This workbook could not be read. Its file is still in the artifact's versions.");
    return;
  }
  if (!book.SheetNames.length) {
    fail("This workbook holds no sheet.");
    return;
  }

  try {
    const { univerAPI } = createUniver({
      locale: LocaleType.PT_BR,
      locales: { [LocaleType.PT_BR]: mergeLocales(UniverPresetSheetsCorePtBR) },
      presets: [UniverSheetsCorePreset({
        container: "app",
        header: false,
        toolbar: false,
        formulaBar: false,
        menu: false,
        contextMenu: false,
        footer: true,
        statusBarStatistic: true,
        disableAutoFocus: true,
      })],
    });
    univerAPI.createWorkbook(toSnapshot(book, info.title));
    const workbook = univerAPI.getActiveWorkbook();
    if (workbook && typeof workbook.setEditable === "function") workbook.setEditable(false);
    document.body.dataset.ready = "1";
  } catch (e) {
    fail("This workbook could not be drawn. Its file is still in the artifact's versions.");
  }
}

main();
