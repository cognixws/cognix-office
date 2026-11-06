import * as XLSX from "xlsx";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCorePtBR from "@univerjs/preset-sheets-core/locales/pt-BR";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import "@univerjs/preset-sheets-core/lib/index.css";

// The Office files page: one spreadsheet artifact, drawn by Univer from the
// workbook itself. The version arrives through the artifact door as base64 (a
// workbook is a ZIP); when that version declared `publish`, the page saves
// what the owner edits as a new version of the same artifact (ADR-0230
// amendment 2026-11-06). It writes nothing else and names no other artifact.

const app = document.getElementById("app");
const state = document.getElementById("state");
const stateText = document.getElementById("state-text");
const saveEl = document.getElementById("save");
const saveText = document.getElementById("save-text");
const conflict = document.getElementById("conflict");
const conflictText = document.getElementById("conflict-text");
const bar = document.getElementById("bar");

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

  return { id: "workbook-" + Date.now(), name: title || "Spreadsheet", sheetOrder: order, styles, sheets };
}

// fromSnapshot writes a Univer workbook back as an .xlsx: values, formulas,
// number formats, column widths and merges — the same set the page reads.
function fromSnapshot(snap) {
  const book = XLSX.utils.book_new();
  for (const name of snap.sheetOrder || []) {
    const sheet = snap.sheets[name];
    if (!sheet) continue;
    const ws = {};
    let maxR = 0;
    let maxC = 0;
    const cellData = sheet.cellData || {};
    for (const r of Object.keys(cellData)) {
      for (const c of Object.keys(cellData[r])) {
        const cell = cellData[r][c] || {};
        const addr = XLSX.utils.encode_cell({ r: Number(r), c: Number(c) });
        const out = {};
        const style = cell.s && snap.styles ? snap.styles[cell.s] : null;
        if (style && style.n && style.n.pattern) out.z = style.n.pattern;
        if (cell.f) {
          out.f = String(cell.f).replace(/^=/, "");
          out.t = "n";
          out.v = typeof cell.v === "number" ? cell.v : 0;
        } else if (typeof cell.v === "number") {
          out.v = cell.v;
          out.t = "n";
        } else if (cell.v !== undefined && cell.v !== null) {
          out.v = String(cell.v);
          out.t = "s";
        } else {
          continue;
        }
        ws[addr] = out;
        maxR = Math.max(maxR, Number(r));
        maxC = Math.max(maxC, Number(c));
      }
    }
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(maxR, 0), c: Math.max(maxC, 0) } });
    if (sheet.columnData) {
      const cols = [];
      for (const c of Object.keys(sheet.columnData)) {
        const w = sheet.columnData[c] && sheet.columnData[c].w;
        if (w) cols[Number(c)] = { wpx: Math.round(w) };
      }
      if (cols.length) ws["!cols"] = cols;
    }
    if (sheet.mergeData && sheet.mergeData.length) {
      ws["!merges"] = sheet.mergeData.map((m) => ({ s: { r: m.startRow, c: m.startColumn }, e: { r: m.endRow, c: m.endColumn } }));
    }
    XLSX.utils.book_append_sheet(book, ws, String(name).slice(0, 31));
  }
  return XLSX.write(book, { bookType: "xlsx", type: "base64" });
}

function setSave(text, tone) {
  if (!saveEl) return;
  saveEl.dataset.tone = tone || "";
  saveText.textContent = text;
}

function showConflict(liveVersion) {
  if (!conflict) return;
  conflict.hidden = false;
  conflictText.textContent = "Um agente publicou a versão " + liveVersion + " enquanto você editava. Sua edição está aqui.";
  setSave("Não salvo · v" + liveVersion + " chegou", "error");
  conflict.dataset.live = String(liveVersion);
}

function hideConflict() {
  if (conflict) conflict.hidden = true;
}

async function main() {
  let door;
  try {
    door = await picode.use("artifact");
  } catch {
    fail("This page draws an artifact version. Open it from the artifact, not on its own.");
    return;
  }
  let info;
  try {
    info = await door.get();
  } catch {
    fail("This workbook could not be read. Its file is still in the artifact's versions.");
    return;
  }
  if (!info || !info.entry) {
    fail("This file is not a workbook COGNIX[WS] can draw.");
    return;
  }
  const writable = Array.isArray(info.capabilities) && info.capabilities.indexOf("publish") >= 0;

  let file;
  try {
    file = await door.read(info.entry);
  } catch {
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

  let univerAPI;
  try {
    const made = createUniver({
      locale: LocaleType.PT_BR,
      locales: { [LocaleType.PT_BR]: mergeLocales(UniverPresetSheetsCorePtBR) },
      presets: [
        UniverSheetsCorePreset({
          container: "app",
          header: true,
          toolbar: true,
          formulaBar: true,
          menu: false,
          contextMenu: true,
          footer: true,
          statusBarStatistic: true,
          disableAutoFocus: !writable,
        }),
      ],
    });
    univerAPI = made.univerAPI;
    univerAPI.createWorkbook(toSnapshot(book, info.title));

    const workbook = univerAPI.getActiveWorkbook();
    if (!writable && workbook && typeof workbook.setEditable === "function") workbook.setEditable(false);
  } catch {
    fail("This workbook could not be drawn. Its file is still in the artifact's versions.");
    return;
  }

  if (bar) bar.hidden = false;
  if (bar) {
    const title = bar.querySelector(".name");
    if (title) title.textContent = info.title || "Planilha";
  }
  if (!writable) {
    setSave("Somente leitura", "");
    document.body.dataset.ready = "1";
    return;
  }

  let baseVersion = info.version || 1;
  // The canonical form of what was drawn: opening a workbook must not make a
  // version; only a real change differs from it.
  let lastSaved = "";
  try {
    lastSaved = fromSnapshot(univerAPI.getActiveWorkbook().save());
  } catch { lastSaved = base64; }
  let timer = 0;
  let saving = false;
  let pending = false;
  setSave("Salvo · v" + baseVersion, "ok");

  async function save() {
    clearTimeout(timer);
    if (saving) { pending = true; return; }
    let snapshot;
    try {
      snapshot = univerAPI.getActiveWorkbook().save();
    } catch {
      return;
    }
    const next = fromSnapshot(snapshot);
    if (next === lastSaved) { setSave("Salvo · v" + baseVersion, "ok"); return; }
    saving = true;
    setSave("Salvando…", "busy");
    try {
      const out = await door.publishFiles([{ path: info.entry, data: next }], baseVersion);
      baseVersion = (out && out.version && out.version.n) || baseVersion + 1;
      lastSaved = next;
      hideConflict();
      setSave("Salvo · v" + baseVersion, "ok");
    } catch (e) {
      if (e && e.code === -32009) {
        const live = Number(e.data && e.data.liveVersion) || 0;
        showConflict(live || baseVersion + 1);
      } else if (e && e.code === -32001) {
        setSave("Somente leitura", "");
      } else {
        setSave("Não salvo · tente de novo", "error");
      }
    } finally {
      saving = false;
      if (pending) { pending = false; timer = setTimeout(save, 200); }
    }
  }

  function schedule() {
    clearTimeout(timer);
    timer = setTimeout(save, 1500);
  }

  // Any command may change a cell or a format; selection alone is cheap to
  // skip because an unchanged workbook saves nothing (the digest check above).
  univerAPI.onCommandExecuted(() => schedule());

  conflict?.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-act]");
    if (!b) return;
    if (b.dataset.act === "open") { location.reload(); return; }
    if (b.dataset.act === "keep") {
      baseVersion = Number(conflict.dataset.live) || baseVersion;
      lastSaved = ""; // force the write on top of the live version
      hideConflict();
      await save();
    }
  });

  document.body.dataset.ready = "1";
}

main();
