import * as XLSX from "xlsx";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import UniverPresetSheetsCorePtBR from "@univerjs/preset-sheets-core/locales/pt-BR";
import UniverPresetSheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import { UniverSheetsFindReplacePreset } from "@univerjs/preset-sheets-find-replace";
import FindReplacePtBR from "@univerjs/preset-sheets-find-replace/locales/pt-BR";
import FindReplaceEnUS from "@univerjs/preset-sheets-find-replace/locales/en-US";
import { localeOf, numfmtOf, resolveLang, stringsFor } from "./i18n.js";
import { createUniver, mergeLocales } from "@univerjs/presets";
import "@univerjs/preset-sheets-core/lib/index.css";
import "./sheet.css";

// The Office files page: one spreadsheet artifact, drawn by Univer from the
// workbook itself. The version arrives through the artifact door as base64 (a
// workbook is a ZIP); when that version declared `publish`, the page saves
// what the owner edits as a new version of the same artifact (ADR-0230
// amendment 2026-11-06). It writes nothing else and names no other artifact.

const app = document.getElementById("app");
const state = document.getElementById("state");
const stateText = document.getElementById("state-text");
const stateTitle = document.getElementById("state-title");
const saveEl = document.getElementById("save");
const saveText = document.getElementById("save-text");
const conflict = document.getElementById("conflict");
const conflictText = document.getElementById("conflict-text");
const conflictOpen = conflict ? conflict.querySelector('[data-act="open"]') : null;
const conflictKeep = conflict ? conflict.querySelector('[data-act="keep"]') : null;
const bar = document.getElementById("bar");
const langEl = document.getElementById("lang");
const retryEl = document.getElementById("retry");

// The page's language: the browser's, changed here and held for this view.
// The artifact frame is sandboxed without storage, so a choice is not kept
// across loads; the browser's language is the default every time.
let lang = resolveLang("", typeof navigator === "undefined" ? "" : navigator.language);
let pageTitle = "";
const S = () => stringsFor(lang);

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

// The save state is kept as a key so a language change repaints it.
let saveState = { kind: "saved", n: 1 };
function paintSave() {
  const s = S();
  if (saveState.kind === "saved") setSave(s.saved(saveState.n), "ok");
  else if (saveState.kind === "saving") setSave(s.saving, "busy");
  else if (saveState.kind === "unsaved") setSave(s.unsaved, "error");
  else if (saveState.kind === "readonly") setSave(s.readOnly, "");
  else setSave(s.conflictSave(saveState.n), "error");
}
function setSaveState(kind, n) {
  saveState = { kind, n: n || saveState.n || 1 };
  paintSave();
}

function showConflict(liveVersion) {
  if (!conflict) return;
  conflict.hidden = false;
  conflictText.textContent = S().conflict(liveVersion);
  setSaveState("conflict", liveVersion);
  conflict.dataset.live = String(liveVersion);
}

function hideConflict() {
  if (conflict) conflict.hidden = true;
  delete conflict?.dataset.live;
}

// applyLanguage paints every string this page owns; Univer's own UI follows
// setLocale, and its number formats follow setNumfmtLocal on the workbook.
function applyLanguage() {
  if (langEl) { langEl.value = lang; langEl.setAttribute("aria-label", S().language); langEl.title = S().language; }
  if (typeof document !== "undefined") document.documentElement.lang = lang;
  if (stateTitle) stateTitle.textContent = S().undrawable;
  if (retryEl) retryEl.textContent = S().retry;
  const nameEl = bar && bar.querySelector(".name");
  if (nameEl) nameEl.textContent = pageTitle || S().title;
  if (conflictOpen) conflictOpen.textContent = S().open;
  if (conflictKeep) conflictKeep.textContent = S().keep;
  const live = Number(conflict?.dataset.live) || 0;
  if (conflict && !conflict.hidden && live) conflictText.textContent = S().conflict(live);
  paintSave();
}

async function main() {
  let door;
  try {
    door = await picode.use("artifact");
  } catch {
    fail(S().notStandalone);
    return;
  }
  let info;
  try {
    info = await door.get();
  } catch {
    fail(S().unreadable);
    return;
  }
  if (!info || !info.entry) {
    fail(S().notWorkbook);
    return;
  }
  pageTitle = info.title || "";
  const writable = Array.isArray(info.capabilities) && info.capabilities.indexOf("publish") >= 0;

  let file;
  try {
    file = await door.read(info.entry);
  } catch {
    fail(S().unreadable);
    return;
  }
  const base64 = file && file.base64;
  if (!base64) {
    fail(S().notWorkbook);
    return;
  }

  let book;
  try {
    book = XLSX.read(base64, { type: "base64", cellNF: true, cellStyles: true });
  } catch {
    fail(S().unreadable);
    return;
  }
  if (!book.SheetNames.length) {
    fail(S().noSheet);
    return;
  }

  let univerAPI;
  try {
    const made = createUniver({
      locale: localeOf(lang),
      locales: {
        [localeOf("pt-BR")]: mergeLocales(UniverPresetSheetsCorePtBR, FindReplacePtBR),
        [localeOf("en-US")]: mergeLocales(UniverPresetSheetsCoreEnUS, FindReplaceEnUS),
      },
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
        // Ctrl+F inside the grid (find and replace), the daily Excel/Sheets
        // habit; the preset brings the dialog and the sheet commands.
        UniverSheetsFindReplacePreset(),
      ],
    });
    univerAPI = made.univerAPI;
    univerAPI.createWorkbook(toSnapshot(book, info.title));

    const workbook = univerAPI.getActiveWorkbook();
    if (!writable && workbook && typeof workbook.setEditable === "function") workbook.setEditable(false);
    // The sheet's numbers and dates read in the chosen language (pt_BR shows
    // 1.234,57 where en_US shows 1,234.57); patterns are saved unchanged.
    if (workbook && typeof workbook.setNumfmtLocal === "function") workbook.setNumfmtLocal(numfmtOf(lang));
    if (langEl) {
      langEl.hidden = false;
      langEl.addEventListener("change", () => {
        if (!["pt-BR", "en-US"].includes(langEl.value)) return;
        lang = langEl.value;
        try { univerAPI.setLocale(localeOf(lang)); } catch { /* the UI keeps its language */ }
        const wb = univerAPI.getActiveWorkbook();
        if (wb && typeof wb.setNumfmtLocal === "function") wb.setNumfmtLocal(numfmtOf(lang));
        applyLanguage();
      });
    }
    applyLanguage();
  } catch {
    fail(S().undrawable);
    return;
  }

  if (bar) bar.hidden = false;
  if (bar) {
    const title = bar.querySelector(".name");
    if (title) title.textContent = pageTitle || S().title;
  }
  if (!writable) {
    setSaveState("readonly");
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
  setSaveState("saved", baseVersion);

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
    if (next === lastSaved) { setSaveState("saved", baseVersion); return; }
    saving = true;
    setSaveState("saving");
    try {
      const out = await door.publishFiles([{ path: info.entry, data: next }], baseVersion);
      baseVersion = (out && out.version && out.version.n) || baseVersion + 1;
      lastSaved = next;
      hideConflict();
      setSaveState("saved", baseVersion);
    } catch (e) {
      if (e && e.code === -32009) {
        const live = Number(e.data && e.data.liveVersion) || 0;
        showConflict(live || baseVersion + 1);
      } else if (e && e.code === -32001) {
        setSaveState("readonly");
      } else {
        setSaveState("unsaved");
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
