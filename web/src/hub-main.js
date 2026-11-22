import { bytesToBase64, createBlankXlsx, detectOoxmlKind } from "./ooxml.js";
import { createBlankDocx } from "./docx-codec.js";
import { createBlankPptx } from "./pptx-codec.js";
import { resolveLang, stringsFor } from "./i18n.js";
import {
  deleteSavedFile,
  editorPageForKind,
  getStateDoor,
  listSavedFiles,
  makeFileId,
  navigateFrame,
  saveStandaloneRecord,
} from "./standalone-store.js";
import "./hub.css";

const hubEl = document.getElementById("hub");
const appTitleEl = document.getElementById("hub-title");
const langEl = document.getElementById("hub-lang");
const newSheetBtn = document.getElementById("btn-new-sheet");
const newDocBtn = document.getElementById("btn-new-doc");
const newDeckBtn = document.getElementById("btn-new-deck");
const openFileBtn = document.getElementById("btn-open-file");
const fileInputEl = document.getElementById("file-picker");
const alertEl = document.getElementById("hub-alert");
const alertTextEl = document.getElementById("hub-alert-text");
const alertCloseBtn = document.getElementById("hub-alert-close");
const emptyEl = document.getElementById("hub-empty");
const emptyTitleEl = document.getElementById("hub-empty-title");
const emptyBodyEl = document.getElementById("hub-empty-body");
const emptyActionEl = document.getElementById("hub-empty-action");
const savedWrapEl = document.getElementById("hub-saved-wrap");
const savedHeadingEl = document.getElementById("hub-saved-heading");
const savedListEl = document.getElementById("hub-saved-list");

let lang = resolveLang("", typeof navigator === "undefined" ? "" : navigator.language);
const S = () => stringsFor(lang);
let stateDoor = null;
let files = [];

function formatShortTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  try {
    return d.toLocaleDateString(lang, { month: "short", day: "numeric" }) + " " +
      d.toLocaleTimeString(lang, { hour: "2-digit", minute: "2-digit" });
  } catch {
    return iso.slice(0, 16).replace("T", " ");
  }
}

function showAlert(msg) {
  if (!alertEl || !alertTextEl) return;
  alertTextEl.textContent = msg;
  alertEl.hidden = false;
}

function hideAlert() {
  if (alertEl) alertEl.hidden = true;
}

function renderList() {
  const s = S();
  if (!files.length) {
    if (savedWrapEl) savedWrapEl.hidden = true;
    if (emptyEl) emptyEl.hidden = false;
    return;
  }
  if (emptyEl) emptyEl.hidden = true;
  if (savedWrapEl) savedWrapEl.hidden = false;
  if (!savedListEl) return;
  savedListEl.replaceChildren();

  for (const item of files) {
    const row = document.createElement("div");
    row.className = "file-row";

    const openBtn = document.createElement("button");
    openBtn.type = "button";
    openBtn.className = "file-row-open";
    openBtn.addEventListener("click", () => {
      const target = editorPageForKind(item.kind);
      if (target) navigateFrame(`${target}?file=${encodeURIComponent(item.id)}`);
    });

    const pill = document.createElement("span");
    pill.className = "file-kind-pill";
    pill.dataset.kind = item.kind;
    pill.textContent = "." + item.kind;

    const textWrap = document.createElement("span");
    textWrap.className = "file-row-text";

    const nameSpan = document.createElement("span");
    nameSpan.className = "file-row-name";
    nameSpan.textContent = item.title || s.untitledDoc;

    const kb = Math.max(1, Math.round((item.sizeBytes || 1024) / 1024));
    const when = formatShortTime(item.updatedAt);
    const metaSpan = document.createElement("span");
    metaSpan.className = "file-row-meta";
    const kindSpan = document.createElement("span");
    kindSpan.className = "file-row-kind";
    kindSpan.textContent = s.kindLabel(item.kind) + " · ";
    const detailSpan = document.createElement("span");
    detailSpan.textContent = when ? s.fileMeta(kb, when) : `${kb} KB`;
    metaSpan.append(kindSpan, detailSpan);

    textWrap.append(nameSpan, metaSpan);
    openBtn.append(pill, textWrap);

    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "file-row-del";
    delBtn.setAttribute("aria-label", s.removeFileAria(item.title || item.id));
    delBtn.title = s.removeFile;
    delBtn.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>`;
    delBtn.addEventListener("click", async (ev) => {
      ev.stopPropagation();
      await deleteSavedFile(stateDoor, item.id);
      files = files.filter((f) => f.id !== item.id);
      renderList();
    });

    row.append(openBtn, delBtn);
    savedListEl.append(row);
  }
}

function applyLanguage() {
  const s = S();
  document.documentElement.lang = lang;
  document.title = s.appTitle;
  if (appTitleEl) appTitleEl.textContent = s.appTitle;
  if (langEl) {
    langEl.value = lang;
    langEl.setAttribute("aria-label", s.language);
    langEl.title = s.language;
  }
  const setBtnLabel = (btn, text) => {
    const span = btn?.querySelector("[data-label]");
    if (span) span.textContent = text;
  };
  setBtnLabel(newSheetBtn, s.newSheet);
  setBtnLabel(newDocBtn, s.newDoc);
  setBtnLabel(newDeckBtn, s.newDeck);
  setBtnLabel(openFileBtn, s.openFile);
  if (emptyTitleEl) emptyTitleEl.textContent = s.emptyHubTitle;
  if (emptyBodyEl) emptyBodyEl.textContent = s.emptyHubBody;
  if (emptyActionEl) emptyActionEl.textContent = s.openArtifacts;
  if (savedHeadingEl) savedHeadingEl.textContent = s.recentHeading;
  if (alertCloseBtn) alertCloseBtn.setAttribute("aria-label", s.dismissAlert);
  renderList();
}

async function openCreated(kind, title, base64) {
  hideAlert();
  const id = makeFileId();
  const { persisted } = await saveStandaloneRecord(stateDoor, {
    id,
    kind,
    title,
    data: base64,
  });
  const target = editorPageForKind(kind);
  const param = persisted ? "file" : "transient";
  navigateFrame(`${target}?${param}=${encodeURIComponent(id)}`);
}

async function handlePickedFile(file) {
  if (!file) return;
  hideAlert();
  let u8;
  try {
    u8 = new Uint8Array(await file.arrayBuffer());
  } catch {
    showAlert(S().badFileError(file.name));
    return;
  }
  const kind = detectOoxmlKind(u8);
  if (!kind) {
    showAlert(S().badFileError(file.name));
    return;
  }
  const rawName = String(file.name || "").replace(/\.(xlsx|docx|pptx)$/i, "").trim();
  const fallbackTitle = kind === "xlsx" ? S().untitledSheet : kind === "docx" ? S().untitledDoc : S().untitledDeck;
  await openCreated(kind, rawName || fallbackTitle, bytesToBase64(u8));
}

async function resolveArtifactRoute() {
  try {
    if (typeof picode === "undefined" || !picode || typeof picode.use !== "function") return "";
    const ctxDoor = await picode.use("context");
    const ctx = ctxDoor && typeof ctxDoor.current === "function" ? ctxDoor.current() : null;
    if (ctx && ctx.source === "artifact" && ctx.artifact && ctx.artifact.kind) {
      const target = editorPageForKind(ctx.artifact.kind);
      if (target) return target;
    }
    const artDoor = await picode.use("artifact");
    if (artDoor && typeof artDoor.get === "function") {
      const info = await artDoor.get();
      if (info && info.kind) {
        const target = editorPageForKind(info.kind);
        if (target) return target;
      }
      const ext = String(info?.entry || "").split(".").pop()?.toLowerCase() || "";
      const byExt = editorPageForKind(ext);
      if (byExt) return byExt;
    }
  } catch {
    // standalone mode
  }
  return "";
}

async function main() {
  const artifactTarget = await resolveArtifactRoute();
  if (artifactTarget) {
    navigateFrame(artifactTarget);
    return;
  }

  stateDoor = await getStateDoor();
  files = await listSavedFiles(stateDoor);

  if (stateDoor && typeof stateDoor.onChange === "function") {
    stateDoor.onChange(async () => {
      files = await listSavedFiles(stateDoor);
      renderList();
    });
  }

  langEl?.addEventListener("change", () => {
    if (["pt-BR", "en-US"].includes(langEl.value)) {
      lang = langEl.value;
      applyLanguage();
    }
  });

  alertCloseBtn?.addEventListener("click", hideAlert);

  newSheetBtn?.addEventListener("click", () => {
    const sheetTab = lang === "pt-BR" ? "Planilha1" : "Sheet1";
    void openCreated("xlsx", S().untitledSheet, createBlankXlsx(sheetTab));
  });

  newDocBtn?.addEventListener("click", () => {
    void openCreated("docx", S().untitledDoc, createBlankDocx());
  });

  newDeckBtn?.addEventListener("click", () => {
    const slideTitle = lang === "pt-BR" ? "Título da apresentação" : "Presentation title";
    const initialSlides = [
      {
        sourceSlidePath: "",
        layoutPath: "ppt/slideLayouts/slideLayout1.xml",
        background: { kind: "solid", css: "#FFFFFF", hex: "#FFFFFF" },
        elements: [
          {
            uid: "s0-1",
            kind: "shape",
            isNew: true,
            name: "Title 1",
            x: 96,
            y: 220,
            w: 1088,
            h: 120,
            rot: 0,
            geom: "rect",
            fill: "transparent",
            fillHex: "",
            vAlign: "middle",
            padding: { top: 8, right: 12, bottom: 8, left: 12 },
            paragraphs: [{ align: "center", bullet: false, level: 0, runs: [{ text: slideTitle, fontSize: 40, bold: true, color: "#181B22" }] }],
          },
        ],
        notes: "",
        dirty: true,
        notesDirty: false,
      },
    ];
    void openCreated("pptx", S().untitledDeck, createBlankPptx(initialSlides));
  });

  openFileBtn?.addEventListener("click", () => {
    if (fileInputEl) {
      fileInputEl.value = "";
      fileInputEl.click();
    }
  });

  fileInputEl?.addEventListener("change", () => {
    const f = fileInputEl.files && fileInputEl.files[0];
    if (f) void handlePickedFile(f);
  });

  // Drag and drop .xlsx, .docx, .pptx anywhere in the page
  window.addEventListener("dragover", (e) => {
    e.preventDefault();
    hubEl?.classList.add("is-dragover");
    emptyEl?.classList.add("is-dragover");
  });
  window.addEventListener("dragleave", (e) => {
    if (!e.relatedTarget) {
      hubEl?.classList.remove("is-dragover");
      emptyEl?.classList.remove("is-dragover");
    }
  });
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    hubEl?.classList.remove("is-dragover");
    emptyEl?.classList.remove("is-dragover");
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) void handlePickedFile(f);
  });

  applyLanguage();
  if (hubEl) hubEl.hidden = false;
  document.body.dataset.ready = "1";
}

void main();
