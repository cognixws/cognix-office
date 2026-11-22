import { Editor, Extension } from "@tiptap/core";
import StarterKit from "@tiptap/starter-kit";
import Underline from "@tiptap/extension-underline";
import TextAlign from "@tiptap/extension-text-align";
import TextStyle from "@tiptap/extension-text-style";
import Color from "@tiptap/extension-color";
import Highlight from "@tiptap/extension-highlight";
import Link from "@tiptap/extension-link";
import Image from "@tiptap/extension-image";
import Table from "@tiptap/extension-table";
import TableRow from "@tiptap/extension-table-row";
import TableCell from "@tiptap/extension-table-cell";
import TableHeader from "@tiptap/extension-table-header";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet } from "@tiptap/pm/view";
import { parseDocx, saveDocx } from "./docx-codec.js";
import { resolveLang, stringsFor } from "./i18n.js";
import { navigateFrame, openStandaloneDraft } from "./standalone-store.js";
import "./doc.css";

// Custom FontSize attribute on TextStyle so OOXML run font sizes survive edits.
const SizedTextStyle = TextStyle.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      fontSize: {
        default: null,
        parseHTML: (el) => el.style.fontSize || null,
        renderHTML: (attrs) => (attrs.fontSize ? { style: `font-size: ${attrs.fontSize}` } : {}),
      },
    };
  },
});

const searchPluginKey = new PluginKey("officeDocSearch");

function findMatches(doc, query) {
  const needle = String(query || "").toLowerCase();
  if (!needle) return [];
  const hits = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const hay = node.text.toLowerCase();
    let idx = 0;
    while ((idx = hay.indexOf(needle, idx)) >= 0) {
      hits.push({ from: pos + idx, to: pos + idx + needle.length });
      idx += needle.length;
    }
  });
  return hits;
}

const SearchHighlight = Extension.create({
  name: "officeSearch",
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: searchPluginKey,
        state: {
          init: () => ({ query: "", activeIndex: 0 }),
          apply(tr, prev) {
            const meta = tr.getMeta(searchPluginKey);
            return meta ? { ...prev, ...meta } : prev;
          },
        },
        props: {
          decorations(state) {
            const { query, activeIndex } = searchPluginKey.getState(state) || {};
            if (!query) return DecorationSet.empty;
            const matches = findMatches(state.doc, query);
            if (!matches.length) return DecorationSet.empty;
            return DecorationSet.create(
              state.doc,
              matches.map((m, i) =>
                Decoration.inline(m.from, m.to, {
                  class: i === activeIndex ? "doc-find-hit is-active" : "doc-find-hit",
                })
              )
            );
          },
        },
      }),
    ];
  },
});

const barEl = document.getElementById("bar");
const standaloneHeadEl = document.getElementById("standalone-head");
const backHubBtn = document.getElementById("back-hub");
const backHubLabelEl = document.getElementById("back-hub-label");
const standaloneTitleEl = document.getElementById("standalone-title");
const toolsEl = document.getElementById("tools");
const docNameEl = document.getElementById("doc-name");
const wordCountEl = document.getElementById("word-count");
const langEl = document.getElementById("lang");
const saveEl = document.getElementById("save");
const saveTextEl = document.getElementById("save-text");
const deskEl = document.getElementById("desk");
const paperEl = document.getElementById("paper");
const findBarEl = document.getElementById("find-bar");
const findInputEl = document.getElementById("find-input");
const replaceInputEl = document.getElementById("replace-input");
const findCountEl = document.getElementById("find-count");
const conflictEl = document.getElementById("conflict");
const conflictTextEl = document.getElementById("conflict-text");
const conflictOpenEl = conflictEl?.querySelector('[data-act="open"]');
const conflictKeepEl = conflictEl?.querySelector('[data-act="keep"]');
const stateEl = document.getElementById("state");
const stateTitleEl = document.getElementById("state-title");
const stateTextEl = document.getElementById("state-text");
const retryEl = document.getElementById("retry");
const styleSelectEl = document.getElementById("block-style");
const textColorEl = document.getElementById("text-color");
const tableMenuEl = document.getElementById("table-menu");
const tableBtnEl = document.querySelector('[data-cmd="tableMenu"]');

let lang = resolveLang("", typeof navigator === "undefined" ? "" : navigator.language);
let pageTitle = "";
const S = () => stringsFor(lang);

let editor = null;
let saveState = { kind: "saved", n: 1 };
let searchMatches = [];
let searchIdx = 0;

function isPhoneViewport() {
  return typeof window !== "undefined" && window.innerWidth <= 560;
}

function fail(title, text) {
  if (barEl) barEl.hidden = true;
  if (deskEl) deskEl.hidden = true;
  if (stateEl) {
    stateEl.hidden = false;
    if (stateTitleEl) stateTitleEl.textContent = title;
    if (stateTextEl) stateTextEl.textContent = text;
  }
  document.body.dataset.ready = "1";
}

retryEl?.addEventListener("click", () => location.reload());

function paintSave() {
  if (!saveEl || !saveTextEl) return;
  const s = S();
  if (saveState.kind === "saved") {
    saveEl.dataset.tone = "ok";
    saveTextEl.textContent = s.saved(saveState.n);
  } else if (saveState.kind === "savedLocal") {
    saveEl.dataset.tone = "ok";
    saveTextEl.textContent = s.savedLocal;
  } else if (saveState.kind === "tabOnly") {
    saveEl.dataset.tone = "";
    saveTextEl.textContent = s.tabOnly;
  } else if (saveState.kind === "saving") {
    saveEl.dataset.tone = "busy";
    saveTextEl.textContent = s.saving;
  } else if (saveState.kind === "unsaved") {
    saveEl.dataset.tone = "error";
    saveTextEl.textContent = s.unsaved;
  } else if (saveState.kind === "readonly") {
    saveEl.dataset.tone = "";
    saveTextEl.textContent = s.readOnly;
  } else {
    saveEl.dataset.tone = "error";
    saveTextEl.textContent = s.conflictSave(saveState.n);
  }
}

function setSaveState(kind, n) {
  saveState = { kind, n: n || saveState.n || 1 };
  paintSave();
}

function countWords(text) {
  const trimmed = String(text || "").trim();
  if (!trimmed) return 0;
  return trimmed.split(/\s+/).length;
}

function updateWordCount() {
  if (!editor || !wordCountEl) return;
  const words = countWords(editor.getText({ blockSeparator: " " }));
  wordCountEl.textContent = S().wordsCount(words);
}

function updateToolbarState() {
  if (!editor || !toolsEl) return;
  const checks = {
    bold: editor.isActive("bold"),
    italic: editor.isActive("italic"),
    underline: editor.isActive("underline"),
    strike: editor.isActive("strike"),
    bulletList: editor.isActive("bulletList"),
    orderedList: editor.isActive("orderedList"),
    alignLeft: editor.isActive({ textAlign: "left" }),
    alignCenter: editor.isActive({ textAlign: "center" }),
    alignRight: editor.isActive({ textAlign: "right" }),
    highlight: editor.isActive("highlight"),
  };
  for (const [cmd, active] of Object.entries(checks)) {
    const btn = toolsEl.querySelector(`[data-cmd="${cmd}"]`);
    if (btn) btn.setAttribute("aria-pressed", active ? "true" : "false");
  }
  if (styleSelectEl) {
    if (editor.isActive("heading", { level: 1 })) styleSelectEl.value = "h1";
    else if (editor.isActive("heading", { level: 2 })) styleSelectEl.value = "h2";
    else if (editor.isActive("heading", { level: 3 })) styleSelectEl.value = "h3";
    else if (editor.isActive("blockquote")) styleSelectEl.value = "quote";
    else styleSelectEl.value = "p";
  }
  const inTable = editor.isActive("table");
  for (const el of document.querySelectorAll("[data-table-only]")) {
    el.hidden = !inTable;
  }
}

function applyLanguage() {
  const s = S();
  document.documentElement.lang = lang;
  if (langEl) {
    langEl.value = lang;
    langEl.setAttribute("aria-label", s.language);
    langEl.title = s.language;
  }
  if (docNameEl) docNameEl.textContent = pageTitle || s.docTitle;
  if (retryEl) retryEl.textContent = s.retry;
  if (conflictOpenEl) conflictOpenEl.textContent = s.open;
  if (conflictKeepEl) conflictKeepEl.textContent = s.keep;
  const live = Number(conflictEl?.dataset.live) || 0;
  if (conflictEl && !conflictEl.hidden && live && conflictTextEl) {
    conflictTextEl.textContent = s.conflict(live);
  }
  if (styleSelectEl) {
    const map = { p: s.styleNormal, h1: s.styleH1, h2: s.styleH2, h3: s.styleH3, quote: s.blockquote };
    for (const opt of styleSelectEl.options) {
      if (opt.value in map) opt.textContent = map[opt.value];
    }
  }
  const tipMap = {
    undo: s.undo,
    redo: s.redo,
    bold: s.bold,
    italic: s.italic,
    underline: s.underline,
    strike: s.strike,
    bulletList: s.bulletList,
    orderedList: s.orderedList,
    alignLeft: s.alignLeft,
    alignCenter: s.alignCenter,
    alignRight: s.alignRight,
    highlight: s.highlightColor,
    tableMenu: s.insertTable,
    find: s.find,
  };
  for (const [cmd, label] of Object.entries(tipMap)) {
    const btn = document.querySelector(`[data-cmd="${cmd}"]`);
    if (btn) {
      btn.title = label;
      btn.setAttribute("aria-label", label);
    }
  }
  if (textColorEl) {
    textColorEl.title = s.textColor;
    textColorEl.setAttribute("aria-label", s.textColor);
  }
  const tblLabels = {
    insertTable: s.insertTable,
    addRow: s.addRow,
    deleteRow: s.deleteRow,
    addColumn: s.addColumn,
    deleteColumn: s.deleteColumn,
    deleteTable: s.deleteTable,
  };
  for (const [cmd, label] of Object.entries(tblLabels)) {
    const btn = tableMenuEl?.querySelector(`[data-tbl="${cmd}"]`);
    if (btn) btn.textContent = label;
  }
  if (findInputEl) findInputEl.placeholder = s.findPlaceholder;
  if (replaceInputEl) replaceInputEl.placeholder = s.replacePlaceholder;
  const repBtn = document.getElementById("find-replace");
  const repAllBtn = document.getElementById("find-replace-all");
  const closeFindBtn = document.getElementById("find-close");
  if (repBtn) repBtn.textContent = s.replace;
  if (repAllBtn) repAllBtn.textContent = s.replaceAll;
  if (closeFindBtn) {
    closeFindBtn.title = s.closeFind;
    closeFindBtn.setAttribute("aria-label", s.closeFind);
  }
  if (findBarEl && !findBarEl.hidden) refreshSearch(false);
  if (backHubBtn) backHubBtn.setAttribute("aria-label", s.backToHubAria);
  if (backHubLabelEl) backHubLabelEl.textContent = s.backToHub;
  if (standaloneTitleEl) standaloneTitleEl.setAttribute("aria-label", s.fileTitleAria);
  updateWordCount();
  paintSave();
}

function refreshSearch(scrollIntoView = true) {
  if (!editor) return;
  const q = findInputEl ? findInputEl.value : "";
  searchMatches = findMatches(editor.state.doc, q);
  if (searchIdx >= searchMatches.length) searchIdx = 0;
  if (searchIdx < 0) searchIdx = Math.max(0, searchMatches.length - 1);
  editor.view.dispatch(editor.state.tr.setMeta(searchPluginKey, { query: q, activeIndex: searchIdx }));
  if (findCountEl) {
    findCountEl.textContent = q ? S().matches(searchMatches.length ? searchIdx + 1 : 0, searchMatches.length) : "";
  }
  if (scrollIntoView && searchMatches[searchIdx]) {
    const { from } = searchMatches[searchIdx];
    const domAt = editor.view.domAtPos(from);
    const el = domAt.node?.nodeType === 1 ? domAt.node : domAt.node?.parentElement;
    el?.scrollIntoView({ block: "nearest" });
  }
}

function openFindBar(writable) {
  if (!findBarEl) return;
  findBarEl.hidden = false;
  for (const el of findBarEl.querySelectorAll("[data-edit-only]")) {
    el.hidden = !writable || isPhoneViewport();
  }
  findInputEl?.focus();
  findInputEl?.select();
  refreshSearch(true);
}

function closeFindBar() {
  if (!findBarEl) return;
  findBarEl.hidden = true;
  if (editor) {
    editor.view.dispatch(editor.state.tr.setMeta(searchPluginKey, { query: "", activeIndex: 0 }));
    editor.commands.focus();
  }
}

function showConflict(liveVersion) {
  if (!conflictEl) return;
  conflictEl.hidden = false;
  if (conflictTextEl) conflictTextEl.textContent = S().conflict(liveVersion);
  conflictEl.dataset.live = String(liveVersion);
  setSaveState("conflict", liveVersion);
}

function hideConflict() {
  if (!conflictEl) return;
  conflictEl.hidden = true;
  delete conflictEl.dataset.live;
}

async function main() {
  const draft = await openStandaloneDraft("docx");
  let door = null;
  let linksDoor = null;
  let info = null;
  let writableCap = false;
  let base64 = "";

  try {
    linksDoor = await picode.use("links").catch(() => null);
  } catch {
    linksDoor = null;
  }

  if (draft) {
    pageTitle = draft.title || S().untitledDoc;
    writableCap = true;
    base64 = draft.data;
    if (standaloneHeadEl) standaloneHeadEl.hidden = false;
    if (standaloneTitleEl) standaloneTitleEl.value = pageTitle;
    backHubBtn?.addEventListener("click", () => draft.back());
  } else {
    try {
      door = await picode.use("artifact");
    } catch {
      navigateFrame("office.html");
      return;
    }
    if (!door || typeof door.get !== "function") {
      navigateFrame("office.html");
      return;
    }

    try {
      info = await door.get();
    } catch {
      fail(S().docUnreadableTitle, S().docUnreadable);
      return;
    }
    if (!info || !info.entry) {
      fail(S().docUnreadableTitle, S().docUnreadable);
      return;
    }

    pageTitle = info.title || "";
    writableCap = Array.isArray(info.capabilities) && info.capabilities.includes("publish");

    let file;
    try {
      file = await door.read(info.entry);
    } catch {
      fail(S().docUnreadableTitle, S().docUnreadable);
      return;
    }
    base64 = (file && file.base64) || "";
    if (!base64) {
      fail(S().docUnreadableTitle, S().docUnreadable);
      return;
    }
  }

  let session;
  try {
    session = parseDocx(base64);
  } catch {
    fail(S().docUnreadableTitle, S().docUnreadable);
    return;
  }

  const editableNow = () => writableCap && !isPhoneViewport();

  if (paperEl && session.page) {
    paperEl.style.setProperty("--doc-max-w", `${Math.min(900, Math.max(640, session.page.widthPx))}px`);
  }

  let baseVersion = (info && info.version) || 1;
  let lastSaved = "";
  let lastSavedTitle = pageTitle;
  let timer = 0;
  let saving = false;
  let pending = false;

  editor = new Editor({
    element: paperEl,
    editable: editableNow(),
    content: session.doc,
    extensions: [
      StarterKit.configure({ heading: { levels: [1, 2, 3, 4] } }),
      Underline,
      SizedTextStyle,
      Color,
      Highlight.configure({ multicolor: true }),
      TextAlign.configure({ types: ["heading", "paragraph"] }),
      Link.configure({ openOnClick: false }),
      Image.configure({ inline: true, allowBase64: true }),
      Table.configure({ resizable: false }),
      TableRow,
      TableHeader,
      TableCell,
      SearchHighlight,
    ],
    onSelectionUpdate() {
      updateToolbarState();
    },
    onUpdate({ transaction }) {
      updateToolbarState();
      updateWordCount();
      if (findBarEl && !findBarEl.hidden) refreshSearch(false);
      if (transaction.docChanged && editableNow()) {
        scheduleSave();
      }
    },
  });

  try {
    lastSaved = saveDocx(session, editor.getJSON());
  } catch {
    lastSaved = base64;
  }

  paperEl.addEventListener("click", (e) => {
    const a = e.target.closest("a[href]");
    if (a && (!editableNow() || e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      const href = a.getAttribute("href");
      if (href && /^https?:\/\//i.test(href) && linksDoor) {
        linksDoor.open(href).catch(() => {});
      }
    }
  });

  async function save() {
    clearTimeout(timer);
    if (!editableNow() && !draft) return;
    if (saving) {
      pending = true;
      return;
    }
    let next;
    try {
      next = saveDocx(session, editor.getJSON());
    } catch {
      setSaveState("unsaved");
      return;
    }
    const nextTitle = draft && standaloneTitleEl ? standaloneTitleEl.value.trim() || S().untitledDoc : pageTitle;
    if (next === lastSaved && nextTitle === lastSavedTitle) {
      setSaveState(draft ? (draft.persisted ? "savedLocal" : "tabOnly") : "saved", baseVersion);
      return;
    }
    saving = true;
    setSaveState("saving");
    try {
      if (draft) {
        const res = await draft.save(next, nextTitle);
        lastSaved = next;
        lastSavedTitle = nextTitle;
        pageTitle = nextTitle;
        setSaveState(res.persisted ? "savedLocal" : "tabOnly");
      } else {
        const out = await door.publishFiles([{ path: info.entry, data: next }], baseVersion);
        baseVersion = (out && out.version && out.version.n) || baseVersion + 1;
        lastSaved = next;
        hideConflict();
        setSaveState("saved", baseVersion);
      }
    } catch (e) {
      if (!draft && e && e.code === -32009) {
        const live = Number(e.data && e.data.liveVersion) || 0;
        showConflict(live || baseVersion + 1);
      } else if (!draft && e && e.code === -32001) {
        setSaveState("readonly");
      } else {
        setSaveState("unsaved");
      }
    } finally {
      saving = false;
      if (pending) {
        pending = false;
        timer = setTimeout(save, 200);
      }
    }
  }

  function scheduleSave() {
    clearTimeout(timer);
    timer = setTimeout(save, 1500);
  }

  standaloneTitleEl?.addEventListener("input", () => scheduleSave());

  // Toolbar bindings
  toolsEl?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-cmd]");
    if (!btn || !editor) return;
    const cmd = btn.dataset.cmd;
    if (cmd === "find") {
      if (findBarEl?.hidden) openFindBar(writableCap);
      else closeFindBar();
      return;
    }
    if (!editableNow()) return;
    const chain = editor.chain().focus();
    if (cmd === "undo") chain.undo().run();
    else if (cmd === "redo") chain.redo().run();
    else if (cmd === "bold") chain.toggleBold().run();
    else if (cmd === "italic") chain.toggleItalic().run();
    else if (cmd === "underline") chain.toggleUnderline().run();
    else if (cmd === "strike") chain.toggleStrike().run();
    else if (cmd === "bulletList") chain.toggleBulletList().run();
    else if (cmd === "orderedList") chain.toggleOrderedList().run();
    else if (cmd === "alignLeft") chain.setTextAlign("left").run();
    else if (cmd === "alignCenter") chain.setTextAlign("center").run();
    else if (cmd === "alignRight") chain.setTextAlign("right").run();
    else if (cmd === "highlight") chain.toggleHighlight({ color: "#FEF08A" }).run();
    else if (cmd === "tableMenu" && tableMenuEl) {
      const open = !tableMenuEl.hidden;
      tableMenuEl.hidden = open;
      if (!open && tableBtnEl) {
        const rect = tableBtnEl.getBoundingClientRect();
        tableMenuEl.style.left = `${Math.max(8, Math.min(window.innerWidth - 190, rect.left))}px`;
      }
    }
    updateToolbarState();
  });

  document.addEventListener("click", (e) => {
    if (tableMenuEl && !tableMenuEl.hidden && !e.target.closest("#table-menu") && !e.target.closest('[data-cmd="tableMenu"]')) {
      tableMenuEl.hidden = true;
    }
  });

  tableMenuEl?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-tbl]");
    if (!btn || !editor || !editableNow()) return;
    const act = btn.dataset.tbl;
    const ch = editor.chain().focus();
    if (act === "insertTable") ch.insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
    else if (act === "addRow") ch.addRowAfter().run();
    else if (act === "deleteRow") ch.deleteRow().run();
    else if (act === "addColumn") ch.addColumnAfter().run();
    else if (act === "deleteColumn") ch.deleteColumn().run();
    else if (act === "deleteTable") ch.deleteTable().run();
    tableMenuEl.hidden = true;
    updateToolbarState();
  });

  styleSelectEl?.addEventListener("change", () => {
    if (!editor || !editableNow()) return;
    const v = styleSelectEl.value;
    const ch = editor.chain().focus();
    if (v === "p") ch.setParagraph().run();
    else if (v === "h1") ch.toggleHeading({ level: 1 }).run();
    else if (v === "h2") ch.toggleHeading({ level: 2 }).run();
    else if (v === "h3") ch.toggleHeading({ level: 3 }).run();
    else if (v === "quote") ch.toggleBlockquote().run();
    updateToolbarState();
  });

  textColorEl?.addEventListener("input", () => {
    if (!editor || !editableNow()) return;
    editor.chain().focus().setColor(textColorEl.value).run();
  });

  langEl?.addEventListener("change", () => {
    if (!["pt-BR", "en-US"].includes(langEl.value)) return;
    lang = langEl.value;
    applyLanguage();
  });

  // Find & Replace events
  findInputEl?.addEventListener("input", () => {
    searchIdx = 0;
    refreshSearch(true);
  });
  findInputEl?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      searchIdx += e.shiftKey ? -1 : 1;
      refreshSearch(true);
    } else if (e.key === "Escape") {
      e.preventDefault();
      closeFindBar();
    }
  });
  replaceInputEl?.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      e.preventDefault();
      closeFindBar();
    }
  });
  document.getElementById("find-prev")?.addEventListener("click", () => {
    searchIdx -= 1;
    refreshSearch(true);
  });
  document.getElementById("find-next")?.addEventListener("click", () => {
    searchIdx += 1;
    refreshSearch(true);
  });
  document.getElementById("find-close")?.addEventListener("click", () => closeFindBar());
  document.getElementById("find-replace")?.addEventListener("click", () => {
    if (!editableNow() || !searchMatches.length) return;
    const hit = searchMatches[searchIdx];
    if (!hit) return;
    const rep = replaceInputEl?.value || "";
    editor.view.dispatch(editor.state.tr.insertText(rep, hit.from, hit.to));
    refreshSearch(true);
  });
  document.getElementById("find-replace-all")?.addEventListener("click", () => {
    if (!editableNow() || !searchMatches.length) return;
    const rep = replaceInputEl?.value || "";
    let tr = editor.state.tr;
    for (let i = searchMatches.length - 1; i >= 0; i--) {
      tr = tr.insertText(rep, searchMatches[i].from, searchMatches[i].to);
    }
    editor.view.dispatch(tr);
    refreshSearch(false);
  });

  window.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
      e.preventDefault();
      openFindBar(writableCap);
    } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      if (editableNow()) void save();
    }
  });

  window.addEventListener("resize", () => {
    const canEdit = editableNow();
    if (editor && editor.isEditable !== canEdit) {
      editor.setEditable(canEdit);
    }
    if (toolsEl) {
      for (const el of toolsEl.querySelectorAll("[data-edit-tool]")) {
        el.hidden = !canEdit;
      }
    }
    const savedKind = draft ? (draft.persisted ? "savedLocal" : "tabOnly") : "saved";
    if (!canEdit && saveState.kind !== "conflict") setSaveState("readonly");
    else if (canEdit && saveState.kind === "readonly") setSaveState(savedKind, baseVersion);
  });

  conflictEl?.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-act]");
    if (!b) return;
    if (b.dataset.act === "open") {
      location.reload();
      return;
    }
    if (b.dataset.act === "keep") {
      baseVersion = Number(conflictEl.dataset.live) || baseVersion;
      lastSaved = "";
      hideConflict();
      await save();
    }
  });

  if (barEl) barEl.hidden = false;
  if (deskEl) deskEl.hidden = false;
  for (const el of toolsEl?.querySelectorAll("[data-edit-tool]") || []) {
    el.hidden = !editableNow();
  }
  const initialSavedKind = draft ? (draft.persisted ? "savedLocal" : "tabOnly") : "saved";
  setSaveState(editableNow() ? initialSavedKind : "readonly", baseVersion);
  applyLanguage();
  updateToolbarState();
  document.body.dataset.ready = "1";
}

main();
