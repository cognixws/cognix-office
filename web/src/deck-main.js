import { parsePptx, savePptx } from "./pptx-codec.js";
import { resolveLang, stringsFor } from "./i18n.js";
import { navigateFrame, openStandaloneDraft } from "./standalone-store.js";
import "./deck.css";

const barEl = document.getElementById("bar");
const standaloneHeadEl = document.getElementById("standalone-head");
const backHubBtn = document.getElementById("back-hub");
const backHubLabelEl = document.getElementById("back-hub-label");
const standaloneTitleEl = document.getElementById("standalone-title");
const toolsEl = document.getElementById("tools");
const deckNameEl = document.getElementById("deck-name");
const slideCounterEl = document.getElementById("slide-counter");
const prevSlideBtn = document.getElementById("prev-slide");
const nextSlideBtn = document.getElementById("next-slide");
const langEl = document.getElementById("lang");
const saveEl = document.getElementById("save");
const saveTextEl = document.getElementById("save-text");

const workspaceEl = document.getElementById("workspace");
const railEl = document.getElementById("rail");
const stageWrapEl = document.getElementById("stage-wrap");
const stageBoxEl = document.getElementById("stage-box");
const slideCanvasEl = document.getElementById("slide-canvas");
const emptyDeckEl = document.getElementById("empty-deck");
const emptyTitleEl = document.getElementById("empty-title");
const emptyBodyEl = document.getElementById("empty-body");
const emptyAddBtn = document.getElementById("empty-add");

const notesPanelEl = document.getElementById("notes-panel");
const notesLabelEl = document.getElementById("notes-label");
const notesInputEl = document.getElementById("notes-input");

const findBarEl = document.getElementById("find-bar");
const findInputEl = document.getElementById("find-input");
const findCountEl = document.getElementById("find-count");

const fontSizeEl = document.getElementById("font-size");
const textColorEl = document.getElementById("text-color");
const fillColorEl = document.getElementById("fill-color");

const conflictEl = document.getElementById("conflict");
const conflictTextEl = document.getElementById("conflict-text");
const conflictOpenEl = conflictEl?.querySelector('[data-act="open"]');
const conflictKeepEl = conflictEl?.querySelector('[data-act="keep"]');

const stateEl = document.getElementById("state");
const stateTitleEl = document.getElementById("state-title");
const stateTextEl = document.getElementById("state-text");
const retryEl = document.getElementById("retry");

let lang = resolveLang("", typeof navigator === "undefined" ? "" : navigator.language);
let pageTitle = "";
const S = () => stringsFor(lang);

let session = null;
let slides = [];
let activeSlideIdx = 0;
let selectedUid = null;
let editingUid = null;
let editingCell = null; // { r, c } for table cell editing
let currentScale = 1;
let writableCap = false;
let saveState = { kind: "saved", n: 1 };
let undoStack = [];
let redoStack = [];
let uidCounter = 1000;
let findHits = [];
let findIdx = 0;

function isPhoneViewport() {
  return typeof window !== "undefined" && window.innerWidth <= 560;
}

function editableNow() {
  return writableCap && !isPhoneViewport();
}

function deepCloneSlides(list) {
  return structuredClone(list);
}

function pushUndo() {
  undoStack.push({ slides: deepCloneSlides(slides), activeSlideIdx, selectedUid });
  if (undoStack.length > 40) undoStack.shift();
  redoStack = [];
}

function fail(title, text) {
  if (barEl) barEl.hidden = true;
  if (workspaceEl) workspaceEl.hidden = true;
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

function getActiveSlide() {
  return slides[activeSlideIdx] || null;
}

function getSelectedElement() {
  const slide = getActiveSlide();
  if (!slide || !selectedUid) return null;
  return slide.elements.find((e) => e.uid === selectedUid) || null;
}

function shapeCssGeometry(geom) {
  if (geom === "ellipse") return { borderRadius: "50%" };
  if (geom === "roundRect") return { borderRadius: "12px" };
  return { borderRadius: "0px" };
}

function renderSlideLayer(container, slide, interactive) {
  container.replaceChildren();
  if (!slide || !session) return;

  // Slide background
  container.style.width = `${session.slideW}px`;
  container.style.height = `${session.slideH}px`;
  if (slide.background?.imageUrl) {
    container.style.background = `url("${slide.background.imageUrl}") center / cover no-repeat #FFFFFF`;
  } else if (slide.background?.css) {
    container.style.background = slide.background.css;
  } else {
    container.style.background = "#FFFFFF";
  }

  const drawElement = (el, isBackdrop) => {
    const box = document.createElement("div");
    box.className = "sld-el" + (isBackdrop ? " is-backdrop" : "");
    box.dataset.uid = el.uid;
    box.style.left = `${el.x}px`;
    box.style.top = `${el.y}px`;
    box.style.width = `${el.w}px`;
    box.style.height = `${el.h}px`;
    if (el.rot) box.style.transform = `rotate(${el.rot}deg)`;

    if (el.kind === "image") {
      box.classList.add("is-image");
      const img = document.createElement("img");
      img.src = el.src;
      img.alt = el.name || "";
      img.draggable = false;
      if (el.crop) {
        img.style.clipPath = `inset(${el.crop.t}% ${el.crop.r}% ${el.crop.b}% ${el.crop.l}%)`;
      }
      box.appendChild(img);
    } else if (el.kind === "connector") {
      box.classList.add("is-connector");
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("viewBox", `0 0 ${Math.max(2, el.w)} ${Math.max(2, el.h)}`);
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      line.setAttribute("x1", el.flipH ? String(el.w) : "0");
      line.setAttribute("y1", el.flipV ? String(el.h) : "0");
      line.setAttribute("x2", el.flipH ? "0" : String(el.w));
      line.setAttribute("y2", el.flipV ? "0" : String(el.h));
      line.setAttribute("stroke", el.stroke?.color || "#667085");
      line.setAttribute("stroke-width", String(el.stroke?.width || 1.5));
      if (el.stroke?.dashed) line.setAttribute("stroke-dasharray", "6 4");
      svg.appendChild(line);
      box.appendChild(svg);
    } else if (el.kind === "table") {
      box.classList.add("is-table");
      const tbl = document.createElement("table");
      const totalColW = (el.colWidths || []).reduce((a, b) => a + b, 0) || el.w;
      if (el.colWidths && el.colWidths.length) {
        const cg = document.createElement("colgroup");
        for (const cw of el.colWidths) {
          const col = document.createElement("col");
          col.style.width = `${((cw / totalColW) * 100).toFixed(2)}%`;
          cg.appendChild(col);
        }
        tbl.appendChild(cg);
      }
      const tbody = document.createElement("tbody");
      (el.rows || []).forEach((row, rIdx) => {
        const tr = document.createElement("tr");
        tr.style.height = `${row.height}px`;
        (row.cells || []).forEach((cell, cIdx) => {
          const td = document.createElement("td");
          if (cell.colspan > 1) td.colSpan = cell.colspan;
          if (cell.rowspan > 1) td.rowSpan = cell.rowspan;
          td.style.background = cell.fill || "transparent";
          td.style.color = cell.color || "#181B22";
          td.style.fontSize = `${cell.fontSize || 14}px`;
          td.style.fontWeight = cell.bold ? "700" : "400";
          td.style.textAlign = cell.align || "left";
          td.textContent = cell.text || "";
          if (interactive && !isBackdrop && editableNow()) {
            const isEditingThis = editingUid === el.uid && editingCell && editingCell.r === rIdx && editingCell.c === cIdx;
            if (isEditingThis) {
              td.contentEditable = "true";
              td.classList.add("is-editing-cell");
              td.addEventListener("input", () => {
                cell.text = td.innerText;
                el.textChanged = true;
                slide.dirty = true;
                onDeckMutated(false);
              });
            }
            td.addEventListener("dblclick", (ev) => {
              ev.stopPropagation();
              if (!editableNow()) return;
              pushUndo();
              selectedUid = el.uid;
              editingUid = el.uid;
              editingCell = { r: rIdx, c: cIdx };
              renderMainStage();
              const targetTd = slideCanvasEl.querySelector(`.sld-el[data-uid="${el.uid}"] tr:nth-child(${rIdx + 1}) td:nth-child(${cIdx + 1})`);
              targetTd?.focus();
            });
          }
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      tbl.appendChild(tbody);
      box.appendChild(tbl);
    } else if (el.kind === "shape") {
      box.classList.add("is-shape");
      const geomCss = shapeCssGeometry(el.geom);
      box.style.borderRadius = geomCss.borderRadius;
      if (el.geom === "line") {
        box.style.background = "transparent";
        box.style.borderTop = `${el.stroke?.width || 2}px ${el.stroke?.dashed ? "dashed" : "solid"} ${el.stroke?.color || "#181B22"}`;
      } else {
        box.style.background = el.fill || "transparent";
        if (el.stroke) {
          box.style.border = `${el.stroke.width}px ${el.stroke.dashed ? "dashed" : "solid"} ${el.stroke.color}`;
        }
      }

      const textWrap = document.createElement("div");
      textWrap.className = "sld-text";
      textWrap.style.justifyContent = el.vAlign === "middle" ? "center" : el.vAlign === "bottom" ? "flex-end" : "flex-start";
      textWrap.style.padding = `${el.padding?.top ?? 6}px ${el.padding?.right ?? 10}px ${el.padding?.bottom ?? 6}px ${el.padding?.left ?? 10}px`;

      const isEditingThis = interactive && !isBackdrop && editingUid === el.uid && editableNow();
      if (isEditingThis) {
        textWrap.contentEditable = "true";
        textWrap.classList.add("is-editing");
      }

      for (const p of el.paragraphs || []) {
        const pEl = document.createElement("div");
        pEl.className = "sld-p" + (p.bullet ? " has-bullet" : "");
        pEl.style.textAlign = p.align || "left";
        if (p.level) pEl.style.paddingLeft = `${p.level * 24 + (p.bullet ? 18 : 0)}px`;
        else if (p.bullet) pEl.style.paddingLeft = "18px";

        const runs = p.runs || [];
        if (!runs.length || runs.every((r) => !r.text)) {
          const first = runs[0] || {};
          pEl.style.fontSize = `${first.fontSize || 18}px`;
          pEl.style.color = first.color || "#181B22";
          pEl.appendChild(document.createElement("br"));
        } else {
          for (const r of runs) {
            if (r.text === "\n") {
              pEl.appendChild(document.createElement("br"));
              continue;
            }
            const span = document.createElement("span");
            span.textContent = r.text;
            span.style.fontSize = `${r.fontSize || 18}px`;
            if (r.fontFamily) span.style.fontFamily = `"${r.fontFamily}", system-ui, sans-serif`;
            if (r.bold) span.style.fontWeight = "700";
            if (r.italic) span.style.fontStyle = "italic";
            if (r.underline) span.style.textDecoration = "underline";
            span.style.color = r.color || "#181B22";
            pEl.appendChild(span);
          }
        }
        textWrap.appendChild(pEl);
      }

      if (isEditingThis) {
        textWrap.addEventListener("input", () => {
          syncEditedTextFromDom(el, textWrap);
          el.textChanged = true;
          slide.dirty = true;
          renderThumbAt(activeSlideIdx);
          scheduleSave();
        });
      }

      box.appendChild(textWrap);
    }

    if (interactive && !isBackdrop) {
      if (selectedUid === el.uid) {
        box.classList.add("is-selected");
        if (editableNow() && !el.inGroup && editingUid !== el.uid) {
          for (const dir of ["nw", "n", "ne", "e", "se", "s", "sw", "w"]) {
            const h = document.createElement("span");
            h.className = `sld-handle is-${dir}`;
            h.dataset.handle = dir;
            box.appendChild(h);
          }
        }
      }
      box.addEventListener("pointerdown", (ev) => onElementPointerDown(ev, el, slide));
      if (el.kind === "shape") {
        box.addEventListener("dblclick", (ev) => {
          ev.stopPropagation();
          if (!editableNow()) return;
          pushUndo();
          selectedUid = el.uid;
          editingUid = el.uid;
          editingCell = null;
          renderMainStage();
          const editorDiv = slideCanvasEl.querySelector(`.sld-el[data-uid="${el.uid}"] .sld-text`);
          editorDiv?.focus();
        });
      }
    }

    container.appendChild(box);
  };

  for (const b of slide.backdropElements || []) drawElement(b, true);
  for (const e of slide.elements || []) drawElement(e, false);
}

function syncEditedTextFromDom(el, textWrap) {
  const baseRun = el.paragraphs?.[0]?.runs?.[0] || {
    fontSize: 18,
    fontFamily: "",
    bold: false,
    italic: false,
    underline: false,
    color: "#181B22",
  };
  const baseAlign = el.paragraphs?.[0]?.align || "left";
  const baseBullet = !!el.paragraphs?.[0]?.bullet;
  const rawLines = String(textWrap.innerText || "").replace(/\r\n/g, "\n").split("\n");
  el.paragraphs = rawLines.map((line, i) => {
    const prevP = el.paragraphs?.[i];
    const prevRun = prevP?.runs?.[0] || baseRun;
    return {
      align: prevP?.align || baseAlign,
      bullet: prevP ? !!prevP.bullet : baseBullet,
      level: prevP?.level || 0,
      runs: [{ ...prevRun, text: line }],
    };
  });
}

function fitStage() {
  if (!stageWrapEl || !stageBoxEl || !slideCanvasEl || !session) return;
  const rect = stageWrapEl.getBoundingClientRect();
  const availW = Math.max(120, rect.width - (window.innerWidth <= 560 ? 16 : 40));
  const availH = Math.max(90, rect.height - (window.innerWidth <= 560 ? 16 : 40));
  currentScale = Math.min(availW / session.slideW, availH / session.slideH, 1.5);
  const scaledW = Math.round(session.slideW * currentScale);
  const scaledH = Math.round(session.slideH * currentScale);
  stageBoxEl.style.width = `${scaledW}px`;
  stageBoxEl.style.height = `${scaledH}px`;
  slideCanvasEl.style.transform = `scale(${currentScale})`;
}

function renderThumbAt(idx) {
  if (!railEl || !session) return;
  const item = railEl.querySelector(`.rail-item[data-idx="${idx}"] .rail-canvas`);
  if (!item || !slides[idx]) return;
  renderSlideLayer(item, slides[idx], false);
  const thumbScale = 144 / session.slideW;
  item.style.transform = `scale(${thumbScale})`;
}

function renderRail() {
  if (!railEl || !session) return;
  railEl.replaceChildren();
  const thumbW = 144;
  const thumbH = Math.round((thumbW * session.slideH) / session.slideW);
  const thumbScale = thumbW / session.slideW;

  slides.forEach((slide, idx) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "rail-item" + (idx === activeSlideIdx ? " is-active" : "");
    btn.dataset.idx = String(idx);
    btn.setAttribute("aria-label", S().slideOf(idx + 1, slides.length));

    const num = document.createElement("span");
    num.className = "rail-num";
    num.textContent = String(idx + 1);

    const frame = document.createElement("span");
    frame.className = "rail-frame";
    frame.style.width = `${thumbW}px`;
    frame.style.height = `${thumbH}px`;

    const canvas = document.createElement("div");
    canvas.className = "rail-canvas";
    renderSlideLayer(canvas, slide, false);
    canvas.style.transform = `scale(${thumbScale})`;

    frame.appendChild(canvas);
    btn.appendChild(num);
    btn.appendChild(frame);
    btn.addEventListener("click", () => selectSlide(idx));
    railEl.appendChild(btn);
  });
}

function formatSlideCounter(cur, total) {
  return (typeof window !== "undefined" && window.innerWidth <= 1020)
    ? S().slideOfCompact(cur, total)
    : S().slideOf(cur, total);
}

function renderMainStage() {
  const s = S();
  const stepperEl = document.querySelector(".slide-stepper");
  const firstBarSep = barEl?.querySelector(".bar-sep");
  if (!slides.length) {
    if (railEl) railEl.hidden = true;
    if (stepperEl) stepperEl.hidden = true;
    if (firstBarSep) firstBarSep.hidden = true;
    if (toolsEl) toolsEl.hidden = true;
    if (notesPanelEl) notesPanelEl.hidden = true;
    if (stageBoxEl) stageBoxEl.hidden = true;
    if (emptyDeckEl) {
      emptyDeckEl.hidden = false;
      if (emptyTitleEl) emptyTitleEl.textContent = s.deckEmptyTitle;
      if (emptyBodyEl) emptyBodyEl.textContent = s.deckEmptyBody;
      if (emptyAddBtn) {
        emptyAddBtn.textContent = s.addSlide;
        emptyAddBtn.hidden = !editableNow();
      }
    }
    if (notesInputEl) notesInputEl.value = "";
    return;
  }

  if (railEl) railEl.hidden = false;
  if (stepperEl) stepperEl.hidden = false;
  if (firstBarSep) firstBarSep.hidden = false;
  if (toolsEl) toolsEl.hidden = false;
  if (emptyDeckEl) emptyDeckEl.hidden = true;
  if (stageBoxEl) stageBoxEl.hidden = false;
  if (activeSlideIdx >= slides.length) activeSlideIdx = slides.length - 1;
  if (activeSlideIdx < 0) activeSlideIdx = 0;

  const slide = getActiveSlide();
  renderSlideLayer(slideCanvasEl, slide, true);
  fitStage();

  if (slideCounterEl) slideCounterEl.textContent = formatSlideCounter(activeSlideIdx + 1, slides.length);
  if (prevSlideBtn) prevSlideBtn.disabled = activeSlideIdx <= 0;
  if (nextSlideBtn) nextSlideBtn.disabled = activeSlideIdx >= slides.length - 1;
  const notesBtn = toolsEl?.querySelector('[data-cmd="notes"]');
  if (notesBtn) notesBtn.hidden = isPhoneViewport() && !slide?.notes;
  if (notesInputEl) {
    notesInputEl.value = slide?.notes || "";
    notesInputEl.readOnly = !editableNow();
  }
  for (const el of railEl?.querySelectorAll(".rail-item") || []) {
    el.classList.toggle("is-active", Number(el.dataset.idx) === activeSlideIdx);
  }
  updateToolbarControls();
}

function selectSlide(idx) {
  if (idx < 0 || idx >= slides.length) return;
  activeSlideIdx = idx;
  selectedUid = null;
  editingUid = null;
  editingCell = null;
  renderMainStage();
  const activeThumb = railEl?.querySelector(`.rail-item[data-idx="${idx}"]`);
  activeThumb?.scrollIntoView({ block: "nearest" });
}

function updateToolbarControls() {
  const canEdit = editableNow();
  for (const el of toolsEl?.querySelectorAll("[data-edit-tool]") || []) {
    el.hidden = !canEdit;
  }
  const sel = getSelectedElement();
  const isShape = sel && sel.kind === "shape";
  const firstRun = isShape ? sel.paragraphs?.[0]?.runs?.[0] : null;

  for (const cmd of ["bold", "italic", "underline"]) {
    const btn = toolsEl?.querySelector(`[data-cmd="${cmd}"]`);
    if (btn) {
      btn.disabled = !isShape;
      btn.setAttribute("aria-pressed", isShape && firstRun?.[cmd] ? "true" : "false");
    }
  }
  const delBtn = toolsEl?.querySelector('[data-cmd="deleteShape"]');
  if (delBtn) delBtn.disabled = !sel || !!sel.inGroup;

  if (fontSizeEl && firstRun?.fontSize) {
    const rounded = String(firstRun.fontSize);
    if ([...fontSizeEl.options].some((o) => o.value === rounded)) {
      fontSizeEl.value = rounded;
    }
  }
  if (textColorEl && firstRun?.color && /^#[0-9A-Fa-f]{6}$/.test(firstRun.color)) {
    textColorEl.value = firstRun.color;
  }
  if (fillColorEl && isShape && sel.fillHex && /^#[0-9A-Fa-f]{6}$/.test(sel.fillHex)) {
    fillColorEl.value = sel.fillHex;
  }
}

function onElementPointerDown(ev, el, slide) {
  ev.stopPropagation();
  if (editingUid === el.uid) return;
  editingUid = null;
  editingCell = null;
  selectedUid = el.uid;
  renderMainStage();

  if (!editableNow() || el.inGroup) return;
  const handle = ev.target.closest("[data-handle]")?.dataset.handle || "";
  const startClientX = ev.clientX;
  const startClientY = ev.clientY;
  const orig = { x: el.x, y: el.y, w: el.w, h: el.h };
  let moved = false;

  function onMove(me) {
    const dx = (me.clientX - startClientX) / currentScale;
    const dy = (me.clientY - startClientY) / currentScale;
    if (!moved && Math.hypot(dx, dy) < 2) return;
    if (!moved) {
      moved = true;
      pushUndo();
    }
    if (!handle) {
      el.x = Math.round(orig.x + dx);
      el.y = Math.round(orig.y + dy);
    } else {
      let nx = orig.x, ny = orig.y, nw = orig.w, nh = orig.h;
      if (handle.includes("e")) nw = Math.max(24, orig.w + dx);
      if (handle.includes("s")) nh = Math.max(20, orig.h + dy);
      if (handle.includes("w")) {
        const d = Math.min(orig.w - 24, dx);
        nx = orig.x + d;
        nw = orig.w - d;
      }
      if (handle.includes("n")) {
        const d = Math.min(orig.h - 20, dy);
        ny = orig.y + d;
        nh = orig.h - d;
      }
      el.x = Math.round(nx);
      el.y = Math.round(ny);
      el.w = Math.round(nw);
      el.h = Math.round(nh);
    }
    const domBox = slideCanvasEl.querySelector(`.sld-el[data-uid="${el.uid}"]`);
    if (domBox) {
      domBox.style.left = `${el.x}px`;
      domBox.style.top = `${el.y}px`;
      domBox.style.width = `${el.w}px`;
      domBox.style.height = `${el.h}px`;
    }
  }

  function onUp() {
    window.removeEventListener("pointermove", onMove);
    window.removeEventListener("pointerup", onUp);
    if (moved) {
      slide.dirty = true;
      renderThumbAt(activeSlideIdx);
      scheduleSave();
    }
  }

  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
}

stageWrapEl?.addEventListener("pointerdown", () => {
  if (selectedUid || editingUid) {
    selectedUid = null;
    editingUid = null;
    editingCell = null;
    renderMainStage();
  }
});

function mutateSelectedRuns(mutator) {
  const slide = getActiveSlide();
  const sel = getSelectedElement();
  if (!slide || !sel || sel.kind !== "shape" || !editableNow()) return;
  pushUndo();
  if (!sel.paragraphs.length) {
    sel.paragraphs = [{ align: "left", bullet: false, level: 0, runs: [{ text: "", fontSize: 18, bold: false, italic: false, underline: false, color: "#181B22" }] }];
  }
  for (const p of sel.paragraphs) {
    if (!p.runs.length) p.runs.push({ text: "", fontSize: 18, bold: false, italic: false, underline: false, color: "#181B22" });
    for (const r of p.runs) mutator(r, p);
  }
  sel.textChanged = true;
  slide.dirty = true;
  onDeckMutated(true);
}

function addShapeToSlide(kind) {
  const slide = getActiveSlide();
  if (!slide || !editableNow() || !session) return;
  pushUndo();
  const uid = `new-${++uidCounter}`;
  const isText = kind === "textbox";
  const w = isText ? 480 : 260;
  const h = isText ? 84 : 160;
  const newEl = {
    uid,
    kind: "shape",
    isNew: true,
    name: isText ? "Text Box" : kind === "ellipse" ? "Oval" : "Rectangle",
    x: Math.round((session.slideW - w) / 2),
    y: Math.round((session.slideH - h) / 2),
    w,
    h,
    rot: 0,
    geom: kind === "ellipse" ? "ellipse" : "rect",
    fill: isText ? "transparent" : "#DBEAFE",
    fillHex: isText ? "" : "#DBEAFE",
    stroke: isText ? null : { color: "#2563EB", hex: "#2563EB", width: 2, dashed: false },
    vAlign: "middle",
    padding: { top: 8, right: 12, bottom: 8, left: 12 },
    paragraphs: [
      {
        align: isText ? "left" : "center",
        bullet: false,
        level: 0,
        runs: [{ text: isText ? S().addTextBox : "", fontSize: isText ? 24 : 18, bold: false, italic: false, underline: false, color: "#181B22" }],
      },
    ],
  };
  slide.elements.push(newEl);
  slide.dirty = true;
  selectedUid = uid;
  editingUid = isText ? uid : null;
  onDeckMutated(true);
  if (isText) {
    const editorDiv = slideCanvasEl.querySelector(`.sld-el[data-uid="${uid}"] .sld-text`);
    editorDiv?.focus();
  }
}

function addBlankSlide() {
  if (!editableNow() || !session) return;
  pushUndo();
  const titleUid = `new-${++uidCounter}`;
  const newSlide = {
    id: `slide-new-${uidCounter}`,
    sourceSlidePath: "",
    layoutPath: session.defaultLayoutPath,
    notesPath: "",
    background: { kind: "solid", css: "#FFFFFF", hex: "#FFFFFF" },
    backdropElements: [],
    elements: [
      {
        uid: titleUid,
        kind: "shape",
        isNew: true,
        name: "Title",
        x: Math.round(session.slideW * 0.08),
        y: Math.round(session.slideH * 0.1),
        w: Math.round(session.slideW * 0.84),
        h: 96,
        rot: 0,
        geom: "rect",
        fill: "transparent",
        fillHex: "",
        stroke: null,
        vAlign: "middle",
        padding: { top: 8, right: 12, bottom: 8, left: 12 },
        paragraphs: [
          {
            align: "left",
            bullet: false,
            level: 0,
            runs: [{ text: S().addSlide, fontSize: 36, bold: true, italic: false, underline: false, color: "#181B22" }],
          },
        ],
      },
    ],
    notes: "",
    dirty: true,
    notesDirty: false,
  };
  const insertAt = slides.length ? activeSlideIdx + 1 : 0;
  slides.splice(insertAt, 0, newSlide);
  activeSlideIdx = insertAt;
  selectedUid = titleUid;
  editingUid = null;
  renderRail();
  renderMainStage();
  scheduleSave();
}

function duplicateActiveSlide() {
  const cur = getActiveSlide();
  if (!cur || !editableNow()) return;
  pushUndo();
  const copy = deepCloneSlides([cur])[0];
  copy.id = `slide-dup-${++uidCounter}`;
  copy.elements.forEach((e) => {
    e.uid = `dup-${++uidCounter}`;
  });
  if (copy.notes) copy.notesDirty = true;
  slides.splice(activeSlideIdx + 1, 0, copy);
  activeSlideIdx += 1;
  selectedUid = null;
  editingUid = null;
  renderRail();
  renderMainStage();
  scheduleSave();
}

function deleteActiveSlide() {
  if (!slides.length || !editableNow()) return;
  pushUndo();
  slides.splice(activeSlideIdx, 1);
  if (activeSlideIdx >= slides.length) activeSlideIdx = Math.max(0, slides.length - 1);
  selectedUid = null;
  editingUid = null;
  renderRail();
  renderMainStage();
  scheduleSave();
}

function moveActiveSlide(delta) {
  const target = activeSlideIdx + delta;
  if (!editableNow() || target < 0 || target >= slides.length) return;
  pushUndo();
  const [moved] = slides.splice(activeSlideIdx, 1);
  slides.splice(target, 0, moved);
  activeSlideIdx = target;
  renderRail();
  renderMainStage();
  scheduleSave();
}

function deleteSelectedElement() {
  const slide = getActiveSlide();
  const sel = getSelectedElement();
  if (!slide || !sel || sel.inGroup || !editableNow()) return;
  pushUndo();
  slide.elements = slide.elements.filter((e) => e.uid !== sel.uid);
  if (sel.xmlId && !sel.isNew) {
    slide.deletedXmlIds = slide.deletedXmlIds || [];
    slide.deletedXmlIds.push(String(sel.xmlId));
  }
  slide.dirty = true;
  selectedUid = null;
  editingUid = null;
  onDeckMutated(true);
}

function applyLanguage() {
  const s = S();
  document.documentElement.lang = lang;
  if (langEl) {
    langEl.value = lang;
    langEl.setAttribute("aria-label", s.language);
    langEl.title = s.language;
  }
  if (deckNameEl) deckNameEl.textContent = pageTitle || s.deckTitle;
  if (retryEl) retryEl.textContent = s.retry;
  if (conflictOpenEl) conflictOpenEl.textContent = s.open;
  if (conflictKeepEl) conflictKeepEl.textContent = s.keep;
  const live = Number(conflictEl?.dataset.live) || 0;
  if (conflictEl && !conflictEl.hidden && live && conflictTextEl) {
    conflictTextEl.textContent = s.conflict(live);
  }
  if (prevSlideBtn) {
    prevSlideBtn.title = s.prevSlide;
    prevSlideBtn.setAttribute("aria-label", s.prevSlide);
  }
  if (nextSlideBtn) {
    nextSlideBtn.title = s.nextSlide;
    nextSlideBtn.setAttribute("aria-label", s.nextSlide);
  }
  const tipMap = {
    undo: s.undo,
    redo: s.redo,
    addSlide: s.addSlide,
    duplicateSlide: s.duplicateSlide,
    deleteSlide: s.deleteSlide,
    moveUp: s.moveSlideUp,
    moveDown: s.moveSlideDown,
    addTextBox: s.addTextBox,
    addRect: s.addRect,
    addEllipse: s.addEllipse,
    deleteShape: s.deleteShape,
    bold: s.bold,
    italic: s.italic,
    underline: s.underline,
    alignLeft: s.alignLeft,
    alignCenter: s.alignCenter,
    alignRight: s.alignRight,
    notes: s.speakerNotes,
    find: s.find,
  };
  for (const [cmd, label] of Object.entries(tipMap)) {
    const btn = document.querySelector(`[data-cmd="${cmd}"]`);
    if (btn) {
      btn.title = label;
      btn.setAttribute("aria-label", label);
    }
  }
  if (fontSizeEl) {
    fontSizeEl.title = s.fontSize;
    fontSizeEl.setAttribute("aria-label", s.fontSize);
  }
  if (textColorEl) {
    textColorEl.title = s.textColor;
    textColorEl.setAttribute("aria-label", s.textColor);
  }
  if (fillColorEl) {
    fillColorEl.title = s.fillColor;
    fillColorEl.setAttribute("aria-label", s.fillColor);
  }
  if (notesLabelEl) notesLabelEl.textContent = s.speakerNotes;
  if (notesInputEl) notesInputEl.placeholder = s.notesPlaceholder;
  if (findInputEl) findInputEl.placeholder = s.findPlaceholder;
  const closeFindBtn = document.getElementById("find-close");
  if (closeFindBtn) {
    closeFindBtn.title = s.closeFind;
    closeFindBtn.setAttribute("aria-label", s.closeFind);
  }
  if (slideCounterEl && slides.length) {
    slideCounterEl.textContent = formatSlideCounter(activeSlideIdx + 1, slides.length);
  }
  if (!slides.length && emptyDeckEl && !emptyDeckEl.hidden) {
    if (emptyTitleEl) emptyTitleEl.textContent = s.deckEmptyTitle;
    if (emptyBodyEl) emptyBodyEl.textContent = s.deckEmptyBody;
    if (emptyAddBtn) emptyAddBtn.textContent = s.addSlide;
  }
  if (backHubBtn) backHubBtn.setAttribute("aria-label", s.backToHubAria);
  if (backHubLabelEl) backHubLabelEl.textContent = s.backToHub;
  if (standaloneTitleEl) standaloneTitleEl.setAttribute("aria-label", s.fileTitleAria);
  paintSave();
}

// Search across all slides
function runDeckSearch(jump = true) {
  const q = String(findInputEl?.value || "").trim().toLowerCase();
  findHits = [];
  if (q) {
    slides.forEach((sl, sIdx) => {
      for (const el of sl.elements) {
        let hay = "";
        if (el.kind === "shape") {
          hay = (el.paragraphs || []).map((p) => (p.runs || []).map((r) => r.text).join("")).join("\n");
        } else if (el.kind === "table") {
          hay = (el.rows || []).map((r) => (r.cells || []).map((c) => c.text).join(" ")).join("\n");
        }
        if (hay.toLowerCase().includes(q)) {
          findHits.push({ slideIdx: sIdx, uid: el.uid });
        }
      }
      if (sl.notes && sl.notes.toLowerCase().includes(q)) {
        findHits.push({ slideIdx: sIdx, uid: null, inNotes: true });
      }
    });
  }
  if (findIdx >= findHits.length) findIdx = 0;
  if (findIdx < 0) findIdx = Math.max(0, findHits.length - 1);
  if (findCountEl) {
    findCountEl.textContent = q ? S().matches(findHits.length ? findIdx + 1 : 0, findHits.length) : "";
  }
  if (jump && findHits[findIdx]) {
    const hit = findHits[findIdx];
    activeSlideIdx = hit.slideIdx;
    selectedUid = hit.uid;
    editingUid = null;
    if (hit.inNotes && notesPanelEl) notesPanelEl.hidden = false;
    renderMainStage();
  }
}

// Save coordinator
let saveFn = null;
let saveTimer = 0;

function scheduleSave(forceTitle = false) {
  if (!saveFn || (!editableNow() && !forceTitle)) return;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => void saveFn(), 1500);
}

function onDeckMutated(rerenderStage = true) {
  if (rerenderStage) renderMainStage();
  renderThumbAt(activeSlideIdx);
  scheduleSave();
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
  const draft = await openStandaloneDraft("pptx");
  let door = null;
  let info = null;
  let base64 = "";

  if (draft) {
    pageTitle = draft.title || S().untitledDeck;
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
      fail(S().deckUnreadableTitle, S().deckUnreadable);
      return;
    }
    if (!info || !info.entry) {
      fail(S().deckUnreadableTitle, S().deckUnreadable);
      return;
    }

    pageTitle = info.title || "";
    writableCap = Array.isArray(info.capabilities) && info.capabilities.includes("publish");

    let file;
    try {
      file = await door.read(info.entry);
    } catch {
      fail(S().deckUnreadableTitle, S().deckUnreadable);
      return;
    }
    base64 = (file && file.base64) || "";
    if (!base64) {
      fail(S().deckUnreadableTitle, S().deckUnreadable);
      return;
    }
  }

  try {
    session = parsePptx(base64);
    slides = deepCloneSlides(session.slides);
  } catch {
    fail(S().deckUnreadableTitle, S().deckUnreadable);
    return;
  }

  let baseVersion = (info && info.version) || 1;
  let lastSaved = "";
  let lastSavedTitle = pageTitle;
  try {
    lastSaved = savePptx(session, slides);
  } catch {
    lastSaved = base64;
  }
  let saving = false;
  let pending = false;

  saveFn = async function save() {
    clearTimeout(saveTimer);
    if (!editableNow() && !draft) return;
    if (saving) {
      pending = true;
      return;
    }
    let next;
    try {
      next = savePptx(session, slides);
    } catch {
      setSaveState("unsaved");
      return;
    }
    const nextTitle = draft && standaloneTitleEl ? standaloneTitleEl.value.trim() || S().untitledDeck : pageTitle;
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
        try {
          session = parsePptx(next);
        } catch { /* keep session */ }
        setSaveState(res.persisted ? "savedLocal" : "tabOnly");
      } else {
        const out = await door.publishFiles([{ path: info.entry, data: next }], baseVersion);
        baseVersion = (out && out.version && out.version.n) || baseVersion + 1;
        lastSaved = next;
        try {
          const refreshed = parsePptx(next);
          session = refreshed;
        } catch { /* keep session */ }
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
        saveTimer = setTimeout(() => void saveFn(), 200);
      }
    }
  };

  standaloneTitleEl?.addEventListener("input", () => scheduleSave(true));

  // Wire UI controls
  prevSlideBtn?.addEventListener("click", () => selectSlide(activeSlideIdx - 1));
  nextSlideBtn?.addEventListener("click", () => selectSlide(activeSlideIdx + 1));
  emptyAddBtn?.addEventListener("click", () => addBlankSlide());

  toolsEl?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-cmd]");
    if (!btn) return;
    const cmd = btn.dataset.cmd;
    if (cmd === "notes") {
      if (notesPanelEl) {
        notesPanelEl.hidden = !notesPanelEl.hidden;
        btn.setAttribute("aria-pressed", notesPanelEl.hidden ? "false" : "true");
        fitStage();
      }
      return;
    }
    if (cmd === "find") {
      if (findBarEl) {
        findBarEl.hidden = !findBarEl.hidden;
        fitStage();
        if (!findBarEl.hidden) {
          findInputEl?.focus();
          findInputEl?.select();
          runDeckSearch(true);
        }
      }
      return;
    }
    if (!editableNow()) return;
    if (cmd === "undo" && undoStack.length) {
      redoStack.push({ slides: deepCloneSlides(slides), activeSlideIdx, selectedUid });
      const prev = undoStack.pop();
      slides = prev.slides;
      activeSlideIdx = prev.activeSlideIdx;
      selectedUid = prev.selectedUid;
      editingUid = null;
      renderRail();
      renderMainStage();
      scheduleSave();
    } else if (cmd === "redo" && redoStack.length) {
      undoStack.push({ slides: deepCloneSlides(slides), activeSlideIdx, selectedUid });
      const next = redoStack.pop();
      slides = next.slides;
      activeSlideIdx = next.activeSlideIdx;
      selectedUid = next.selectedUid;
      editingUid = null;
      renderRail();
      renderMainStage();
      scheduleSave();
    } else if (cmd === "addSlide") addBlankSlide();
    else if (cmd === "duplicateSlide") duplicateActiveSlide();
    else if (cmd === "deleteSlide") deleteActiveSlide();
    else if (cmd === "moveUp") moveActiveSlide(-1);
    else if (cmd === "moveDown") moveActiveSlide(1);
    else if (cmd === "addTextBox") addShapeToSlide("textbox");
    else if (cmd === "addRect") addShapeToSlide("rect");
    else if (cmd === "addEllipse") addShapeToSlide("ellipse");
    else if (cmd === "deleteShape") deleteSelectedElement();
    else if (cmd === "bold") {
      const sel = getSelectedElement();
      const nextVal = !(sel?.paragraphs?.[0]?.runs?.[0]?.bold);
      mutateSelectedRuns((r) => { r.bold = nextVal; });
    } else if (cmd === "italic") {
      const sel = getSelectedElement();
      const nextVal = !(sel?.paragraphs?.[0]?.runs?.[0]?.italic);
      mutateSelectedRuns((r) => { r.italic = nextVal; });
    } else if (cmd === "underline") {
      const sel = getSelectedElement();
      const nextVal = !(sel?.paragraphs?.[0]?.runs?.[0]?.underline);
      mutateSelectedRuns((r) => { r.underline = nextVal; });
    } else if (cmd === "alignLeft") mutateSelectedRuns((_r, p) => { p.align = "left"; });
    else if (cmd === "alignCenter") mutateSelectedRuns((_r, p) => { p.align = "center"; });
    else if (cmd === "alignRight") mutateSelectedRuns((_r, p) => { p.align = "right"; });
  });

  fontSizeEl?.addEventListener("change", () => {
    const sz = Number(fontSizeEl.value);
    if (sz > 0) mutateSelectedRuns((r) => { r.fontSize = sz; });
  });

  textColorEl?.addEventListener("input", () => {
    const clr = textColorEl.value.toUpperCase();
    mutateSelectedRuns((r) => { r.color = clr; });
  });

  fillColorEl?.addEventListener("input", () => {
    const slide = getActiveSlide();
    const sel = getSelectedElement();
    const clr = fillColorEl.value.toUpperCase();
    if (!editableNow() || !slide) return;
    pushUndo();
    if (sel && sel.kind === "shape") {
      sel.fill = clr;
      sel.fillHex = clr;
      sel.fillChanged = true;
      slide.dirty = true;
      onDeckMutated(true);
    } else if (!sel) {
      slide.background = { kind: "solid", css: clr, hex: clr };
      slide.backgroundChanged = true;
      slide.dirty = true;
      onDeckMutated(true);
    }
  });

  notesInputEl?.addEventListener("input", () => {
    const slide = getActiveSlide();
    if (!slide || !editableNow()) return;
    slide.notes = notesInputEl.value;
    slide.notesDirty = true;
    scheduleSave();
  });

  langEl?.addEventListener("change", () => {
    if (!["pt-BR", "en-US"].includes(langEl.value)) return;
    lang = langEl.value;
    applyLanguage();
  });

  // Find bar events
  findInputEl?.addEventListener("input", () => {
    findIdx = 0;
    runDeckSearch(true);
  });
  findInputEl?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      findIdx += e.shiftKey ? -1 : 1;
      runDeckSearch(true);
    } else if (e.key === "Escape") {
      e.preventDefault();
      if (findBarEl) { findBarEl.hidden = true; fitStage(); }
    }
  });
  document.getElementById("find-prev")?.addEventListener("click", () => {
    findIdx -= 1;
    runDeckSearch(true);
  });
  document.getElementById("find-next")?.addEventListener("click", () => {
    findIdx += 1;
    runDeckSearch(true);
  });
  document.getElementById("find-close")?.addEventListener("click", () => {
    if (findBarEl) { findBarEl.hidden = true; fitStage(); }
  });

  // Keyboard shortcuts
  window.addEventListener("keydown", (e) => {
    const inInput = e.target.closest("input, textarea, select, [contenteditable='true']");
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
      e.preventDefault();
      if (findBarEl) {
        findBarEl.hidden = false;
        fitStage();
        findInputEl?.focus();
        findInputEl?.select();
        runDeckSearch(true);
      }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      if (editableNow() && saveFn) void saveFn();
      return;
    }
    if (inInput) {
      if (e.key === "Escape") {
        editingUid = null;
        editingCell = null;
        renderMainStage();
      }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z" && editableNow()) {
      e.preventDefault();
      const btn = toolsEl?.querySelector(e.shiftKey ? '[data-cmd="redo"]' : '[data-cmd="undo"]');
      btn?.click();
      return;
    }
    if ((e.key === "Delete" || e.key === "Backspace") && selectedUid && editableNow()) {
      e.preventDefault();
      deleteSelectedElement();
      return;
    }
    if (e.key === "Enter" && selectedUid && editableNow()) {
      const sel = getSelectedElement();
      if (sel && sel.kind === "shape") {
        e.preventDefault();
        pushUndo();
        editingUid = sel.uid;
        renderMainStage();
        const editorDiv = slideCanvasEl.querySelector(`.sld-el[data-uid="${sel.uid}"] .sld-text`);
        editorDiv?.focus();
      }
      return;
    }
    if (e.key === "PageDown" || (e.key === "ArrowDown" && !selectedUid) || (e.key === "ArrowRight" && !selectedUid)) {
      e.preventDefault();
      selectSlide(activeSlideIdx + 1);
    } else if (e.key === "PageUp" || (e.key === "ArrowUp" && !selectedUid) || (e.key === "ArrowLeft" && !selectedUid)) {
      e.preventDefault();
      selectSlide(activeSlideIdx - 1);
    }
  });

  if (typeof ResizeObserver !== "undefined" && stageWrapEl) {
    new ResizeObserver(() => fitStage()).observe(stageWrapEl);
  }
  window.addEventListener("resize", () => {
    const canEdit = editableNow();
    const savedKind = draft ? (draft.persisted ? "savedLocal" : "tabOnly") : "saved";
    if (!canEdit && saveState.kind !== "conflict") setSaveState("readonly");
    else if (canEdit && saveState.kind === "readonly") setSaveState(savedKind, baseVersion);
    renderMainStage();
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
      if (saveFn) await saveFn();
    }
  });

  if (barEl) barEl.hidden = false;
  if (workspaceEl) workspaceEl.hidden = false;
  const initialSavedKind = draft ? (draft.persisted ? "savedLocal" : "tabOnly") : "saved";
  setSaveState(editableNow() ? initialSavedKind : "readonly", baseVersion);
  renderRail();
  applyLanguage();
  renderMainStage();
  document.body.dataset.ready = "1";
}

main();
