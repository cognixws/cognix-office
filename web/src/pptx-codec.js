import { strToU8 } from "fflate";
import {
  attrOf,
  childrenByLocal,
  cssColor,
  descendantsByLocal,
  ensureContentTypeOverride,
  escapeXml,
  firstChildByLocal,
  firstDescendantByLocal,
  localOf,
  mediaToDataUrl,
  parseRels,
  parseTheme,
  parseXml,
  readTextPart,
  relsPathFor,
  resolveDrawingColor,
  serializeXml,
  unzipPackage,
  writeTextPart,
  zipPackage,
} from "./ooxml.js";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

export const EMU_PER_PX = 9525; // 914400 EMU/inch @ 96 px/inch
const EMU_PER_PT = 12700;

const emuToPx = (v) => Math.round((Number(v) || 0) / EMU_PER_PX);
const pxToEmu = (v) => Math.round((Number(v) || 0) * EMU_PER_PX);

function parseClrMap(el) {
  if (!el) return null;
  const slots = ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const out = {};
  let any = false;
  for (const s of slots) {
    const v = attrOf(el, s);
    if (v) {
      out[s] = v;
      any = true;
    }
  }
  return any ? out : null;
}

function parseXfrm(xfrmEl) {
  if (!xfrmEl) return null;
  const off = firstChildByLocal(xfrmEl, "off");
  const ext = firstChildByLocal(xfrmEl, "ext");
  if (!off || !ext) return null;
  const rotRaw = Number(attrOf(xfrmEl, "rot")) || 0;
  return {
    x: emuToPx(attrOf(off, "x")),
    y: emuToPx(attrOf(off, "y")),
    w: Math.max(4, emuToPx(attrOf(ext, "cx"))),
    h: Math.max(4, emuToPx(attrOf(ext, "cy"))),
    rot: Math.round((rotRaw / 60000) * 100) / 100,
    flipH: attrOf(xfrmEl, "flipH") === "1" || attrOf(xfrmEl, "flipH") === "true",
    flipV: attrOf(xfrmEl, "flipV") === "1" || attrOf(xfrmEl, "flipV") === "true",
  };
}

function parseFillFromContainer(containerEl, ctx, phColor = null) {
  if (!containerEl) return null;
  if (firstChildByLocal(containerEl, "noFill")) {
    return { kind: "none", css: "transparent", hex: "" };
  }
  const solid = firstChildByLocal(containerEl, "solidFill");
  if (solid) {
    const resolved = resolveDrawingColor(solid, ctx.theme, ctx.clrMap, phColor);
    if (resolved) {
      return { kind: "solid", css: cssColor(resolved), hex: resolved.hex };
    }
  }
  const grad = firstChildByLocal(containerEl, "gradFill");
  if (grad) {
    const gsLst = firstChildByLocal(grad, "gsLst");
    const stops = [];
    for (const gs of childrenByLocal(gsLst, "gs")) {
      const pos = Math.round((Number(attrOf(gs, "pos")) || 0) / 1000);
      const clr = resolveDrawingColor(gs, ctx.theme, ctx.clrMap, phColor);
      if (clr) stops.push(`${cssColor(clr)} ${pos}%`);
    }
    if (stops.length >= 2) {
      const lin = firstChildByLocal(grad, "lin");
      const ooxmlDeg = lin ? (Number(attrOf(lin, "ang")) || 0) / 60000 : 90;
      const cssDeg = Math.round((ooxmlDeg + 90) % 360);
      return { kind: "gradient", css: `linear-gradient(${cssDeg}deg, ${stops.join(", ")})`, hex: "" };
    }
  }
  const blipFill = firstChildByLocal(containerEl, "blipFill");
  if (blipFill && ctx.rels) {
    const blip = firstChildByLocal(blipFill, "blip");
    const rId = blip && attrOf(blip, "embed");
    const rel = rId ? ctx.rels.get(rId) : null;
    if (rel && !rel.external) {
      const url = mediaToDataUrl(ctx.files, rel.target);
      if (url) return { kind: "image", css: "", hex: "", imageUrl: url };
    }
  }
  return null;
}

function parseBg(rootDoc, ctx) {
  const bg = firstDescendantByLocal(rootDoc, "bg");
  if (!bg) return null;
  const bgPr = firstChildByLocal(bg, "bgPr");
  if (bgPr) {
    const fill = parseFillFromContainer(bgPr, ctx);
    if (fill) return fill;
  }
  const bgRef = firstChildByLocal(bg, "bgRef");
  if (bgRef) {
    const clr = resolveDrawingColor(bgRef, ctx.theme, ctx.clrMap);
    if (clr) return { kind: "solid", css: cssColor(clr), hex: clr.hex };
  }
  return null;
}

function parseStroke(spPr, styleEl, ctx) {
  const ln = spPr && firstChildByLocal(spPr, "ln");
  if (ln) {
    if (firstChildByLocal(ln, "noFill")) return null;
    const wEmu = Number(attrOf(ln, "w")) || EMU_PER_PT;
    const width = Math.max(1, Math.round((wEmu / EMU_PER_PX) * 10) / 10);
    const solid = firstChildByLocal(ln, "solidFill");
    const clr = solid ? resolveDrawingColor(solid, ctx.theme, ctx.clrMap) : null;
    const prstDash = firstChildByLocal(ln, "prstDash");
    const dashVal = prstDash ? attrOf(prstDash, "val") : "solid";
    if (clr) {
      return { color: cssColor(clr), hex: clr.hex, width, dashed: dashVal !== "solid" };
    }
  }
  const lnRef = styleEl && firstChildByLocal(styleEl, "lnRef");
  if (lnRef && Number(attrOf(lnRef, "idx")) > 0) {
    const clr = resolveDrawingColor(lnRef, ctx.theme, ctx.clrMap);
    if (clr) return { color: cssColor(clr), hex: clr.hex, width: 1, dashed: false };
  }
  return null;
}

function resolveTypeface(raw, theme) {
  if (!raw) return "";
  if (raw === "+mj-lt" || raw === "+mj-ea" || raw === "+mj-cs") return (theme && theme.majorFont) || "system-ui";
  if (raw === "+mn-lt" || raw === "+mn-ea" || raw === "+mn-cs") return (theme && theme.minorFont) || "system-ui";
  return raw;
}

function defaultPlaceholderDefaults(phType) {
  if (phType === "ctrTitle" || phType === "title") {
    return { fontSize: 36, bold: true, align: phType === "ctrTitle" ? "center" : "left", color: "" };
  }
  if (phType === "subTitle") {
    return { fontSize: 20, bold: false, align: "center", color: "#555E6D" };
  }
  if (phType === "sldNum" || phType === "dt" || phType === "ftr") {
    return { fontSize: 11, bold: false, align: "left", color: "#667085" };
  }
  return { fontSize: 18, bold: false, align: "left", color: "" };
}

function parseTxBody(txBody, ctx, defaults = {}) {
  if (!txBody) return { paragraphs: [], vAlign: "top", padding: { top: 6, right: 10, bottom: 6, left: 10 } };
  const bodyPr = firstChildByLocal(txBody, "bodyPr");
  let vAlign = defaults.vAlign || "top";
  const padding = { top: 6, right: 10, bottom: 6, left: 10 };
  if (bodyPr) {
    const anchor = attrOf(bodyPr, "anchor");
    if (anchor === "ctr") vAlign = "middle";
    else if (anchor === "b") vAlign = "bottom";
    else if (anchor === "t") vAlign = "top";
    const ins = (attr, fallback) => {
      const raw = attrOf(bodyPr, attr);
      return raw !== "" && Number.isFinite(Number(raw)) ? Math.max(0, emuToPx(raw)) : fallback;
    };
    padding.left = ins("lIns", 10);
    padding.top = ins("tIns", 6);
    padding.right = ins("rIns", 10);
    padding.bottom = ins("bIns", 6);
  }

  // Shape-level lstStyle level defaults
  const lstStyle = firstChildByLocal(txBody, "lstStyle");
  const lvlDefaults = new Map();
  if (lstStyle) {
    for (let i = 1; i <= 9; i++) {
      const lp = firstChildByLocal(lstStyle, `lvl${i}pPr`);
      if (!lp) continue;
      const defRPr = firstChildByLocal(lp, "defRPr");
      const sz = defRPr && Number(attrOf(defRPr, "sz"));
      const solid = defRPr && firstChildByLocal(defRPr, "solidFill");
      const clr = solid ? resolveDrawingColor(solid, ctx.theme, ctx.clrMap) : null;
      lvlDefaults.set(i - 1, {
        fontSize: sz ? Math.round(sz / 100) : null,
        color: clr ? clr.hex : "",
        align: mapPptAlign(attrOf(lp, "algn")),
      });
    }
  }

  const paragraphs = [];
  for (const p of childrenByLocal(txBody, "p")) {
    const pPr = firstChildByLocal(p, "pPr");
    const lvl = pPr ? Math.max(0, Math.min(5, Number(attrOf(pPr, "lvl")) || 0)) : 0;
    const ld = lvlDefaults.get(lvl) || {};
    const align = (pPr && mapPptAlign(attrOf(pPr, "algn"))) || ld.align || defaults.align || "left";
    let bullet = false;
    if (pPr) {
      if (firstChildByLocal(pPr, "buNone")) bullet = false;
      else if (firstChildByLocal(pPr, "buChar") || firstChildByLocal(pPr, "buAutoNum")) bullet = true;
    }
    const defRPr = pPr && firstChildByLocal(pPr, "defRPr");
    const pDefSz = defRPr && Number(attrOf(defRPr, "sz")) ? Math.round(Number(attrOf(defRPr, "sz")) / 100) : null;

    const runs = [];
    for (const ch of childrenByLocal(p)) {
      const tag = localOf(ch);
      if (tag === "r" || tag === "fld") {
        const rPr = firstChildByLocal(ch, "rPr");
        const tEl = firstChildByLocal(ch, "t");
        const text = tEl ? tEl.textContent || "" : "";
        const szRaw = rPr && Number(attrOf(rPr, "sz"));
        const fontSize = szRaw ? Math.round(szRaw / 100) : pDefSz || ld.fontSize || defaults.fontSize || 18;
        const bAttr = rPr ? attrOf(rPr, "b") : "";
        const bold = bAttr ? bAttr === "1" || bAttr === "true" : !!defaults.bold;
        const iAttr = rPr ? attrOf(rPr, "i") : "";
        const italic = iAttr === "1" || iAttr === "true";
        const uAttr = rPr ? attrOf(rPr, "u") : "";
        const underline = !!uAttr && uAttr !== "none";
        const solid = rPr && firstChildByLocal(rPr, "solidFill");
        const clr = solid ? resolveDrawingColor(solid, ctx.theme, ctx.clrMap) : null;
        const latin = rPr && firstChildByLocal(rPr, "latin");
        const fontFamily = resolveTypeface(latin && attrOf(latin, "typeface"), ctx.theme);
        runs.push({
          text,
          fontSize,
          fontFamily,
          bold,
          italic,
          underline,
          color: (clr && clr.hex) || ld.color || defaults.color || defaults.fontRefColor || "#181B22",
        });
      } else if (tag === "br") {
        runs.push({
          text: "\n",
          fontSize: pDefSz || ld.fontSize || defaults.fontSize || 18,
          fontFamily: "",
          bold: !!defaults.bold,
          italic: false,
          underline: false,
          color: ld.color || defaults.color || "#181B22",
        });
      }
    }
    paragraphs.push({ align, bullet, level: lvl, runs });
  }
  return { paragraphs, vAlign, padding };
}

function mapPptAlign(algn) {
  if (algn === "ctr") return "center";
  if (algn === "r") return "right";
  if (algn === "just") return "justify";
  if (algn === "l") return "left";
  return "";
}

function unmapPptAlign(align) {
  if (align === "center") return "ctr";
  if (align === "right") return "r";
  if (align === "justify") return "just";
  return "l";
}

function collectPlaceholders(spTree, ctx) {
  const byKey = new Map();
  if (!spTree) return byKey;
  for (const sp of childrenByLocal(spTree, "sp")) {
    const nvSpPr = firstChildByLocal(sp, "nvSpPr");
    const nvPr = nvSpPr && firstChildByLocal(nvSpPr, "nvPr");
    const ph = nvPr && firstChildByLocal(nvPr, "ph");
    if (!ph) continue;
    const type = attrOf(ph, "type") || "body";
    const idx = attrOf(ph, "idx") || "";
    const spPr = firstChildByLocal(sp, "spPr");
    const xfrm = parseXfrm(spPr && firstChildByLocal(spPr, "xfrm"));
    const txBody = firstChildByLocal(sp, "txBody");
    let defFontSize = null;
    let defColor = "";
    let defAlign = "";
    if (txBody) {
      const lstStyle = firstChildByLocal(txBody, "lstStyle");
      const lvl1 = lstStyle && firstChildByLocal(lstStyle, "lvl1pPr");
      if (lvl1) {
        defAlign = mapPptAlign(attrOf(lvl1, "algn"));
        const defRPr = firstChildByLocal(lvl1, "defRPr");
        if (defRPr && Number(attrOf(defRPr, "sz"))) defFontSize = Math.round(Number(attrOf(defRPr, "sz")) / 100);
        const solid = defRPr && firstChildByLocal(defRPr, "solidFill");
        const clr = solid ? resolveDrawingColor(solid, ctx.theme, ctx.clrMap) : null;
        if (clr) defColor = clr.hex;
      }
    }
    const entry = { type, idx, xfrm, fontSize: defFontSize, color: defColor, align: defAlign };
    if (idx) byKey.set(`idx:${idx}`, entry);
    if (!byKey.has(`type:${type}`)) byKey.set(`type:${type}`, entry);
  }
  return byKey;
}

function lookupPlaceholder(phEl, layoutPhs, masterPhs) {
  if (!phEl) return null;
  const type = attrOf(phEl, "type") || (attrOf(phEl, "idx") ? "body" : "");
  const idx = attrOf(phEl, "idx") || "";
  const l = (idx && layoutPhs.get(`idx:${idx}`)) || (type && layoutPhs.get(`type:${type}`)) || null;
  const m = (idx && masterPhs.get(`idx:${idx}`)) || (type && masterPhs.get(`type:${type}`)) || null;
  return {
    type: type || l?.type || m?.type || "",
    xfrm: l?.xfrm || m?.xfrm || null,
    fontSize: l?.fontSize || m?.fontSize || null,
    color: l?.color || m?.color || "",
    align: l?.align || m?.align || "",
  };
}

function mapGroupCoord(xfrm, groupMap) {
  if (!xfrm || !groupMap) return xfrm;
  const sx = groupMap.extW / Math.max(1, groupMap.chExtW);
  const sy = groupMap.extH / Math.max(1, groupMap.chExtH);
  return {
    ...xfrm,
    x: Math.round(groupMap.offX + (xfrm.x - groupMap.chOffX) * sx),
    y: Math.round(groupMap.offY + (xfrm.y - groupMap.chOffY) * sy),
    w: Math.max(4, Math.round(xfrm.w * sx)),
    h: Math.max(4, Math.round(xfrm.h * sy)),
  };
}

function parseSpTreeElements(spTree, ctx, opts = {}) {
  const out = [];
  if (!spTree) return out;
  let seq = opts.seqStart || 1;

  for (const node of childrenByLocal(spTree)) {
    const tag = localOf(node);
    if (tag === "sp") {
      const nvSpPr = firstChildByLocal(node, "nvSpPr");
      const cNvPr = nvSpPr && firstChildByLocal(nvSpPr, "cNvPr");
      const nvPr = nvSpPr && firstChildByLocal(nvSpPr, "nvPr");
      const ph = nvPr && firstChildByLocal(nvPr, "ph");
      if (opts.skipPlaceholders && ph) continue;
      if (cNvPr && attrOf(cNvPr, "hidden") === "1") continue;

      const xmlId = (cNvPr && attrOf(cNvPr, "id")) || String(seq);
      const name = (cNvPr && attrOf(cNvPr, "name")) || `Shape ${seq}`;
      const phInherited = ph ? lookupPlaceholder(ph, ctx.layoutPhs, ctx.masterPhs) : null;
      const phType = phInherited?.type || "";

      // Skip empty date/footer/slide-number placeholders on slides so blank badges don't clutter
      const spPr = firstChildByLocal(node, "spPr");
      let xfrm = parseXfrm(spPr && firstChildByLocal(spPr, "xfrm")) || phInherited?.xfrm;
      if (!xfrm) {
        if (phType === "ctrTitle" || phType === "title") xfrm = { x: 80, y: 60, w: ctx.slideW - 160, h: 110, rot: 0 };
        else xfrm = { x: 80, y: 190, w: ctx.slideW - 160, h: 380, rot: 0 };
      }
      xfrm = mapGroupCoord(xfrm, opts.groupMap);

      const prstGeom = spPr && firstChildByLocal(spPr, "prstGeom");
      const geom = (prstGeom && attrOf(prstGeom, "prst")) || "rect";
      const styleEl = firstChildByLocal(node, "style");

      // Style reference fill & font color
      let fill = parseFillFromContainer(spPr, ctx);
      if (!fill && styleEl) {
        const fillRef = firstChildByLocal(styleEl, "fillRef");
        if (fillRef && Number(attrOf(fillRef, "idx")) > 0) {
          const refClr = resolveDrawingColor(fillRef, ctx.theme, ctx.clrMap);
          if (refClr) fill = { kind: "solid", css: cssColor(refClr), hex: refClr.hex };
        }
      }
      let fontRefColor = "";
      if (styleEl) {
        const fontRef = firstChildByLocal(styleEl, "fontRef");
        const fClr = fontRef ? resolveDrawingColor(fontRef, ctx.theme, ctx.clrMap) : null;
        if (fClr) fontRefColor = fClr.hex;
      }

      const stroke = parseStroke(spPr, styleEl, ctx);
      const phDefs = defaultPlaceholderDefaults(phType);
      const tx = parseTxBody(firstChildByLocal(node, "txBody"), ctx, {
        fontSize: phInherited?.fontSize || phDefs.fontSize,
        bold: phDefs.bold,
        align: phInherited?.align || phDefs.align,
        color: phInherited?.color || phDefs.color,
        fontRefColor,
      });

      const hasText = tx.paragraphs.some((p) => p.runs.some((r) => r.text.length > 0));
      if (!hasText && (!fill || fill.kind === "none") && !stroke && (phType === "dt" || phType === "ftr" || phType === "sldNum")) {
        continue;
      }

      out.push({
        uid: `${opts.prefix || "el"}-${seq++}`,
        kind: "shape",
        xmlId,
        name,
        phType,
        inGroup: !!opts.groupMap,
        x: xfrm.x,
        y: xfrm.y,
        w: xfrm.w,
        h: xfrm.h,
        rot: xfrm.rot || 0,
        geom,
        fill: fill ? fill.css : "transparent",
        fillHex: fill && fill.kind === "solid" ? fill.hex : "",
        stroke,
        vAlign: tx.vAlign,
        padding: tx.padding,
        paragraphs: tx.paragraphs,
      });
    } else if (tag === "pic") {
      const nvPicPr = firstChildByLocal(node, "nvPicPr");
      const cNvPr = nvPicPr && firstChildByLocal(nvPicPr, "cNvPr");
      const nvPr = nvPicPr && firstChildByLocal(nvPicPr, "nvPr");
      if (opts.skipPlaceholders && nvPr && firstChildByLocal(nvPr, "ph")) continue;
      if (cNvPr && attrOf(cNvPr, "hidden") === "1") continue;

      const xmlId = (cNvPr && attrOf(cNvPr, "id")) || String(seq);
      const name = (cNvPr && attrOf(cNvPr, "name")) || `Picture ${seq}`;
      const blipFill = firstChildByLocal(node, "blipFill");
      const blip = blipFill && firstChildByLocal(blipFill, "blip");
      const rId = blip && attrOf(blip, "embed");
      const rel = rId && ctx.rels ? ctx.rels.get(rId) : null;
      if (!rel || rel.external) continue;
      const src = mediaToDataUrl(ctx.files, rel.target);
      if (!src) continue;

      const spPr = firstChildByLocal(node, "spPr");
      let xfrm = parseXfrm(spPr && firstChildByLocal(spPr, "xfrm")) || { x: 0, y: 0, w: ctx.slideW, h: ctx.slideH, rot: 0 };
      xfrm = mapGroupCoord(xfrm, opts.groupMap);

      let crop = null;
      const srcRect = blipFill && firstChildByLocal(blipFill, "srcRect");
      if (srcRect) {
        const pct = (k) => Math.max(0, Math.min(95, (Number(attrOf(srcRect, k)) || 0) / 1000));
        const l = pct("l"), t = pct("t"), r = pct("r"), b = pct("b");
        if (l || t || r || b) crop = { l, t, r, b };
      }

      out.push({
        uid: `${opts.prefix || "el"}-${seq++}`,
        kind: "image",
        xmlId,
        name,
        inGroup: !!opts.groupMap,
        x: xfrm.x,
        y: xfrm.y,
        w: xfrm.w,
        h: xfrm.h,
        rot: xfrm.rot || 0,
        src,
        mediaPath: rel.target,
        crop,
      });
    } else if (tag === "graphicFrame") {
      const nvPr = firstChildByLocal(node, "nvGraphicFramePr");
      const cNvPr = nvPr && firstChildByLocal(nvPr, "cNvPr");
      const xmlId = (cNvPr && attrOf(cNvPr, "id")) || String(seq);
      const name = (cNvPr && attrOf(cNvPr, "name")) || `Frame ${seq}`;
      let xfrm = parseXfrm(firstChildByLocal(node, "xfrm"));
      if (!xfrm) continue;
      xfrm = mapGroupCoord(xfrm, opts.groupMap);

      const tbl = firstDescendantByLocal(node, "tbl");
      if (tbl) {
        const tblGrid = firstChildByLocal(tbl, "tblGrid");
        const rawCols = childrenByLocal(tblGrid, "gridCol").map((c) => Math.max(20, emuToPx(attrOf(c, "w"))));
        const rows = [];
        const trEls = childrenByLocal(tbl, "tr");
        trEls.forEach((tr, rIdx) => {
          const rowH = Math.max(22, emuToPx(attrOf(tr, "h")));
          const cells = [];
          for (const tc of childrenByLocal(tr, "tc")) {
            if (attrOf(tc, "vMerge") === "1" || attrOf(tc, "hMerge") === "1") continue;
            const colspan = Math.max(1, Number(attrOf(tc, "gridSpan")) || 1);
            const rowspan = Math.max(1, Number(attrOf(tc, "rowSpan")) || 1);
            const tcPr = firstChildByLocal(tc, "tcPr");
            const fillObj = parseFillFromContainer(tcPr, ctx);
            const tx = parseTxBody(firstChildByLocal(tc, "txBody"), ctx, {
              fontSize: 14,
              bold: rIdx === 0,
              color: "#181B22",
            });
            const text = tx.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
            const firstRun = tx.paragraphs[0]?.runs[0];
            cells.push({
              text,
              align: tx.paragraphs[0]?.align || "left",
              bold: firstRun ? !!firstRun.bold : rIdx === 0,
              fontSize: firstRun?.fontSize || 14,
              color: firstRun?.color || "#181B22",
              fill: fillObj?.css || (rIdx === 0 ? "rgba(59, 130, 246, 0.12)" : "transparent"),
              colspan,
              rowspan,
            });
          }
          if (cells.length) rows.push({ height: rowH, cells });
        });
        if (rows.length) {
          out.push({
            uid: `${opts.prefix || "el"}-${seq++}`,
            kind: "table",
            xmlId,
            name,
            inGroup: !!opts.groupMap,
            x: xfrm.x,
            y: xfrm.y,
            w: xfrm.w,
            h: xfrm.h,
            rot: 0,
            colWidths: rawCols,
            rows,
          });
        }
      }
    } else if (tag === "cxnSp") {
      const nvCxnPr = firstChildByLocal(node, "nvCxnSpPr");
      const cNvPr = nvCxnPr && firstChildByLocal(nvCxnPr, "cNvPr");
      const xmlId = (cNvPr && attrOf(cNvPr, "id")) || String(seq);
      const spPr = firstChildByLocal(node, "spPr");
      let xfrm = parseXfrm(spPr && firstChildByLocal(spPr, "xfrm"));
      if (!xfrm) continue;
      xfrm = mapGroupCoord(xfrm, opts.groupMap);
      const stroke = parseStroke(spPr, firstChildByLocal(node, "style"), ctx) || { color: "#667085", hex: "#667085", width: 1.5, dashed: false };
      out.push({
        uid: `${opts.prefix || "el"}-${seq++}`,
        kind: "connector",
        xmlId,
        inGroup: !!opts.groupMap,
        x: xfrm.x,
        y: xfrm.y,
        w: Math.max(2, xfrm.w),
        h: Math.max(2, xfrm.h),
        rot: xfrm.rot || 0,
        flipH: xfrm.flipH,
        flipV: xfrm.flipV,
        stroke,
      });
    } else if (tag === "grpSp") {
      const grpSpPr = firstChildByLocal(node, "grpSpPr");
      const xfrmEl = grpSpPr && firstChildByLocal(grpSpPr, "xfrm");
      const off = xfrmEl && firstChildByLocal(xfrmEl, "off");
      const ext = xfrmEl && firstChildByLocal(xfrmEl, "ext");
      const chOff = xfrmEl && firstChildByLocal(xfrmEl, "chOff");
      const chExt = xfrmEl && firstChildByLocal(xfrmEl, "chExt");
      let gMap = opts.groupMap || null;
      if (off && ext && chOff && chExt) {
        const localGroup = {
          offX: emuToPx(attrOf(off, "x")),
          offY: emuToPx(attrOf(off, "y")),
          extW: Math.max(1, emuToPx(attrOf(ext, "cx"))),
          extH: Math.max(1, emuToPx(attrOf(ext, "cy"))),
          chOffX: emuToPx(attrOf(chOff, "x")),
          chOffY: emuToPx(attrOf(chOff, "y")),
          chExtW: Math.max(1, emuToPx(attrOf(chExt, "cx"))),
          chExtH: Math.max(1, emuToPx(attrOf(chExt, "cy"))),
        };
        if (gMap) {
          const mapped = mapGroupCoord({ x: localGroup.offX, y: localGroup.offY, w: localGroup.extW, h: localGroup.extH }, gMap);
          gMap = { ...localGroup, offX: mapped.x, offY: mapped.y, extW: mapped.w, extH: mapped.h };
        } else {
          gMap = localGroup;
        }
      }
      const nested = parseSpTreeElements(node, ctx, { ...opts, groupMap: gMap, seqStart: seq });
      seq += nested.length + 1;
      out.push(...nested);
    }
  }
  return out;
}

function parseSlideNotes(files, slideRels) {
  for (const rel of slideRels.values()) {
    if (rel.type.endsWith("/notesSlide") && !rel.external) {
      const xml = readTextPart(files, rel.target);
      if (!xml) return { notes: "", notesPath: rel.target };
      try {
        const doc = parseXml(xml);
        const spTree = firstDescendantByLocal(doc, "spTree");
        for (const sp of childrenByLocal(spTree, "sp")) {
          const nvPr = firstDescendantByLocal(sp, "nvPr");
          const ph = nvPr && firstChildByLocal(nvPr, "ph");
          if (ph && attrOf(ph, "type") === "body") {
            const txBody = firstChildByLocal(sp, "txBody");
            const lines = [];
            for (const p of childrenByLocal(txBody, "p")) {
              let line = "";
              for (const ch of childrenByLocal(p)) {
                if (localOf(ch) === "r" || localOf(ch) === "fld") {
                  const t = firstChildByLocal(ch, "t");
                  if (t) line += t.textContent || "";
                } else if (localOf(ch) === "br") {
                  line += "\n";
                }
              }
              lines.push(line);
            }
            return { notes: lines.join("\n").trim(), notesPath: rel.target };
          }
        }
      } catch {
        // ignore malformed notes
      }
      return { notes: "", notesPath: rel.target };
    }
  }
  return { notes: "", notesPath: "" };
}

export function parsePptx(input) {
  const files = unzipPackage(input);
  const presXml = readTextPart(files, "ppt/presentation.xml");
  if (!presXml) throw new Error("missing ppt/presentation.xml");
  const presDoc = parseXml(presXml);
  const presRels = parseRels(files, "ppt/presentation.xml");

  // Slide size
  let slideW = 1280;
  let slideH = 720;
  const sldSz = firstDescendantByLocal(presDoc, "sldSz");
  if (sldSz) {
    const cx = Number(attrOf(sldSz, "cx"));
    const cy = Number(attrOf(sldSz, "cy"));
    if (cx > 914400 && cy > 914400) {
      slideW = Math.round(cx / EMU_PER_PX);
      slideH = Math.round(cy / EMU_PER_PX);
    }
  }

  // Theme
  let themePath = "ppt/theme/theme1.xml";
  for (const rel of presRels.values()) {
    if (rel.type.endsWith("/theme") && !rel.external) {
      themePath = rel.target;
      break;
    }
  }
  const theme = parseTheme(readTextPart(files, themePath));

  // Ordered slide paths
  const sldIdLst = firstDescendantByLocal(presDoc, "sldIdLst");
  const slideRefs = [];
  if (sldIdLst) {
    for (const sldId of childrenByLocal(sldIdLst, "sldId")) {
      const rId = attrOf(sldId, "id"); // Note: sldId has both id="256" and r:id="rId2"
      let relId = "";
      if (sldId.attributes) {
        for (let i = 0; i < sldId.attributes.length; i++) {
          const a = sldId.attributes[i];
          const nm = a.name || a.nodeName || "";
          if (nm.endsWith(":id") || nm === "r:id") relId = a.value;
        }
      }
      const rel = presRels.get(relId || rId);
      if (rel && !rel.external && files[rel.target]) {
        slideRefs.push(rel.target);
      }
    }
  } else {
    // Fallback if sldIdLst omitted: discover ppt/slides/slideN.xml sorted numerically
    const found = Object.keys(files)
      .filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k))
      .sort((a, b) => {
        const na = Number(/(\d+)/.exec(a)?.[1] || 0);
        const nb = Number(/(\d+)/.exec(b)?.[1] || 0);
        return na - nb;
      });
    slideRefs.push(...found);
  }

  let defaultLayoutPath = "ppt/slideLayouts/slideLayout1.xml";
  const slides = slideRefs.map((slidePath, idx) => {
    const xml = readTextPart(files, slidePath);
    const sDoc = parseXml(xml);
    const sRels = parseRels(files, slidePath);

    let layoutPath = "";
    for (const r of sRels.values()) {
      if (r.type.endsWith("/slideLayout") && !r.external) {
        layoutPath = r.target;
        defaultLayoutPath = r.target;
        break;
      }
    }

    let layoutDoc = null;
    let layoutRels = new Map();
    let masterPath = "";
    let masterDoc = null;
    let masterRels = new Map();
    let clrMap = {
      bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2",
      accent1: "accent1", accent2: "accent2", accent3: "accent3",
      accent4: "accent4", accent5: "accent5", accent6: "accent6",
      hlink: "hlink", folHlink: "folHlink",
    };

    if (layoutPath && files[layoutPath]) {
      try {
        layoutDoc = parseXml(readTextPart(files, layoutPath));
        layoutRels = parseRels(files, layoutPath);
        for (const r of layoutRels.values()) {
          if (r.type.endsWith("/slideMaster") && !r.external) {
            masterPath = r.target;
            break;
          }
        }
      } catch { /* ignore */ }
    }
    if (masterPath && files[masterPath]) {
      try {
        masterDoc = parseXml(readTextPart(files, masterPath));
        masterRels = parseRels(files, masterPath);
        const mClr = parseClrMap(firstDescendantByLocal(masterDoc, "clrMap"));
        if (mClr) clrMap = { ...clrMap, ...mClr };
      } catch { /* ignore */ }
    }
    // Check slide or layout clrMapOvr
    for (const d of [layoutDoc, sDoc]) {
      if (!d) continue;
      const ovr = firstDescendantByLocal(d, "clrMapOvr");
      const overrideMapping = ovr && firstChildByLocal(ovr, "overrideClrMapping");
      const parsed = parseClrMap(overrideMapping);
      if (parsed) clrMap = { ...clrMap, ...parsed };
    }

    const masterPhs = masterDoc ? collectPlaceholders(firstDescendantByLocal(masterDoc, "spTree"), { theme, clrMap }) : new Map();
    const layoutPhs = layoutDoc ? collectPlaceholders(firstDescendantByLocal(layoutDoc, "spTree"), { theme, clrMap }) : new Map();

    // Background cascade: slide -> layout -> master -> #FFFFFF
    const bg =
      parseBg(sDoc, { files, rels: sRels, theme, clrMap }) ||
      (layoutDoc && parseBg(layoutDoc, { files, rels: layoutRels, theme, clrMap })) ||
      (masterDoc && parseBg(masterDoc, { files, rels: masterRels, theme, clrMap })) ||
      { kind: "solid", css: "#FFFFFF", hex: "#FFFFFF" };

    // Non-placeholder decorative elements from master & layout (behind slide elements)
    const showMasterOnSlide = attrOf(sDoc.documentElement, "showMasterSp") !== "0";
    const showMasterOnLayout = !layoutDoc || attrOf(layoutDoc.documentElement, "showMasterSp") !== "0";
    const backdropElements = [];
    if (showMasterOnSlide) {
      if (masterDoc && showMasterOnLayout) {
        backdropElements.push(
          ...parseSpTreeElements(
            firstDescendantByLocal(masterDoc, "spTree"),
            { files, rels: masterRels, theme, clrMap, layoutPhs: new Map(), masterPhs: new Map(), slideW, slideH },
            { skipPlaceholders: true, prefix: `m${idx}` }
          )
        );
      }
      if (layoutDoc) {
        backdropElements.push(
          ...parseSpTreeElements(
            firstDescendantByLocal(layoutDoc, "spTree"),
            { files, rels: layoutRels, theme, clrMap, layoutPhs: new Map(), masterPhs, slideW, slideH },
            { skipPlaceholders: true, prefix: `l${idx}` }
          )
        );
      }
    }

    const elements = parseSpTreeElements(
      firstDescendantByLocal(sDoc, "spTree"),
      { files, rels: sRels, theme, clrMap, layoutPhs, masterPhs, slideW, slideH },
      { skipPlaceholders: false, prefix: `s${idx}` }
    );

    const { notes, notesPath } = parseSlideNotes(files, sRels);

    return {
      id: `slide-${idx + 1}`,
      sourceSlidePath: slidePath,
      layoutPath: layoutPath || defaultLayoutPath,
      notesPath,
      background: bg,
      backdropElements,
      elements,
      notes,
      dirty: false,
      notesDirty: false,
    };
  });

  return {
    slideW,
    slideH,
    theme,
    defaultLayoutPath,
    slides,
    files,
  };
}

// Lossless-preserving PPTX serializer: untouched slides keep their exact bytes;
// edited slides mutate their XML DOM in place; added/reordered/deleted slides
// update presentation.xml, .rels, and [Content_Types].xml cleanly.

function buildTxBodyXml(paragraphs, vAlign = "top", padding = { top: 6, right: 10, bottom: 6, left: 10 }) {
  const anchor = vAlign === "middle" ? "ctr" : vAlign === "bottom" ? "b" : "t";
  const lIns = pxToEmu(padding?.left ?? 10);
  const tIns = pxToEmu(padding?.top ?? 6);
  const rIns = pxToEmu(padding?.right ?? 10);
  const bIns = pxToEmu(padding?.bottom ?? 6);
  const paras = (paragraphs && paragraphs.length ? paragraphs : [{ align: "left", bullet: false, level: 0, runs: [{ text: "" }] }])
    .map((p) => {
      const algn = unmapPptAlign(p.align);
      const lvlAttr = p.level ? ` lvl="${p.level}"` : "";
      const buXml = p.bullet ? `<a:buChar char="•"/>` : `<a:buNone/>`;
      const pPr = `<a:pPr algn="${algn}"${lvlAttr}>${buXml}</a:pPr>`;
      const runs = (p.runs && p.runs.length ? p.runs : [{ text: "" }])
        .map((r) => {
          const text = String(r.text ?? "");
          if (text === "\n") return `<a:br/>`;
          const sz = Math.round((Number(r.fontSize) || 18) * 100);
          const b = r.bold ? ` b="1"` : ` b="0"`;
          const i = r.italic ? ` i="1"` : ` i="0"`;
          const u = r.underline ? ` u="sng"` : "";
          const hex = String(r.color || "#181B22").replace(/^#/, "").toUpperCase();
          const clrXml = /^[0-9A-F]{6}$/.test(hex) ? `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>` : "";
          const fontXml = r.fontFamily ? `<a:latin typeface="${escapeXml(r.fontFamily)}"/>` : "";
          return `<a:r><a:rPr lang="en-US" sz="${sz}"${b}${i}${u} dirty="0">${clrXml}${fontXml}</a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`;
        })
        .join("");
      return `<a:p>${pPr}${runs}</a:p>`;
    })
    .join("");
  return `<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr wrap="square" rtlCol="0" anchor="${anchor}" lIns="${lIns}" tIns="${tIns}" rIns="${rIns}" bIns="${bIns}"/><a:lstStyle/>${paras}</p:txBody>`;
}

function buildShapeXml(el, idNum) {
  const x = pxToEmu(el.x);
  const y = pxToEmu(el.y);
  const cx = pxToEmu(Math.max(8, el.w));
  const cy = pxToEmu(Math.max(8, el.h));
  const rotAttr = el.rot ? ` rot="${Math.round(el.rot * 60000)}"` : "";
  const geom = escapeXml(el.geom || "rect");
  const fillHex = String(el.fillHex || "").replace(/^#/, "").toUpperCase();
  const fillXml = /^[0-9A-F]{6}$/.test(fillHex) ? `<a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>` : `<a:noFill/>`;
  let lnXml = `<a:ln><a:noFill/></a:ln>`;
  if (el.stroke && el.stroke.hex) {
    const sHex = String(el.stroke.hex).replace(/^#/, "").toUpperCase();
    const w = pxToEmu(el.stroke.width || 1.5);
    if (/^[0-9A-F]{6}$/.test(sHex)) {
      lnXml = `<a:ln w="${w}"><a:solidFill><a:srgbClr val="${sHex}"/></a:solidFill></a:ln>`;
    }
  }
  const txXml = buildTxBodyXml(el.paragraphs, el.vAlign, el.padding);
  return (
    `<p:sp xmlns:p="${P_NS}" xmlns:a="${A_NS}">` +
    `<p:nvSpPr><p:cNvPr id="${idNum}" name="${escapeXml(el.name || `Shape ${idNum}`)}"/><p:cNvSpPr txBox="${fillHex ? "0" : "1"}"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${fillXml}${lnXml}</p:spPr>` +
    `${txXml}</p:sp>`
  );
}

function applyXfrmUpdate(spPrOrFrame, el, doc) {
  let xfrm = firstChildByLocal(spPrOrFrame, "xfrm");
  if (!xfrm) {
    xfrm = doc.createElementNS(A_NS, "a:xfrm");
    if (spPrOrFrame.firstChild) spPrOrFrame.insertBefore(xfrm, spPrOrFrame.firstChild);
    else spPrOrFrame.appendChild(xfrm);
  }
  if (el.rot) xfrm.setAttribute("rot", String(Math.round(el.rot * 60000)));
  let off = firstChildByLocal(xfrm, "off");
  if (!off) {
    off = doc.createElementNS(A_NS, "a:off");
    xfrm.appendChild(off);
  }
  off.setAttribute("x", String(pxToEmu(el.x)));
  off.setAttribute("y", String(pxToEmu(el.y)));

  let ext = firstChildByLocal(xfrm, "ext");
  if (!ext) {
    ext = doc.createElementNS(A_NS, "a:ext");
    xfrm.appendChild(ext);
  }
  ext.setAttribute("cx", String(pxToEmu(Math.max(4, el.w))));
  ext.setAttribute("cy", String(pxToEmu(Math.max(4, el.h))));
}

function mutateSlideDom(slideXml, slide) {
  const doc = parseXml(slideXml);
  const cSld = firstDescendantByLocal(doc, "cSld");
  const spTree = cSld && firstChildByLocal(cSld, "spTree");
  if (!spTree) return slideXml;

  // Update slide solid background if user picked one
  if (slide.backgroundChanged && slide.background?.hex) {
    const hex = slide.background.hex.replace(/^#/, "").toUpperCase();
    if (/^[0-9A-F]{6}$/.test(hex)) {
      const oldBg = firstChildByLocal(cSld, "bg");
      if (oldBg) cSld.removeChild(oldBg);
      const bgFrag = parseXml(
        `<p:bg xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`
      ).documentElement;
      cSld.insertBefore(doc.importNode(bgFrag, true), spTree);
    }
  }

  const byXmlId = new Map();
  let maxId = 10;
  for (const el of slide.elements) {
    if (el.xmlId && !el.isNew) byXmlId.set(String(el.xmlId), el);
    const n = Number(el.xmlId);
    if (Number.isFinite(n)) maxId = Math.max(maxId, n);
  }

  // Walk existing direct children of spTree (keep nvGrpSpPr / grpSpPr; only update/remove non-grouped shapes/pics/tables)
  const toRemove = [];
  for (const ch of childrenByLocal(spTree)) {
    const tag = localOf(ch);
    if (tag === "nvGrpSpPr" || tag === "grpSpPr" || tag === "grpSp") continue;
    const cNvPr = firstDescendantByLocal(ch, "cNvPr");
    const id = cNvPr ? attrOf(cNvPr, "id") : "";
    const n = Number(id);
    if (Number.isFinite(n)) maxId = Math.max(maxId, n);
    if (!id) continue;

    if (!byXmlId.has(id)) {
      // Deleted by user (unless it was a hidden/skipped placeholder we never loaded — check if slide tracked deletedXmlIds)
      if (slide.deletedXmlIds && slide.deletedXmlIds.includes(id)) {
        toRemove.push(ch);
      }
      continue;
    }

    const updated = byXmlId.get(id);
    if (tag === "sp" && updated.kind === "shape") {
      const spPr = firstChildByLocal(ch, "spPr");
      if (spPr && !updated.inGroup) applyXfrmUpdate(spPr, updated, doc);
      if (spPr && updated.fillChanged) {
        for (const fTag of ["noFill", "solidFill", "gradFill"]) {
          const old = firstChildByLocal(spPr, fTag);
          if (old) spPr.removeChild(old);
        }
        const hex = String(updated.fillHex || "").replace(/^#/, "").toUpperCase();
        const fillNode = /^[0-9A-F]{6}$/.test(hex)
          ? parseXml(`<a:solidFill xmlns:a="${A_NS}"><a:srgbClr val="${hex}"/></a:solidFill>`).documentElement
          : parseXml(`<a:noFill xmlns:a="${A_NS}"/>`).documentElement;
        const ln = firstChildByLocal(spPr, "ln");
        if (ln) spPr.insertBefore(doc.importNode(fillNode, true), ln);
        else spPr.appendChild(doc.importNode(fillNode, true));
      }
      if (updated.textChanged) {
        const oldTx = firstChildByLocal(ch, "txBody");
        const newTx = parseXml(buildTxBodyXml(updated.paragraphs, updated.vAlign, updated.padding)).documentElement;
        if (oldTx) ch.replaceChild(doc.importNode(newTx, true), oldTx);
        else ch.appendChild(doc.importNode(newTx, true));
      }
    } else if (tag === "pic" && updated.kind === "image") {
      const spPr = firstChildByLocal(ch, "spPr");
      if (spPr && !updated.inGroup) applyXfrmUpdate(spPr, updated, doc);
    } else if (tag === "graphicFrame" && updated.kind === "table") {
      if (!updated.inGroup) applyXfrmUpdate(ch, updated, doc);
      if (updated.textChanged) {
        const tbl = firstDescendantByLocal(ch, "tbl");
        const trEls = tbl ? childrenByLocal(tbl, "tr") : [];
        trEls.forEach((trEl, rIdx) => {
          const modelRow = updated.rows?.[rIdx];
          if (!modelRow) return;
          let cModelIdx = 0;
          for (const tcEl of childrenByLocal(trEl, "tc")) {
            if (attrOf(tcEl, "vMerge") === "1" || attrOf(tcEl, "hMerge") === "1") continue;
            const modelCell = modelRow.cells?.[cModelIdx++];
            if (!modelCell) continue;
            const oldTx = firstChildByLocal(tcEl, "txBody");
            const paras = String(modelCell.text ?? "")
              .split("\n")
              .map((line) => ({
                align: modelCell.align || "left",
                bullet: false,
                level: 0,
                runs: [{ text: line, fontSize: modelCell.fontSize || 14, bold: !!modelCell.bold, color: modelCell.color || "#181B22" }],
              }));
            const newTx = parseXml(buildTxBodyXml(paras, "top", { top: 4, right: 6, bottom: 4, left: 6 })).documentElement;
            if (oldTx) tcEl.replaceChild(doc.importNode(newTx, true), oldTx);
          }
        });
      }
    }
  }

  for (const rem of toRemove) spTree.removeChild(rem);

  for (const el of slide.elements) {
    if (el.isNew && el.kind === "shape") {
      const newId = ++maxId;
      const frag = parseXml(buildShapeXml(el, newId)).documentElement;
      spTree.appendChild(doc.importNode(frag, true));
    }
  }

  return serializeXml(doc);
}

function buildFreshSlideXml(slide) {
  let bgXml = "";
  const hex = String(slide.background?.hex || "#FFFFFF").replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex) && hex !== "FFFFFF") {
    bgXml = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
  }
  let idNum = 2;
  const shapesXml = (slide.elements || [])
    .filter((el) => el.kind === "shape")
    .map((el) => buildShapeXml(el, idNum++))
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:cSld>${bgXml}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `${shapesXml}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

function writeSlideNotes(files, slidePath, existingNotesPath, notesText) {
  let notesPath = existingNotesPath;
  if (!notesPath) {
    let idx = 1;
    while (files[`ppt/notesSlides/notesSlide${idx}.xml`]) idx++;
    notesPath = `ppt/notesSlides/notesSlide${idx}.xml`;

    // Link slide -> notesSlide
    const sRelPath = relsPathFor(slidePath);
    const sRelTxt = readTextPart(files, sRelPath) || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"/>`;
    const sRelDoc = parseXml(sRelTxt);
    let maxR = 1;
    for (const r of childrenByLocal(sRelDoc.documentElement, "Relationship")) {
      const m = /^rId(\d+)$/.exec(attrOf(r, "Id"));
      if (m) maxR = Math.max(maxR, Number(m[1]) + 1);
    }
    const relEl = sRelDoc.createElementNS(PKG_REL_NS, "Relationship");
    relEl.setAttribute("Id", `rId${maxR}`);
    relEl.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide");
    relEl.setAttribute("Target", "../" + notesPath.slice("ppt/".length));
    sRelDoc.documentElement.appendChild(relEl);
    writeTextPart(files, sRelPath, serializeXml(sRelDoc));

    ensureContentTypeOverride(
      files,
      "/" + notesPath,
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
    );
  }

  const paras = String(notesText || "")
    .split("\n")
    .map((line) => `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`)
    .join("");
  const notesXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>` +
    `</p:spTree></p:cSld></p:notes>`;
  writeTextPart(files, notesPath, notesXml);
}

export function savePptx(session, slides) {
  const files = { ...session.files };
  let nextSlideNum = 1;
  for (const k of Object.keys(files)) {
    const m = /^ppt\/slides\/slide(\d+)\.xml$/i.exec(k);
    if (m) nextSlideNum = Math.max(nextSlideNum, Number(m[1]) + 1);
  }

  const usedPaths = new Set();
  const finalSlidePaths = [];

  for (const s of slides) {
    let slidePath = s.sourceSlidePath;
    const isDuplicateReuse = slidePath && usedPaths.has(slidePath);
    if (!slidePath || !files[slidePath] || isDuplicateReuse) {
      const targetPath = `ppt/slides/slide${nextSlideNum++}.xml`;
      if (slidePath && files[slidePath]) {
        // Duplicated slide: clone original XML + .rels first, then apply mutations if dirty
        const baseXml = readTextPart(files, slidePath);
        writeTextPart(files, targetPath, s.dirty ? mutateSlideDom(baseXml, s) : baseXml);
        const origRels = readTextPart(files, relsPathFor(slidePath));
        if (origRels) writeTextPart(files, relsPathFor(targetPath), origRels);
      } else {
        // Brand-new slide
        writeTextPart(files, targetPath, buildFreshSlideXml(s));
        const layoutRelTarget = "../" + (s.layoutPath || session.defaultLayoutPath || "ppt/slideLayouts/slideLayout1.xml").replace(/^ppt\//, "");
        writeTextPart(
          files,
          relsPathFor(targetPath),
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="${PKG_REL_NS}">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${escapeXml(layoutRelTarget)}"/>` +
            `</Relationships>`
        );
      }
      ensureContentTypeOverride(
        files,
        "/" + targetPath,
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
      );
      slidePath = targetPath;
    } else if (s.dirty) {
      const origXml = readTextPart(files, slidePath);
      writeTextPart(files, slidePath, mutateSlideDom(origXml, s));
    }

    if (s.notesDirty) {
      writeSlideNotes(files, slidePath, isDuplicateReuse ? "" : s.notesPath, s.notes);
    }

    usedPaths.add(slidePath);
    finalSlidePaths.push(slidePath);
  }

  // Check if slide list order/count changed compared to original session.slides
  const origOrder = (session.slides || []).map((s) => s.sourceSlidePath);
  const orderChanged =
    origOrder.length !== finalSlidePaths.length ||
    origOrder.some((p, i) => p !== finalSlidePaths[i]);

  if (orderChanged) {
    const presRelPath = relsPathFor("ppt/presentation.xml");
    const presRelDoc = parseXml(readTextPart(files, presRelPath));
    const relRoot = presRelDoc.documentElement;
    let maxRelId = 1;
    const pathToRelId = new Map();
    for (const r of childrenByLocal(relRoot, "Relationship")) {
      const id = attrOf(r, "Id");
      const m = /^rId(\d+)$/.exec(id);
      if (m) maxRelId = Math.max(maxRelId, Number(m[1]) + 1);
      if (attrOf(r, "Type").endsWith("/slide")) {
        const norm = "ppt/" + attrOf(r, "Target").replace(/^\.\//, "").replace(/^\/ppt\//, "");
        if (!finalSlidePaths.includes(norm)) {
          relRoot.removeChild(r);
        } else {
          pathToRelId.set(norm, id);
        }
      }
    }
    for (const sp of finalSlidePaths) {
      if (!pathToRelId.has(sp)) {
        const id = `rId${maxRelId++}`;
        const el = presRelDoc.createElementNS(PKG_REL_NS, "Relationship");
        el.setAttribute("Id", id);
        el.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
        el.setAttribute("Target", sp.slice("ppt/".length));
        relRoot.appendChild(el);
        pathToRelId.set(sp, id);
      }
    }
    writeTextPart(files, presRelPath, serializeXml(presRelDoc));

    // Update p:sldIdLst in ppt/presentation.xml
    const presDoc = parseXml(readTextPart(files, "ppt/presentation.xml"));
    let sldIdLst = firstDescendantByLocal(presDoc, "sldIdLst");
    if (!sldIdLst) {
      sldIdLst = presDoc.createElementNS(P_NS, "p:sldIdLst");
      const sldSz = firstDescendantByLocal(presDoc, "sldSz");
      if (sldSz && sldSz.parentNode) sldSz.parentNode.insertBefore(sldIdLst, sldSz);
      else presDoc.documentElement.appendChild(sldIdLst);
    }
    while (sldIdLst.firstChild) sldIdLst.removeChild(sldIdLst.firstChild);
    let numId = 256;
    for (const sp of finalSlidePaths) {
      const rId = pathToRelId.get(sp);
      const item = presDoc.createElementNS(P_NS, "p:sldId");
      item.setAttribute("id", String(numId++));
      item.setAttributeNS(R_NS, "r:id", rId);
      sldIdLst.appendChild(item);
    }
    writeTextPart(files, "ppt/presentation.xml", serializeXml(presDoc));

    // Remove orphaned deleted slide XML files
    for (const oldPath of origOrder) {
      if (oldPath && !usedPaths.has(oldPath)) {
        delete files[oldPath];
        delete files[relsPathFor(oldPath)];
      }
    }
  }

  return zipPackage(files);
}

export function createBlankPptx(initialSlides = null) {
  const cx = 12192000;
  const cy = 6858000;
  const files = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
        `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
        `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
        `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
        `</Types>`
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
        `</Relationships>`
    ),
    "ppt/presentation.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
        `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
        `<p:sldIdLst/>` +
        `<p:sldSz cx="${cx}" cy="${cy}"/>` +
        `<p:notesSz cx="${cy}" cy="${cx}"/>` +
        `</p:presentation>`
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
        `</Relationships>`
    ),
    "ppt/slideMasters/slideMaster1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
        `<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
        `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
        `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
        `</p:sldMaster>`
    ),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>` +
        `</Relationships>`
    ),
    "ppt/slideLayouts/slideLayout1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`
    ),
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
        `</Relationships>`
    ),
    "ppt/theme/theme1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<a:theme xmlns:a="${A_NS}" name="COGNIX"><a:themeElements>` +
        `<a:clrScheme name="COGNIX"><a:dk1><a:srgbClr val="181B22"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="F2F4F7"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="DC2626"/></a:accent2><a:accent3><a:srgbClr val="16A34A"/></a:accent3><a:accent4><a:srgbClr val="7C3AED"/></a:accent4><a:accent5><a:srgbClr val="0891B2"/></a:accent5><a:accent6><a:srgbClr val="EA580C"/></a:accent6><a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink></a:clrScheme>` +
        `<a:fontScheme name="COGNIX"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>` +
        `<a:fmtScheme name="COGNIX"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
        `</a:themeElements></a:theme>`
    ),
  };

  const slides =
    initialSlides ||
    [
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
            paragraphs: [{ align: "center", bullet: false, level: 0, runs: [{ text: "Presentation", fontSize: 40, bold: true, color: "#181B22" }] }],
          },
        ],
        notes: "",
        dirty: true,
        notesDirty: false,
      },
    ];

  return savePptx({ files, defaultLayoutPath: "ppt/slideLayouts/slideLayout1.xml", slides: [] }, slides);
}
