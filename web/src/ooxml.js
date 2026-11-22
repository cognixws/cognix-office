import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

// Shared OOXML container, XML DOM, relationships, media, and DrawingML theme
// helpers for the Word (.docx) and PowerPoint (.pptx) artifact editors.
// Runs pure-JS inside the strict artifact sandbox (connect-src 'none', no WASM).

export function base64ToBytes(b64) {
  if (b64 instanceof Uint8Array) return b64;
  const clean = String(b64 || "").replace(/\s+/g, "");
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(clean, "base64"));
  }
  const bin = atob(clean);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function bytesToBase64(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (typeof Buffer !== "undefined") {
    return Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength).toString("base64");
  }
  let out = "";
  for (let i = 0; i < u8.length; i += 0x8000) {
    out += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  }
  return btoa(out);
}

export function unzipPackage(input) {
  const bytes = base64ToBytes(input);
  if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error("not a ZIP container");
  }
  return unzipSync(bytes);
}

export function zipPackage(filesMap) {
  const entries = {};
  for (const [k, v] of Object.entries(filesMap)) {
    if (v == null) continue;
    const u8 = typeof v === "string" ? strToU8(v) : v;
    const lower = k.toLowerCase();
    const isCompressedMedia = /\.(png|jpe?g|gif|webp)$/i.test(lower);
    entries[k] = isCompressedMedia ? [u8, { level: 0 }] : [u8, { level: 6 }];
  }
  return bytesToBase64(zipSync(entries));
}

export function readTextPart(files, path) {
  const raw = files[path];
  if (!raw) return null;
  return typeof raw === "string" ? raw : strFromU8(raw);
}

export function writeTextPart(files, path, text) {
  files[path] = strToU8(text);
}

export function parseXml(xmlText) {
  const Parser = globalThis.DOMParser;
  if (!Parser) throw new Error("DOMParser unavailable");
  const doc = new Parser().parseFromString(String(xmlText || ""), "application/xml");
  const err = doc.getElementsByTagName("parsererror")[0];
  if (err) throw new Error("invalid XML");
  return doc;
}

export function serializeXml(doc) {
  const Serializer = globalThis.XMLSerializer;
  if (!Serializer) throw new Error("XMLSerializer unavailable");
  let out = new Serializer().serializeToString(doc);
  if (!out.startsWith("<?xml")) {
    out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out;
  }
  return out;
}

export function localOf(node) {
  if (!node || node.nodeType !== 1) return "";
  if (node.localName) return node.localName;
  const tag = node.nodeName || node.tagName || "";
  const idx = tag.indexOf(":");
  return idx >= 0 ? tag.slice(idx + 1) : tag;
}

export function attrOf(el, name) {
  if (!el || el.nodeType !== 1) return "";
  if (el.hasAttribute(name)) return el.getAttribute(name) || "";
  const attrs = el.attributes;
  if (!attrs) return "";
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs[i];
    const aName = a.name || a.nodeName || "";
    if (aName === name || aName.endsWith(":" + name)) return a.value || "";
  }
  return "";
}

export function childrenByLocal(parent, localName) {
  const out = [];
  if (!parent || !parent.childNodes) return out;
  for (let i = 0; i < parent.childNodes.length; i++) {
    const ch = parent.childNodes[i];
    if (ch.nodeType === 1 && (!localName || localOf(ch) === localName)) out.push(ch);
  }
  return out;
}

export function firstChildByLocal(parent, localName) {
  if (!parent || !parent.childNodes) return null;
  for (let i = 0; i < parent.childNodes.length; i++) {
    const ch = parent.childNodes[i];
    if (ch.nodeType === 1 && localOf(ch) === localName) return ch;
  }
  return null;
}

export function descendantsByLocal(root, localName) {
  const out = [];
  function walk(node) {
    if (!node || !node.childNodes) return;
    for (let i = 0; i < node.childNodes.length; i++) {
      const ch = node.childNodes[i];
      if (ch.nodeType === 1) {
        if (localOf(ch) === localName) out.push(ch);
        walk(ch);
      }
    }
  }
  walk(root);
  return out;
}

export function firstDescendantByLocal(root, localName) {
  if (!root || !root.childNodes) return null;
  for (let i = 0; i < root.childNodes.length; i++) {
    const ch = root.childNodes[i];
    if (ch.nodeType === 1) {
      if (localOf(ch) === localName) return ch;
      const found = firstDescendantByLocal(ch, localName);
      if (found) return found;
    }
  }
  return null;
}

export function normalizeZipPath(baseDir, target) {
  const raw = String(target || "").replace(/\\/g, "/");
  if (!raw) return "";
  if (/^[a-z][a-z0-9+.-]*:/i.test(raw)) return raw;
  const combined = raw.startsWith("/") ? raw.slice(1) : (baseDir ? baseDir + "/" + raw : raw);
  const parts = [];
  for (const seg of combined.split("/")) {
    if (!seg || seg === ".") continue;
    if (seg === "..") parts.pop();
    else parts.push(seg);
  }
  return parts.join("/");
}

export function relsPathFor(partPath) {
  const idx = partPath.lastIndexOf("/");
  const dir = idx >= 0 ? partPath.slice(0, idx) : "";
  const base = idx >= 0 ? partPath.slice(idx + 1) : partPath;
  return (dir ? dir + "/_rels/" : "_rels/") + base + ".rels";
}

export function dirOf(partPath) {
  const idx = partPath.lastIndexOf("/");
  return idx >= 0 ? partPath.slice(0, idx) : "";
}

export function parseRels(files, partPath) {
  const relFile = relsPathFor(partPath);
  const txt = readTextPart(files, relFile);
  const map = new Map();
  if (!txt) return map;
  let doc;
  try {
    doc = parseXml(txt);
  } catch {
    return map;
  }
  const baseDir = dirOf(partPath);
  for (const rel of descendantsByLocal(doc, "Relationship")) {
    const id = attrOf(rel, "Id");
    const target = attrOf(rel, "Target");
    const mode = attrOf(rel, "TargetMode");
    const type = attrOf(rel, "Type");
    if (!id || !target) continue;
    map.set(id, {
      id,
      type,
      external: mode === "External",
      target: mode === "External" ? target : normalizeZipPath(baseDir, target),
      rawTarget: target,
    });
  }
  return map;
}

const MIME_BY_EXT = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  svg: "image/svg+xml",
  bmp: "image/bmp",
};

export function mediaToDataUrl(files, pkgPath) {
  const raw = files[pkgPath];
  if (!raw) return "";
  const ext = (pkgPath.split(".").pop() || "").toLowerCase();
  const mime = MIME_BY_EXT[ext];
  if (!mime) return "";
  const u8 = typeof raw === "string" ? strToU8(raw) : raw;
  return "data:" + mime + ";base64," + bytesToBase64(u8);
}

export function dataUrlToMedia(dataUrl) {
  const m = /^data:image\/([a-z0-9+.-]+);base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(dataUrl || ""));
  if (!m) return null;
  let ext = m[1].toLowerCase();
  if (ext === "jpeg") ext = "jpg";
  if (ext === "svg+xml") ext = "svg";
  return { ext, mime: "image/" + m[1].toLowerCase(), bytes: base64ToBytes(m[2]) };
}

// DrawingML color and theme helpers
const PRESET_COLORS = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#008000",
  blue: "#0000FF",
  yellow: "#FFFF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  gray: "#808080",
  grey: "#808080",
  darkGray: "#A9A9A9",
  lightGray: "#D3D3D3",
  navy: "#000080",
  teal: "#008080",
  maroon: "#800000",
  olive: "#808000",
  purple: "#800080",
  silver: "#C0C0C0",
  orange: "#FFA500",
};

const DEFAULT_THEME_COLORS = {
  dk1: "#000000",
  lt1: "#FFFFFF",
  dk2: "#1F497D",
  lt2: "#EEECE1",
  accent1: "#4F81BD",
  accent2: "#C0504D",
  accent3: "#9BBB59",
  accent4: "#8064A2",
  accent5: "#4BACC6",
  accent6: "#F79646",
  hlink: "#0000FF",
  folHlink: "#800080",
};

export function parseTheme(xmlText) {
  const colors = { ...DEFAULT_THEME_COLORS };
  let majorFont = "Calibri";
  let minorFont = "Calibri";
  if (!xmlText) return { colors, majorFont, minorFont };
  try {
    const doc = parseXml(xmlText);
    const clrScheme = firstDescendantByLocal(doc, "clrScheme");
    if (clrScheme) {
      for (const slot of childrenByLocal(clrScheme)) {
        const name = localOf(slot);
        const srgb = firstChildByLocal(slot, "srgbClr");
        const sys = firstChildByLocal(slot, "sysClr");
        const hex = (srgb && attrOf(srgb, "val")) || (sys && (attrOf(sys, "lastClr") || attrOf(sys, "val")));
        if (hex && /^[0-9a-fA-F]{6}$/.test(hex)) {
          colors[name] = "#" + hex.toUpperCase();
        }
      }
    }
    const fontScheme = firstDescendantByLocal(doc, "fontScheme");
    if (fontScheme) {
      const maj = firstChildByLocal(fontScheme, "majorFont");
      const min = firstChildByLocal(fontScheme, "minorFont");
      const majLat = maj && firstChildByLocal(maj, "latin");
      const minLat = min && firstChildByLocal(min, "latin");
      if (majLat && attrOf(majLat, "typeface")) majorFont = attrOf(majLat, "typeface");
      if (minLat && attrOf(minLat, "typeface")) minorFont = attrOf(minLat, "typeface");
    }
  } catch {
    // keep defaults
  }
  return { colors, majorFont, minorFont };
}

function hexToRgb(hex) {
  const h = String(hex || "").replace(/^#/, "");
  if (h.length !== 6) return [0, 0, 0];
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

function rgbToHex(r, g, b) {
  const clamp = (v) => Math.max(0, Math.min(255, Math.round(v)));
  return (
    "#" +
    [clamp(r), clamp(g), clamp(b)]
      .map((x) => x.toString(16).padStart(2, "0"))
      .join("")
      .toUpperCase()
  );
}

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r: h = (g - b) / d + (g < b ? 6 : 0); break;
      case g: h = (b - r) / d + 2; break;
      case b: h = (r - g) / d + 4; break;
    }
    h /= 6;
  }
  return [h, s, l];
}

function hslToRgb(h, s, l) {
  if (s === 0) {
    const v = l * 255;
    return [v, v, v];
  }
  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue2rgb(p, q, h + 1 / 3) * 255, hue2rgb(p, q, h) * 255, hue2rgb(p, q, h - 1 / 3) * 255];
}

function applyColorModifiers(baseHex, clrEl) {
  let [r, g, b] = hexToRgb(baseHex);
  let [h, s, l] = rgbToHsl(r, g, b);
  let alpha = 1;
  for (const mod of childrenByLocal(clrEl)) {
    const kind = localOf(mod);
    const val = Number(attrOf(mod, "val"));
    if (!Number.isFinite(val)) continue;
    const f = val / 100000;
    if (kind === "lumMod") l = Math.max(0, Math.min(1, l * f));
    else if (kind === "lumOff") l = Math.max(0, Math.min(1, l + f));
    else if (kind === "satMod") s = Math.max(0, Math.min(1, s * f));
    else if (kind === "satOff") s = Math.max(0, Math.min(1, s + f));
    else if (kind === "alpha") alpha = Math.max(0, Math.min(1, f));
    else if (kind === "tint") {
      [r, g, b] = hslToRgb(h, s, l);
      r = r + (255 - r) * (1 - f);
      g = g + (255 - g) * (1 - f);
      b = b + (255 - b) * (1 - f);
      [h, s, l] = rgbToHsl(r, g, b);
    } else if (kind === "shade") {
      [r, g, b] = hslToRgb(h, s, l);
      r *= f; g *= f; b *= f;
      [h, s, l] = rgbToHsl(r, g, b);
    }
  }
  [r, g, b] = hslToRgb(h, s, l);
  const hex = rgbToHex(r, g, b);
  return { hex, alpha };
}

const DEFAULT_CLR_MAP = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

export function resolveDrawingColor(containerEl, theme, clrMap = DEFAULT_CLR_MAP, phColor = null) {
  if (!containerEl) return null;
  const srgb = firstChildByLocal(containerEl, "srgbClr");
  if (srgb) {
    const v = attrOf(srgb, "val");
    if (/^[0-9a-fA-F]{6}$/.test(v)) return applyColorModifiers("#" + v.toUpperCase(), srgb);
  }
  const sys = firstChildByLocal(containerEl, "sysClr");
  if (sys) {
    const v = attrOf(sys, "lastClr") || attrOf(sys, "val");
    if (/^[0-9a-fA-F]{6}$/.test(v)) return applyColorModifiers("#" + v.toUpperCase(), sys);
    if (v === "windowText") return applyColorModifiers("#000000", sys);
    if (v === "window") return applyColorModifiers("#FFFFFF", sys);
  }
  const prst = firstChildByLocal(containerEl, "prstClr");
  if (prst) {
    const v = attrOf(prst, "val");
    const hex = PRESET_COLORS[v];
    if (hex) return applyColorModifiers(hex, prst);
  }
  const scheme = firstChildByLocal(containerEl, "schemeClr");
  if (scheme) {
    const rawSlot = attrOf(scheme, "val");
    if (rawSlot === "phClr" && phColor) return applyColorModifiers(phColor, scheme);
    const mapped = (clrMap && clrMap[rawSlot]) || DEFAULT_CLR_MAP[rawSlot] || rawSlot;
    const hex = (theme && theme.colors && theme.colors[mapped]) || DEFAULT_THEME_COLORS[mapped];
    if (hex) return applyColorModifiers(hex, scheme);
  }
  return null;
}

export function cssColor(resolved) {
  if (!resolved) return "";
  if (resolved.alpha != null && resolved.alpha < 0.999) {
    const [r, g, b] = hexToRgb(resolved.hex);
    return `rgba(${r}, ${g}, ${b}, ${Number(resolved.alpha.toFixed(3))})`;
  }
  return resolved.hex;
}

export function escapeXml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function ensureContentTypeDefault(files, ext, contentType) {
  const txt = readTextPart(files, "[Content_Types].xml");
  if (!txt) return;
  try {
    const doc = parseXml(txt);
    const root = doc.documentElement;
    for (const def of childrenByLocal(root, "Default")) {
      if (attrOf(def, "Extension").toLowerCase() === ext.toLowerCase()) return;
    }
    const ns = root.namespaceURI || "http://schemas.openxmlformats.org/package/2006/content-types";
    const el = doc.createElementNS(ns, "Default");
    el.setAttribute("Extension", ext);
    el.setAttribute("ContentType", contentType);
    root.appendChild(el);
    writeTextPart(files, "[Content_Types].xml", serializeXml(doc));
  } catch {
    // leave untouched
  }
}

export function ensureContentTypeOverride(files, partName, contentType) {
  const txt = readTextPart(files, "[Content_Types].xml");
  if (!txt) return;
  try {
    const doc = parseXml(txt);
    const root = doc.documentElement;
    const normalized = partName.startsWith("/") ? partName : "/" + partName;
    for (const ov of childrenByLocal(root, "Override")) {
      if (attrOf(ov, "PartName") === normalized) return;
    }
    const ns = root.namespaceURI || "http://schemas.openxmlformats.org/package/2006/content-types";
    const el = doc.createElementNS(ns, "Override");
    el.setAttribute("PartName", normalized);
    el.setAttribute("ContentType", contentType);
    root.appendChild(el);
    writeTextPart(files, "[Content_Types].xml", serializeXml(doc));
  } catch {
    // leave untouched
  }
}

export function detectOoxmlKind(input) {
  let files;
  try {
    files = unzipPackage(input);
  } catch {
    return "";
  }
  if (!files["[Content_Types].xml"]) return "";
  if (files["word/document.xml"]) return "docx";
  if (files["ppt/presentation.xml"]) return "pptx";
  if (files["xl/workbook.xml"]) return "xlsx";
  return "";
}

export function createBlankXlsx(sheetName = "Sheet1") {
  const safeName = escapeXml(String(sheetName || "Sheet1").slice(0, 31));
  const files = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
        `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
        `</Types>`
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
        `</Relationships>`
    ),
    "xl/workbook.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        `<sheets><sheet name="${safeName}" sheetId="1" r:id="rId1"/></sheets>` +
        `</workbook>`
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
        `</Relationships>`
    ),
    "xl/worksheets/sheet1.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        `<dimension ref="A1"/>` +
        `<sheetData/>` +
        `</worksheet>`
    ),
  };
  return zipPackage(files);
}

