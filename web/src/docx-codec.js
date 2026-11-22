import { strToU8 } from "fflate";
import {
  attrOf,
  childrenByLocal,
  dataUrlToMedia,
  descendantsByLocal,
  ensureContentTypeDefault,
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
  serializeXml,
  unzipPackage,
  writeTextPart,
  zipPackage,
} from "./ooxml.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

const HIGHLIGHT_HEX = {
  yellow: "#FEF08A",
  green: "#BBF7D0",
  cyan: "#A5F3FC",
  magenta: "#F5D0FE",
  blue: "#BFDBFE",
  red: "#FECACA",
  darkBlue: "#93C5FD",
  darkCyan: "#67E8F9",
  darkGreen: "#86EFAC",
  darkMagenta: "#E879F9",
  darkRed: "#FCA5A5",
  darkYellow: "#FDE047",
  lightGray: "#E5E7EB",
  darkGray: "#D1D5DB",
};

function isNearBlackInk(hex) {
  const h = String(hex || "").replace(/^#/, "");
  if (h.length !== 6) return false;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return r < 38 && g < 38 && b < 42;
}

function parseStylesMap(xmlText) {
  const map = new Map();
  if (!xmlText) return map;
  try {
    const doc = parseXml(xmlText);
    for (const st of descendantsByLocal(doc, "style")) {
      const id = attrOf(st, "styleId");
      if (!id) continue;
      const nameEl = firstChildByLocal(st, "name");
      const name = (nameEl && attrOf(nameEl, "val")) || id;
      const lower = (id + " " + name).toLowerCase();
      let kind = "paragraph";
      let level = 0;
      const hm = /heading\s*([1-6])|t[íi]tulo\s*([1-6])|^h([1-6])\b/.exec(lower);
      if (hm) {
        kind = "heading";
        level = Math.min(4, Number(hm[1] || hm[2] || hm[3]));
      } else if (/\btitle\b/.test(lower) && !/subtitle/.test(lower)) {
        kind = "heading";
        level = 1;
      } else if (/subtitle/.test(lower)) {
        kind = "heading";
        level = 2;
      } else if (/quote|cita/.test(lower)) {
        kind = "blockquote";
      } else if (/code|preformatted|verbatim/.test(lower)) {
        kind = "code";
      }
      const pPr = firstChildByLocal(st, "pPr");
      const outline = pPr && firstChildByLocal(pPr, "outlineLvl");
      if (kind === "paragraph" && outline) {
        const v = Number(attrOf(outline, "val"));
        if (Number.isFinite(v) && v >= 0 && v <= 5) {
          kind = "heading";
          level = Math.min(4, v + 1);
        }
      }
      const rPr = firstChildByLocal(st, "rPr");
      map.set(id, { id, name, kind, level, rPr });
    }
  } catch {
    // ignore malformed styles.xml
  }
  return map;
}

function parseNumberingMap(xmlText) {
  const abstractFormats = new Map();
  const numToAbstract = new Map();
  if (!xmlText) return { listTypeOf: () => "bullet" };
  try {
    const doc = parseXml(xmlText);
    for (const abs of descendantsByLocal(doc, "abstractNum")) {
      const absId = attrOf(abs, "abstractNumId");
      const lvlMap = new Map();
      for (const lvl of childrenByLocal(abs, "lvl")) {
        const ilvl = attrOf(lvl, "ilvl") || "0";
        const fmtEl = firstChildByLocal(lvl, "numFmt");
        const fmt = (fmtEl && attrOf(fmtEl, "val")) || "bullet";
        lvlMap.set(ilvl, fmt === "bullet" || fmt === "none" ? "bullet" : "ordered");
      }
      abstractFormats.set(absId, lvlMap);
    }
    for (const num of descendantsByLocal(doc, "num")) {
      const numId = attrOf(num, "numId");
      const absEl = firstChildByLocal(num, "abstractNumId");
      if (numId && absEl) numToAbstract.set(numId, attrOf(absEl, "val"));
    }
  } catch {
    // fallback to bullet
  }
  return {
    listTypeOf(numId, ilvl = "0") {
      const absId = numToAbstract.get(String(numId));
      const lvlMap = absId != null ? abstractFormats.get(absId) : null;
      if (lvlMap && lvlMap.has(String(ilvl))) return lvlMap.get(String(ilvl));
      if (lvlMap && lvlMap.has("0")) return lvlMap.get("0");
      return String(numId) === "2" ? "ordered" : "bullet";
    },
  };
}

function marksFromRPr(rPr, styleRPr, linkHref, theme) {
  const marks = [];
  const hasFlag = (parent, name) => {
    if (!parent) return null;
    const el = firstChildByLocal(parent, name);
    if (!el) return null;
    const v = attrOf(el, "val").toLowerCase();
    return v !== "0" && v !== "false" && v !== "none";
  };
  const bold = hasFlag(rPr, "b") ?? hasFlag(styleRPr, "b") ?? false;
  const italic = hasFlag(rPr, "i") ?? hasFlag(styleRPr, "i") ?? false;
  const uEl = (rPr && firstChildByLocal(rPr, "u")) || (styleRPr && firstChildByLocal(styleRPr, "u"));
  const uVal = uEl ? (attrOf(uEl, "val") || "single").toLowerCase() : "none";
  const underline = uVal !== "none" && uVal !== "0" && uVal !== "false";
  const strike =
    (hasFlag(rPr, "strike") ?? hasFlag(rPr, "dstrike") ?? hasFlag(styleRPr, "strike")) || false;

  const fontsEl = (rPr && firstChildByLocal(rPr, "rFonts")) || (styleRPr && firstChildByLocal(styleRPr, "rFonts"));
  const asciiFont = fontsEl ? attrOf(fontsEl, "ascii") || attrOf(fontsEl, "hAnsi") : "";
  const isMono = /consolas|courier|menlo|monaco|fira code|jetbrains mono|monospace/i.test(asciiFont);

  if (bold) marks.push({ type: "bold" });
  if (italic) marks.push({ type: "italic" });
  if (underline) marks.push({ type: "underline" });
  if (strike) marks.push({ type: "strike" });
  if (isMono) marks.push({ type: "code" });

  const textStyleAttrs = {};
  const colorEl = (rPr && firstChildByLocal(rPr, "color")) || (styleRPr && firstChildByLocal(styleRPr, "color"));
  if (colorEl) {
    const val = attrOf(colorEl, "val");
    const themeClr = attrOf(colorEl, "themeColor");
    let hex = "";
    if (/^[0-9a-fA-F]{6}$/.test(val)) {
      hex = "#" + val.toUpperCase();
    } else if (themeClr && theme && theme.colors && theme.colors[themeClr]) {
      hex = theme.colors[themeClr];
    }
    // Keep body default ink adaptive across dark/light themes unless explicitly colored.
    if (hex && !isNearBlackInk(hex)) {
      textStyleAttrs.color = hex;
    }
  }
  const szEl = rPr && firstChildByLocal(rPr, "sz");
  if (szEl) {
    const halfPt = Number(attrOf(szEl, "val"));
    if (Number.isFinite(halfPt) && halfPt >= 14 && halfPt <= 144 && halfPt !== 22 && halfPt !== 24) {
      textStyleAttrs.fontSize = Math.round(halfPt / 2) + "pt";
    }
  }
  if (Object.keys(textStyleAttrs).length > 0) {
    marks.push({ type: "textStyle", attrs: textStyleAttrs });
  }

  const hlEl = rPr && firstChildByLocal(rPr, "highlight");
  const shdEl = rPr && firstChildByLocal(rPr, "shd");
  if (hlEl) {
    const name = attrOf(hlEl, "val");
    const hex = HIGHLIGHT_HEX[name];
    if (hex) marks.push({ type: "highlight", attrs: { color: hex } });
  } else if (shdEl) {
    const fill = attrOf(shdEl, "fill");
    if (/^[0-9a-fA-F]{6}$/.test(fill) && fill.toUpperCase() !== "FFFFFF") {
      marks.push({ type: "highlight", attrs: { color: "#" + fill.toUpperCase() } });
    }
  }

  if (linkHref) {
    marks.push({ type: "link", attrs: { href: linkHref, target: "_blank" } });
  }
  return marks;
}

function extractDrawingImages(el, files, rels) {
  const images = [];
  for (const blip of descendantsByLocal(el, "blip")) {
    const embed = attrOf(blip, "embed");
    if (!embed) continue;
    const rel = rels.get(embed);
    if (!rel || rel.external) continue;
    const src = mediaToDataUrl(files, rel.target);
    if (!src) continue;
    let width = null;
    const extent = firstDescendantByLocal(el, "extent");
    if (extent) {
      const cx = Number(attrOf(extent, "cx"));
      if (Number.isFinite(cx) && cx > 0) width = Math.min(680, Math.max(24, Math.round(cx / 9525)));
    }
    const docPr = firstDescendantByLocal(el, "docPr");
    const alt = (docPr && (attrOf(docPr, "descr") || attrOf(docPr, "title"))) || "";
    images.push({
      type: "image",
      attrs: { src, alt, title: rel.target, ...(width ? { width } : {}) },
    });
  }
  return images;
}

function parseInlineChildren(container, ctx, linkHref = "") {
  const nodes = [];
  for (const ch of childrenByLocal(container)) {
    const tag = localOf(ch);
    if (tag === "r") {
      const rPr = firstChildByLocal(ch, "rPr");
      let styleRPr = null;
      const rStyle = rPr && firstChildByLocal(rPr, "rStyle");
      if (rStyle) {
        const st = ctx.styles.get(attrOf(rStyle, "val"));
        if (st) styleRPr = st.rPr;
      }
      const marks = marksFromRPr(rPr, styleRPr, linkHref, ctx.theme);
      let textBuf = "";
      const flushText = () => {
        if (!textBuf) return;
        const node = { type: "text", text: textBuf };
        if (marks.length) node.marks = marks;
        nodes.push(node);
        textBuf = "";
      };
      for (const rChild of childrenByLocal(ch)) {
        const rTag = localOf(rChild);
        if (rTag === "t" || rTag === "delText") {
          if (rTag === "t") textBuf += rChild.textContent || "";
        } else if (rTag === "tab") {
          textBuf += "\t";
        } else if (rTag === "br" || rTag === "cr") {
          flushText();
          nodes.push({ type: "hardBreak" });
        } else if (rTag === "drawing" || rTag === "pict") {
          flushText();
          nodes.push(...extractDrawingImages(rChild, ctx.files, ctx.rels));
        }
      }
      flushText();
    } else if (tag === "hyperlink") {
      const rId = attrOf(ch, "id");
      const anchor = attrOf(ch, "anchor");
      const rel = rId ? ctx.rels.get(rId) : null;
      const href = (rel && rel.target) || (anchor ? "#" + anchor : "");
      nodes.push(...parseInlineChildren(ch, ctx, href));
    } else if (tag === "sdt") {
      const content = firstChildByLocal(ch, "sdtContent");
      if (content) nodes.push(...parseInlineChildren(content, ctx, linkHref));
    } else if (tag === "ins" || tag === "smartTag" || tag === "fldSimple") {
      nodes.push(...parseInlineChildren(ch, ctx, linkHref));
    }
  }
  return nodes;
}

function parseParagraph(pEl, ctx) {
  const pPr = firstChildByLocal(pEl, "pPr");
  let kind = "paragraph";
  let level = 1;
  let align = "";
  let listInfo = null;

  if (pPr) {
    const pStyle = firstChildByLocal(pPr, "pStyle");
    if (pStyle) {
      const st = ctx.styles.get(attrOf(pStyle, "val"));
      if (st) {
        kind = st.kind;
        if (st.level) level = st.level;
      }
    }
    const outline = firstChildByLocal(pPr, "outlineLvl");
    if (kind === "paragraph" && outline) {
      const v = Number(attrOf(outline, "val"));
      if (Number.isFinite(v) && v >= 0 && v <= 3) {
        kind = "heading";
        level = v + 1;
      }
    }
    const jc = firstChildByLocal(pPr, "jc");
    if (jc) {
      const v = attrOf(jc, "val").toLowerCase();
      if (v === "center") align = "center";
      else if (v === "right" || v === "end") align = "right";
      else if (v === "both" || v === "distribute") align = "justify";
    }
    const numPr = firstChildByLocal(pPr, "numPr");
    if (numPr) {
      const numIdEl = firstChildByLocal(numPr, "numId");
      const ilvlEl = firstChildByLocal(numPr, "ilvl");
      const numId = numIdEl ? attrOf(numIdEl, "val") : "";
      const ilvl = ilvlEl ? Number(attrOf(ilvlEl, "val")) || 0 : 0;
      if (numId && numId !== "0") {
        listInfo = {
          listType: ctx.numbering.listTypeOf(numId, String(ilvl)),
          numId,
          ilvl: Math.min(3, Math.max(0, ilvl)),
        };
      }
    }
  }

  const inlines = parseInlineChildren(pEl, ctx);
  // Check if paragraph is a horizontal rule (empty text with bottom border).
  if (!inlines.length && pPr) {
    const pBdr = firstChildByLocal(pPr, "pBdr");
    const bottom = pBdr && firstChildByLocal(pBdr, "bottom");
    if (bottom && attrOf(bottom, "val") !== "none") {
      return { block: { type: "horizontalRule" }, listInfo: null };
    }
  }

  // Separate block-level images if any (Tiptap standard Image is block by default unless inline: true; we configure Image with inline: true so images live cleanly inside paragraphs too!).
  const attrs = {};
  if (align) attrs.textAlign = align;

  if (listInfo) {
    const para = { type: "paragraph", ...(Object.keys(attrs).length ? { attrs } : {}), ...(inlines.length ? { content: inlines } : {}) };
    return { block: para, listInfo };
  }

  if (kind === "heading") {
    return {
      block: {
        type: "heading",
        attrs: { level, ...(align ? { textAlign: align } : {}) },
        ...(inlines.length ? { content: inlines } : {}),
      },
      listInfo: null,
    };
  }
  if (kind === "blockquote") {
    return {
      block: {
        type: "blockquote",
        content: [{ type: "paragraph", ...(Object.keys(attrs).length ? { attrs } : {}), ...(inlines.length ? { content: inlines } : {}) }],
      },
      listInfo: null,
    };
  }
  if (kind === "code") {
    const raw = inlines.map((n) => (n.type === "hardBreak" ? "\n" : n.text || "")).join("");
    return {
      block: {
        type: "codeBlock",
        ...(raw ? { content: [{ type: "text", text: raw }] } : {}),
      },
      listInfo: null,
    };
  }

  return {
    block: {
      type: "paragraph",
      ...(Object.keys(attrs).length ? { attrs } : {}),
      ...(inlines.length ? { content: inlines } : {}),
    },
    listInfo: null,
  };
}

function parseTable(tblEl, ctx) {
  const rows = [];
  const trList = childrenByLocal(tblEl, "tr");
  trList.forEach((trEl, rowIdx) => {
    const trPr = firstChildByLocal(trEl, "trPr");
    const isHeaderRow = !!(trPr && firstChildByLocal(trPr, "tblHeader")) || (rowIdx === 0 && trList.length > 1 && hasHeaderShading(trEl));
    const cells = [];
    for (const tcEl of childrenByLocal(trEl, "tc")) {
      const tcPr = firstChildByLocal(tcEl, "tcPr");
      const vMerge = tcPr && firstChildByLocal(tcPr, "vMerge");
      if (vMerge && !attrOf(vMerge, "val")) {
        // Continuation cell of vertical merge
        continue;
      }
      const gridSpanEl = tcPr && firstChildByLocal(tcPr, "gridSpan");
      const colspan = gridSpanEl ? Math.max(1, Number(attrOf(gridSpanEl, "val")) || 1) : 1;
      const cellBlocks = parseBlockSequence(tcEl, ctx);
      if (!cellBlocks.length) cellBlocks.push({ type: "paragraph" });
      cells.push({
        type: isHeaderRow ? "tableHeader" : "tableCell",
        attrs: { colspan, rowspan: 1 },
        content: cellBlocks,
      });
    }
    if (cells.length) rows.push({ type: "tableRow", content: cells });
  });
  if (!rows.length) return null;
  return { type: "table", content: rows };
}

function hasHeaderShading(trEl) {
  const tcs = childrenByLocal(trEl, "tc");
  if (!tcs.length) return false;
  return tcs.every((tc) => {
    const tcPr = firstChildByLocal(tc, "tcPr");
    const shd = tcPr && firstChildByLocal(tcPr, "shd");
    const fill = shd ? attrOf(shd, "fill").toUpperCase() : "";
    return fill && fill !== "AUTO" && fill !== "FFFFFF";
  });
}

function parseBlockSequence(container, ctx) {
  const blocks = [];
  let activeList = null;

  const flushList = () => {
    if (activeList) {
      blocks.push(activeList.node);
      activeList = null;
    }
  };

  for (const ch of childrenByLocal(container)) {
    const tag = localOf(ch);
    if (tag === "p") {
      const { block, listInfo } = parseParagraph(ch, ctx);
      if (listInfo) {
        const listType = listInfo.listType === "ordered" ? "orderedList" : "bulletList";
        if (!activeList || activeList.type !== listType) {
          flushList();
          activeList = { type: listType, node: { type: listType, content: [] } };
        }
        activeList.node.content.push({ type: "listItem", content: [block] });
      } else {
        flushList();
        blocks.push(block);
      }
    } else if (tag === "tbl") {
      flushList();
      const tbl = parseTable(ch, ctx);
      if (tbl) blocks.push(tbl);
    } else if (tag === "sdt") {
      const content = firstChildByLocal(ch, "sdtContent");
      if (content) {
        flushList();
        blocks.push(...parseBlockSequence(content, ctx));
      }
    }
  }
  flushList();
  return blocks;
}

export function parseDocx(input) {
  const files = unzipPackage(input);
  const docXml = readTextPart(files, "word/document.xml");
  if (!docXml) throw new Error("missing word/document.xml");
  const dom = parseXml(docXml);
  const body = firstDescendantByLocal(dom, "body");
  if (!body) throw new Error("missing w:body");

  const rels = parseRels(files, "word/document.xml");
  const styles = parseStylesMap(readTextPart(files, "word/styles.xml"));
  const numbering = parseNumberingMap(readTextPart(files, "word/numbering.xml"));
  const theme = parseTheme(readTextPart(files, "word/theme/theme1.xml"));

  const sectPr = firstChildByLocal(body, "sectPr");
  let sectPrXml = "";
  let pageWidthPx = 816;
  let marginPx = { top: 72, right: 72, bottom: 72, left: 72 };
  if (sectPr) {
    const Serializer = globalThis.XMLSerializer;
    if (Serializer) sectPrXml = new Serializer().serializeToString(sectPr);
    const pgSz = firstChildByLocal(sectPr, "pgSz");
    if (pgSz) {
      const wTwip = Number(attrOf(pgSz, "w"));
      if (Number.isFinite(wTwip) && wTwip > 2880) pageWidthPx = Math.round((wTwip / 1440) * 96);
    }
    const pgMar = firstChildByLocal(sectPr, "pgMar");
    if (pgMar) {
      const tw = (k, d) => {
        const v = Number(attrOf(pgMar, k));
        return Number.isFinite(v) && v >= 0 ? Math.min(120, Math.round((v / 1440) * 96)) : d;
      };
      marginPx = {
        top: tw("top", 72),
        right: tw("right", 72),
        bottom: tw("bottom", 72),
        left: tw("left", 72),
      };
    }
  }

  const content = parseBlockSequence(body, { files, rels, styles, numbering, theme });
  if (!content.length) content.push({ type: "paragraph" });

  return {
    doc: { type: "doc", content },
    page: { widthPx: pageWidthPx, marginPx },
    files,
    sectPrXml,
  };
}

// Serialization: Tiptap JSON AST -> OOXML word/document.xml while preserving
// untouched package parts, headers/footers, styles, media, and sectPr.

const DEFAULT_NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;

const DEFAULT_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="34"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="40"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:pPr><w:keepNext/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:pPr><w:ind w:left="720"/><w:spacing w:before="120" w:after="120"/></w:pPr><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>
</w:styles>`;

function buildRelsManager(files) {
  const relPath = relsPathFor("word/document.xml");
  const existingXml = readTextPart(files, relPath);
  let doc;
  if (existingXml) {
    doc = parseXml(existingXml);
  } else {
    doc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"/>`);
  }
  const root = doc.documentElement;
  let maxId = 0;
  const byTargetAndType = new Map();
  for (const rel of childrenByLocal(root, "Relationship")) {
    const id = attrOf(rel, "Id");
    const m = /^rId(\d+)$/.exec(id);
    if (m) maxId = Math.max(maxId, Number(m[1]));
    const key = attrOf(rel, "Type") + "|" + attrOf(rel, "Target") + "|" + attrOf(rel, "TargetMode");
    byTargetAndType.set(key, id);
  }

  let mediaSeq = 1;
  for (const k of Object.keys(files)) {
    const m = /^word\/media\/image(\d+)\./i.exec(k);
    if (m) mediaSeq = Math.max(mediaSeq, Number(m[1]) + 1);
  }

  function ensureRel(type, target, external = false) {
    const mode = external ? "External" : "";
    const key = type + "|" + target + "|" + mode;
    if (byTargetAndType.has(key)) return byTargetAndType.get(key);
    const id = "rId" + (++maxId);
    const el = doc.createElementNS(PKG_REL_NS, "Relationship");
    el.setAttribute("Id", id);
    el.setAttribute("Type", type);
    el.setAttribute("Target", target);
    if (external) el.setAttribute("TargetMode", "External");
    root.appendChild(el);
    byTargetAndType.set(key, id);
    return id;
  }

  function addHyperlink(href) {
    return ensureRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", href, true);
  }

  function addImage(attrs) {
    // Reuse original package media path if stored in attrs.title and still present
    if (attrs.title && files[attrs.title] && attrs.title.startsWith("word/")) {
      const relTarget = attrs.title.slice("word/".length);
      return ensureRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", relTarget, false);
    }
    const media = dataUrlToMedia(attrs.src);
    if (!media) return "";
    const fileName = `image${mediaSeq++}.${media.ext}`;
    files["word/media/" + fileName] = media.bytes;
    ensureContentTypeDefault(files, media.ext, media.mime);
    return ensureRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/" + fileName, false);
  }

  function ensureNumberingPart() {
    if (!files["word/numbering.xml"]) {
      writeTextPart(files, "word/numbering.xml", DEFAULT_NUMBERING_XML);
      ensureContentTypeOverride(
        files,
        "/word/numbering.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
      );
    }
    ensureRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml", false);
  }

  function ensureStylesPart() {
    if (!files["word/styles.xml"]) {
      writeTextPart(files, "word/styles.xml", DEFAULT_STYLES_XML);
      ensureContentTypeOverride(
        files,
        "/word/styles.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
      );
    }
    ensureRel("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml", false);
  }

  function flush() {
    writeTextPart(files, relPath, serializeXml(doc));
  }

  return { addHyperlink, addImage, ensureNumberingPart, ensureStylesPart, flush };
}

function findListNumIds(files) {
  const xml = readTextPart(files, "word/numbering.xml");
  if (!xml) return { bullet: "1", ordered: "2", exists: false };
  let bullet = "";
  let ordered = "";
  try {
    const doc = parseXml(xml);
    const absType = new Map();
    for (const abs of descendantsByLocal(doc, "abstractNum")) {
      const id = attrOf(abs, "abstractNumId");
      const lvl0 = childrenByLocal(abs, "lvl")[0];
      const fmt = lvl0 && firstChildByLocal(lvl0, "numFmt");
      const val = fmt ? attrOf(fmt, "val") : "bullet";
      absType.set(id, val === "bullet" || val === "none" ? "bullet" : "ordered");
    }
    for (const num of descendantsByLocal(doc, "num")) {
      const numId = attrOf(num, "numId");
      const absEl = firstChildByLocal(num, "abstractNumId");
      const t = absEl ? absType.get(attrOf(absEl, "val")) : null;
      if (t === "bullet" && !bullet) bullet = numId;
      if (t === "ordered" && !ordered) ordered = numId;
    }
  } catch {
    // fallback below
  }
  return { bullet: bullet || "1", ordered: ordered || "2", exists: !!(bullet && ordered) };
}

function serializeInline(nodes, relMgr, extraRPr = "") {
  let out = "";
  let imgId = 100;
  for (const n of nodes || []) {
    if (n.type === "hardBreak") {
      out += `<w:r><w:br/></w:r>`;
      continue;
    }
    if (n.type === "image") {
      const rId = relMgr.addImage(n.attrs || {});
      if (!rId) continue;
      const wPx = Number(n.attrs?.width) || 320;
      const hPx = Math.round(wPx * 0.66);
      const cx = Math.round(wPx * 9525);
      const cy = Math.round(hPx * 9525);
      const id = imgId++;
      const alt = escapeXml(n.attrs?.alt || "Image");
      out += `<w:r><w:drawing><wp:inline xmlns:wp="${WP_NS}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}" descr="${alt}"/><a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}"><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
      continue;
    }
    if (n.type !== "text") continue;
    const marks = n.marks || [];
    let rPr = extraRPr;
    let linkHref = "";
    for (const m of marks) {
      if (m.type === "bold") rPr += `<w:b/>`;
      else if (m.type === "italic") rPr += `<w:i/>`;
      else if (m.type === "underline") rPr += `<w:u w:val="single"/>`;
      else if (m.type === "strike") rPr += `<w:strike/>`;
      else if (m.type === "code") rPr += `<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/>`;
      else if (m.type === "textStyle") {
        const clr = String(m.attrs?.color || "").replace(/^#/, "");
        if (/^[0-9a-fA-F]{6}$/.test(clr)) rPr += `<w:color w:val="${clr.toUpperCase()}"/>`;
        const fs = String(m.attrs?.fontSize || "");
        const ptMatch = /^(\d+(?:\.\d+)?)pt$/i.exec(fs);
        if (ptMatch) rPr += `<w:sz w:val="${Math.round(Number(ptMatch[1]) * 2)}"/>`;
      } else if (m.type === "highlight") {
        const clr = String(m.attrs?.color || "#FEF08A").replace(/^#/, "");
        if (/^[0-9a-fA-F]{6}$/.test(clr)) rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${clr.toUpperCase()}"/>`;
      } else if (m.type === "link" && m.attrs?.href) {
        linkHref = String(m.attrs.href);
        if (!marks.some((x) => x.type === "textStyle" && x.attrs?.color)) {
          rPr += `<w:color w:val="2563EB"/><w:u w:val="single"/>`;
        }
      }
    }
    const parts = String(n.text || "").split("\t");
    let runInner = "";
    parts.forEach((seg, idx) => {
      if (idx > 0) runInner += `<w:tab/>`;
      if (seg.length > 0) {
        const preserve = /^\s|\s$|\s{2,}/.test(seg) ? ` xml:space="preserve"` : "";
        runInner += `<w:t${preserve}>${escapeXml(seg)}</w:t>`;
      }
    });
    const runXml = `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}${runInner}</w:r>`;
    if (linkHref && !linkHref.startsWith("#")) {
      const rId = relMgr.addHyperlink(linkHref);
      out += `<w:hyperlink r:id="${rId}">${runXml}</w:hyperlink>`;
    } else if (linkHref.startsWith("#")) {
      out += `<w:hyperlink w:anchor="${escapeXml(linkHref.slice(1))}">${runXml}</w:hyperlink>`;
    } else {
      out += runXml;
    }
  }
  return out;
}

function jcXml(align) {
  if (!align || align === "left") return "";
  if (align === "justify") return `<w:jc w:val="both"/>`;
  if (align === "center" || align === "right") return `<w:jc w:val="${align}"/>`;
  return "";
}

function serializeBlocks(blocks, ctx, listCtx = null) {
  let out = "";
  for (const b of blocks || []) {
    if (b.type === "paragraph") {
      let pPr = "";
      if (ctx.quoteMode) {
        pPr += `<w:pStyle w:val="Quote"/><w:ind w:left="720"/>`;
      }
      if (listCtx) {
        ctx.usedLists = true;
        const numId = listCtx.ordered ? ctx.numIds.ordered : ctx.numIds.bullet;
        pPr += `<w:numPr><w:ilvl w:val="${listCtx.level}"/><w:numId w:val="${numId}"/></w:numPr>`;
      }
      pPr += jcXml(b.attrs?.textAlign);
      const extraRPr = ctx.quoteMode ? `<w:i/>` : "";
      const runs = serializeInline(b.content, ctx.relMgr, extraRPr);
      out += `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${runs}</w:p>`;
    } else if (b.type === "heading") {
      const lvl = Math.min(4, Math.max(1, Number(b.attrs?.level) || 1));
      const halfPts = { 1: 44, 2: 34, 3: 28, 4: 24 }[lvl] || 28;
      const before = { 1: 240, 2: 200, 3: 160, 4: 120 }[lvl] || 160;
      const pPr = `<w:pStyle w:val="Heading${lvl}"/><w:spacing w:before="${before}" w:after="80"/><w:outlineLvl w:val="${lvl - 1}"/>${jcXml(b.attrs?.textAlign)}`;
      const runs = serializeInline(b.content, ctx.relMgr, `<w:b/><w:sz w:val="${halfPts}"/>`);
      out += `<w:p><w:pPr>${pPr}</w:pPr>${runs}</w:p>`;
    } else if (b.type === "bulletList" || b.type === "orderedList") {
      const ordered = b.type === "orderedList";
      const nextLevel = listCtx ? Math.min(3, listCtx.level + 1) : 0;
      for (const item of b.content || []) {
        out += serializeBlocks(item.content || [], ctx, { ordered, level: nextLevel });
      }
    } else if (b.type === "blockquote") {
      const prev = ctx.quoteMode;
      ctx.quoteMode = true;
      out += serializeBlocks(b.content || [], ctx, listCtx);
      ctx.quoteMode = prev;
    } else if (b.type === "codeBlock") {
      const text = (b.content || []).map((n) => n.text || "").join("");
      const lines = text.split("\n");
      for (const line of lines) {
        const preserve = /^\s|\s$|\s{2,}/.test(line) ? ` xml:space="preserve"` : "";
        out += `<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/><w:spacing w:before="20" w:after="20"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr>${line ? `<w:t${preserve}>${escapeXml(line)}</w:t>` : ""}</w:r></w:p>`;
      }
    } else if (b.type === "horizontalRule") {
      out += `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D0D5DD"/></w:pBdr></w:pPr></w:p>`;
    } else if (b.type === "image") {
      const runs = serializeInline([b], ctx.relMgr);
      if (runs) out += `<w:p>${runs}</w:p>`;
    } else if (b.type === "table") {
      out += serializeTable(b, ctx);
    }
  }
  return out;
}

function serializeTable(tblNode, ctx) {
  const rows = tblNode.content || [];
  if (!rows.length) return "";
  let trXml = "";
  let maxCols = 0;
  for (const row of rows) {
    const cells = row.content || [];
    let colCount = 0;
    let isHeader = false;
    let tcXml = "";
    for (const cell of cells) {
      const span = Math.max(1, Number(cell.attrs?.colspan) || 1);
      colCount += span;
      if (cell.type === "tableHeader") isHeader = true;
      let tcPr = `<w:tcW w:w="${Math.round(9000 / Math.max(1, cells.length)) * span}" w:type="dxa"/>`;
      if (span > 1) tcPr += `<w:gridSpan w:val="${span}"/>`;
      if (cell.type === "tableHeader") tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="F2F4F7"/>`;
      let cellInner = serializeBlocks(cell.content || [], ctx, null);
      if (!cellInner) cellInner = `<w:p/>`;
      tcXml += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${cellInner}</w:tc>`;
    }
    maxCols = Math.max(maxCols, colCount);
    trXml += `<w:tr>${isHeader ? `<w:trPr><w:tblHeader/></w:trPr>` : ""}${tcXml}</w:tr>`;
  }
  let gridXml = "";
  const colW = Math.round(9000 / Math.max(1, maxCols));
  for (let i = 0; i < maxCols; i++) gridXml += `<w:gridCol w:w="${colW}"/>`;
  const borders = `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:left w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:right w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/></w:tblBorders>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>${borders}<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${gridXml}</w:tblGrid>${trXml}</w:tbl>`;
}

export function saveDocx(session, tiptapDoc) {
  const files = { ...session.files };
  const relMgr = buildRelsManager(files);
  relMgr.ensureStylesPart();
  const numIds = findListNumIds(files);

  const ctx = {
    relMgr,
    numIds,
    usedLists: false,
    quoteMode: false,
  };
  let bodyInner = serializeBlocks(tiptapDoc?.content || [], ctx, null);
  if (!bodyInner) bodyInner = `<w:p/>`;
  if (ctx.usedLists && !numIds.exists) {
    writeTextPart(files, "word/numbering.xml", DEFAULT_NUMBERING_XML);
    ensureContentTypeOverride(
      files,
      "/word/numbering.xml",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
    );
    relMgr.ensureNumberingPart();
  } else if (ctx.usedLists) {
    relMgr.ensureNumberingPart();
  }

  const sectXml =
    session.sectPrXml ||
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">` +
    `<w:body>${bodyInner}${sectXml}</w:body></w:document>`;

  writeTextPart(files, "word/document.xml", documentXml);
  relMgr.flush();
  return zipPackage(files);
}

export function createBlankDocx(tiptapDoc = { type: "doc", content: [{ type: "paragraph" }] }) {
  const files = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
        `</Types>`
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="${PKG_REL_NS}">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
        `</Relationships>`
    ),
    "word/styles.xml": strToU8(DEFAULT_STYLES_XML),
  };
  return saveDocx({ files, sectPrXml: "" }, tiptapDoc);
}
