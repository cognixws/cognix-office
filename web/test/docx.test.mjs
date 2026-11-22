import { test } from "node:test";
import assert from "node:assert/strict";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { strFromU8, strToU8 } from "fflate";
import { createBlankDocx, parseDocx, saveDocx } from "../src/docx-codec.js";
import { unzipPackage, zipPackage } from "../src/ooxml.js";

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const TINY_PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ioAAAAASUVORK5CYII=";

test("docx round-trip: headings, rich marks, lists, tables, links, and images", () => {
  const initialDoc = {
    type: "doc",
    content: [
      {
        type: "heading",
        attrs: { level: 1, textAlign: "center" },
        content: [{ type: "text", text: "Quarterly Review" }],
      },
      {
        type: "paragraph",
        content: [
          { type: "text", text: "Status: " },
          { type: "text", text: "on track", marks: [{ type: "bold" }, { type: "underline" }, { type: "textStyle", attrs: { color: "#15803D" } }] },
          { type: "text", text: " — see " },
          { type: "text", text: "report", marks: [{ type: "link", attrs: { href: "https://example.com/q4" } }] },
        ],
      },
      {
        type: "bulletList",
        content: [
          { type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "First milestone" }] }] },
          { type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "Second milestone" }] }] },
        ],
      },
      {
        type: "orderedList",
        content: [
          { type: "listItem", content: [{ type: "paragraph", content: [{ type: "text", text: "Step one" }] }] },
        ],
      },
      {
        type: "table",
        content: [
          {
            type: "tableRow",
            content: [
              { type: "tableHeader", attrs: { colspan: 1, rowspan: 1 }, content: [{ type: "paragraph", content: [{ type: "text", text: "Area" }] }] },
              { type: "tableHeader", attrs: { colspan: 1, rowspan: 1 }, content: [{ type: "paragraph", content: [{ type: "text", text: "Owner" }] }] },
            ],
          },
          {
            type: "tableRow",
            content: [
              { type: "tableCell", attrs: { colspan: 1, rowspan: 1 }, content: [{ type: "paragraph", content: [{ type: "text", text: "Core" }] }] },
              { type: "tableCell", attrs: { colspan: 1, rowspan: 1 }, content: [{ type: "paragraph", content: [{ type: "text", text: "Ada" }] }] },
            ],
          },
        ],
      },
      {
        type: "paragraph",
        content: [
          { type: "image", attrs: { src: `data:image/png;base64,${TINY_PNG_B64}`, alt: "Dot", width: 64 } },
        ],
      },
    ],
  };

  const b64 = createBlankDocx(initialDoc);
  const parsed = parseDocx(b64);

  assert.equal(parsed.doc.content[0].type, "heading");
  assert.equal(parsed.doc.content[0].attrs.level, 1);
  assert.equal(parsed.doc.content[0].attrs.textAlign, "center");
  assert.equal(parsed.doc.content[0].content[0].text, "Quarterly Review");

  const paraRuns = parsed.doc.content[1].content;
  const coloredRun = paraRuns.find((r) => r.text === "on track");
  assert.ok(coloredRun);
  assert.ok(coloredRun.marks.some((m) => m.type === "bold"));
  assert.ok(coloredRun.marks.some((m) => m.type === "underline"));
  assert.ok(coloredRun.marks.some((m) => m.type === "textStyle" && m.attrs.color === "#15803D"));

  const linkRun = paraRuns.find((r) => r.text === "report");
  assert.ok(linkRun.marks.some((m) => m.type === "link" && m.attrs.href === "https://example.com/q4"));

  assert.equal(parsed.doc.content[2].type, "bulletList");
  assert.equal(parsed.doc.content[2].content.length, 2);
  assert.equal(parsed.doc.content[3].type, "orderedList");
  assert.equal(parsed.doc.content[4].type, "table");
  assert.equal(parsed.doc.content[4].content[1].content[1].content[0].content[0].text, "Ada");

  const imgNode = parsed.doc.content[5].content[0];
  assert.equal(imgNode.type, "image");
  assert.match(imgNode.attrs.src, /^data:image\/png;base64,/);
});

test("docx preserves untouched package parts and sectPr across edits", () => {
  const baseB64 = createBlankDocx({
    type: "doc",
    content: [{ type: "paragraph", content: [{ type: "text", text: "Original text" }] }],
  });
  const files = unzipPackage(baseB64);
  files["word/header1.xml"] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Confidential Header</w:t></w:r></w:p></w:hdr>`);
  files["docProps/core.xml"] = strToU8(`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">Author X</dc:creator></cp:coreProperties>`);
  const withExtras = zipPackage(files);

  const session = parseDocx(withExtras);
  session.doc.content[0].content[0].text = "Updated text";
  const savedB64 = saveDocx(session, session.doc);

  const outFiles = unzipPackage(savedB64);
  assert.match(strFromU8(outFiles["word/header1.xml"]), /Confidential Header/);
  assert.match(strFromU8(outFiles["docProps/core.xml"]), /Author X/);

  const reparsed = parseDocx(savedB64);
  assert.equal(reparsed.doc.content[0].content[0].text, "Updated text");
});

test("docx refuses non-zip and zip missing word/document.xml", () => {
  assert.throws(() => parseDocx(Buffer.from("not a zip").toString("base64")), /not a ZIP/);
  const fakeZip = zipPackage({ "[Content_Types].xml": "<Types/>" });
  assert.throws(() => parseDocx(fakeZip), /missing word\/document\.xml/);
});
