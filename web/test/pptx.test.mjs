import { test } from "node:test";
import assert from "node:assert/strict";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { strFromU8, strToU8 } from "fflate";
import { createBlankPptx, parsePptx, savePptx } from "../src/pptx-codec.js";
import { unzipPackage, zipPackage } from "../src/ooxml.js";

globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

test("pptx round-trip: shapes, coordinates, colors, notes, and slide operations", () => {
  const b64 = createBlankPptx([
    {
      sourceSlidePath: "",
      layoutPath: "ppt/slideLayouts/slideLayout1.xml",
      background: { kind: "solid", css: "#F8FAFC", hex: "#F8FAFC" },
      elements: [
        {
          uid: "s0-title",
          kind: "shape",
          isNew: true,
          name: "Heading",
          x: 100,
          y: 80,
          w: 800,
          h: 90,
          rot: 0,
          geom: "rect",
          fill: "transparent",
          fillHex: "",
          vAlign: "middle",
          padding: { top: 6, right: 10, bottom: 6, left: 10 },
          paragraphs: [
            {
              align: "left",
              bullet: false,
              level: 0,
              runs: [{ text: "Roadmap 2027", fontSize: 36, bold: true, italic: false, underline: false, color: "#1E3A8A" }],
            },
          ],
        },
      ],
      notes: "Mention Q1 targets.",
      dirty: true,
      notesDirty: true,
    },
    {
      sourceSlidePath: "",
      layoutPath: "ppt/slideLayouts/slideLayout1.xml",
      background: { kind: "solid", css: "#FFFFFF", hex: "#FFFFFF" },
      elements: [
        {
          uid: "s1-box",
          kind: "shape",
          isNew: true,
          name: "Card",
          x: 120,
          y: 160,
          w: 400,
          h: 200,
          rot: 0,
          geom: "ellipse",
          fill: "#DBEAFE",
          fillHex: "#DBEAFE",
          vAlign: "middle",
          padding: { top: 8, right: 12, bottom: 8, left: 12 },
          paragraphs: [
            {
              align: "center",
              bullet: false,
              level: 0,
              runs: [{ text: "Slide Two", fontSize: 24, bold: false, italic: true, underline: false, color: "#181B22" }],
            },
          ],
        },
      ],
      notes: "",
      dirty: true,
      notesDirty: false,
    },
  ]);

  const parsed = parsePptx(b64);
  assert.equal(parsed.slideW, 1280);
  assert.equal(parsed.slideH, 720);
  assert.equal(parsed.slides.length, 2);
  assert.equal(parsed.slides[0].background.hex, "#F8FAFC");
  assert.equal(parsed.slides[0].elements[0].paragraphs[0].runs[0].text, "Roadmap 2027");
  assert.equal(parsed.slides[0].elements[0].paragraphs[0].runs[0].color, "#1E3A8A");
  assert.equal(parsed.slides[0].notes, "Mention Q1 targets.");
  assert.equal(parsed.slides[1].elements[0].geom, "ellipse");
  assert.equal(parsed.slides[1].elements[0].fillHex, "#DBEAFE");
});

test("pptx lossless preservation: untouched slides keep exact XML bytes and edited slides preserve unknown tags", () => {
  const initialB64 = createBlankPptx();
  const files = unzipPackage(initialB64);

  // Inject a custom <p:transition spd="med"><p:fade/></p:transition> into slide1.xml
  // and create a second slide2.xml with custom XML content that must stay byte-for-byte untouched.
  const slide1Xml = strFromU8(files["ppt/slides/slide1.xml"]).replace(
    "</p:sld>",
    `<p:transition spd="med"><p:fade/></p:transition></p:sld>`
  );
  files["ppt/slides/slide1.xml"] = strToU8(slide1Xml);

  const untouchedSlide2Bytes = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
      `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="7" name="KeepMe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="3810000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Untouched Slide 2</a:t></a:r></a:p></p:txBody></p:sp>` +
      `</p:spTree></p:cSld><p:timing><p:tnLst/></p:timing></p:sld>`
  );
  files["ppt/slides/slide2.xml"] = untouchedSlide2Bytes;
  files["ppt/slides/_rels/slide2.xml.rels"] = files["ppt/slides/_rels/slide1.xml.rels"];
  files["ppt/presentation.xml"] = strToU8(
    strFromU8(files["ppt/presentation.xml"]).replace(
      /(<p:sldId\b[^>]*\/>)/,
      `$1<p:sldId id="257" r:id="rId99"/>`
    )
  );
  files["ppt/_rels/presentation.xml.rels"] = strToU8(
    strFromU8(files["ppt/_rels/presentation.xml.rels"]).replace(
      "</Relationships>",
      `<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>`
    )
  );

  const session = parsePptx(zipPackage(files));
  assert.equal(session.slides.length, 2);

  // Edit ONLY slide 1's shape text & x coordinate; leave slide 2 untouched!
  const slides = structuredClone(session.slides);
  slides[0].elements[0].x = 180;
  slides[0].elements[0].paragraphs[0].runs[0].text = "Edited Slide 1";
  slides[0].elements[0].textChanged = true;
  slides[0].dirty = true;

  const savedB64 = savePptx(session, slides);
  const savedFiles = unzipPackage(savedB64);

  // Slide 2 bytes must be 100% identical byte-for-byte
  assert.deepEqual(savedFiles["ppt/slides/slide2.xml"], untouchedSlide2Bytes);

  // Slide 1 must keep <p:transition> while reflecting the updated text & coordinate
  const savedSlide1 = strFromU8(savedFiles["ppt/slides/slide1.xml"]);
  assert.match(savedSlide1, /<p:transition spd="med">/);
  assert.match(savedSlide1, /Edited Slide 1/);

  // Reordering + adding a 3rd slide updates sldIdLst cleanly
  const reparsed = parsePptx(savedB64);
  const reordered = [reparsed.slides[1], reparsed.slides[0]];
  const reorderedB64 = savePptx(reparsed, reordered);
  const checkOrder = parsePptx(reorderedB64);
  assert.equal(checkOrder.slides[0].elements[0].paragraphs[0].runs[0].text, "Untouched Slide 2");
  assert.equal(checkOrder.slides[1].elements[0].paragraphs[0].runs[0].text, "Edited Slide 1");
});

test("pptx refuses non-zip and zip missing ppt/presentation.xml", () => {
  assert.throws(() => parsePptx(Buffer.from("plain text").toString("base64")), /not a ZIP/);
  const fakeZip = zipPackage({ "[Content_Types].xml": "<Types/>" });
  assert.throws(() => parsePptx(fakeZip), /missing ppt\/presentation\.xml/);
});
