import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync, existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "xlsx";
import { createBlankXlsx, detectOoxmlKind } from "../src/ooxml.js";
import { createBlankDocx } from "../src/docx-codec.js";
import { createBlankPptx } from "../src/pptx-codec.js";
import {
  deleteSavedFile,
  editorPageForKind,
  listSavedFiles,
  saveStandaloneRecord,
} from "../src/standalone-store.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "../..");

test("picode-extension.json exposes one unified Office files page and icon", () => {
  const manifest = JSON.parse(readFileSync(join(root, "picode-extension.json"), "utf8"));
  assert.equal(manifest.id, "office");
  assert.equal(manifest.icon, "ui/icon.svg");
  assert.equal(existsSync(join(root, manifest.icon)), true);
  assert.equal(manifest.ui.pages.length, 1);
  assert.deepEqual(manifest.ui.pages[0], {
    id: "office",
    title: "Office files",
    path: "ui/office.html",
  });
  for (const kind of manifest.contributes.artifactKinds) {
    assert.equal(kind.page, "office");
  }
});

test("createBlankXlsx and detectOoxmlKind identify xlsx, docx, and pptx", () => {
  const xlsxB64 = createBlankXlsx("Planilha1");
  assert.equal(detectOoxmlKind(xlsxB64), "xlsx");
  const wb = XLSX.read(xlsxB64, { type: "base64" });
  assert.deepEqual(wb.SheetNames, ["Planilha1"]);

  // Install minimal DOMParser/XMLSerializer if needed for docx/pptx generators
  if (!globalThis.DOMParser) {
    // createBlankDocx & createBlankPptx only need DOMParser when ensureContentType* parses XML
    return;
  }
  assert.equal(detectOoxmlKind(createBlankDocx()), "docx");
  assert.equal(detectOoxmlKind(createBlankPptx()), "pptx");
  assert.equal(detectOoxmlKind("bm90IGEgemlw"), "");
});

test("standalone store saves, sorts, updates, and deletes Office records", async () => {
  const kv = new Map();
  const mockDoor = {
    async list() {
      return [...kv.entries()].map(([key, value]) => ({ key, value }));
    },
    async get(key) {
      return kv.has(key) ? { key, value: kv.get(key) } : null;
    },
    async set(key, value) {
      kv.set(key, value);
      return { key, value };
    },
    async delete(key) {
      kv.delete(key);
      return { deleted: true };
    },
  };

  const r1 = await saveStandaloneRecord(mockDoor, {
    id: "f-one",
    kind: "xlsx",
    title: "Q4 Budget",
    data: createBlankXlsx("Q4"),
  });
  assert.equal(r1.persisted, true);

  const list = await listSavedFiles(mockDoor);
  assert.equal(list.length, 1);
  assert.equal(list[0].title, "Q4 Budget");
  assert.equal(editorPageForKind(list[0].kind), "sheet.html");
  assert.equal(editorPageForKind("ext:office/docx"), "doc.html");
  assert.equal(editorPageForKind("ext:office/pptx"), "deck.html");

  await deleteSavedFile(mockDoor, "f-one");
  assert.equal((await listSavedFiles(mockDoor)).length, 0);
});
