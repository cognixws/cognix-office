import { test } from "node:test";
import assert from "node:assert/strict";
import { LANGS, STRINGS, localeOf, numfmtOf, resolveLang, stringsFor } from "../src/i18n.js";

test("the browser's language is the default", () => {
  assert.equal(resolveLang("", "pt-BR"), "pt-BR");
  assert.equal(resolveLang("", "pt-PT"), "pt-BR");
  assert.equal(resolveLang("", "en-US"), "en-US");
  assert.equal(resolveLang("", "fr-FR"), "en-US");
  assert.equal(resolveLang(null, ""), "en-US");
});

test("a known saved choice wins, an unknown one does not", () => {
  assert.equal(resolveLang("en-US", "pt-BR"), "en-US");
  assert.equal(resolveLang("pt-BR", "en-US"), "pt-BR");
  assert.equal(resolveLang("de-DE", "pt-BR"), "pt-BR");
});

test("the engine's spellings follow the language", () => {
  assert.deepEqual(LANGS.map(localeOf), ["ptBR", "enUS"]);
  assert.deepEqual(LANGS.map(numfmtOf), ["pt_BR", "en_US"]);
});

test("both languages carry every string", () => {
  const keys = Object.keys(STRINGS["pt-BR"]).sort();
  assert.deepEqual(Object.keys(STRINGS["en-US"]).sort(), keys);
  for (const lang of LANGS) for (const [k, v] of Object.entries(STRINGS[lang])) assert.ok(typeof v === "function" || (typeof v === "string" && v.length > 0), lang + " " + k);
  assert.equal(stringsFor("pt-BR").saved(3), "Salvo · v3");
  assert.equal(stringsFor("en-US").saved(3), "Saved · v3");
  assert.match(stringsFor("pt-BR").conflict(4), /versão 4/);
  assert.match(stringsFor("en-US").conflict(4), /version 4/);
});
