// Standalone Office files persistence through picode.use("state"), with an
// in-tab window.name fallback when a local file exceeds ExtensionKVValueMax
// (256 KiB per key) or ExtensionKVQuota (1 MiB per extension).

const KEY_PREFIX = "file/";
const MAX_JSON_BYTES = 240 * 1024;
const TRANSIENT_TAG = "__cognixOfficeTransient";

function utf8Length(str) {
  if (typeof TextEncoder !== "undefined") {
    return new TextEncoder().encode(str).byteLength;
  }
  return str.length * 2;
}

export function makeFileId() {
  const rand = Math.random().toString(36).slice(2, 8);
  return "f-" + Date.now().toString(36) + "-" + rand;
}

export function approxBytesOfBase64(b64) {
  const len = String(b64 || "").replace(/=+$/, "").length;
  return Math.max(1, Math.floor((len * 3) / 4));
}

export function editorPageForKind(kind) {
  if (kind === "xlsx" || kind === "ext:office/xlsx") return "sheet.html";
  if (kind === "docx" || kind === "ext:office/docx") return "doc.html";
  if (kind === "pptx" || kind === "ext:office/pptx") return "deck.html";
  return "";
}

export function writeTransientRecord(record) {
  try {
    if (typeof window !== "undefined") {
      window.name = JSON.stringify({ [TRANSIENT_TAG]: record });
      return true;
    }
  } catch {
    // ignore
  }
  return false;
}

export function readTransientRecord(id) {
  try {
    if (typeof window === "undefined" || !window.name) return null;
    const parsed = JSON.parse(window.name);
    const rec = parsed && parsed[TRANSIENT_TAG];
    if (rec && (!id || rec.id === id) && rec.data) return rec;
  } catch {
    // ignore
  }
  return null;
}

export async function getStateDoor() {
  try {
    if (typeof picode !== "undefined" && picode && typeof picode.use === "function") {
      return await picode.use("state");
    }
  } catch {
    // standalone browser or test
  }
  return null;
}

export async function listSavedFiles(stateDoor) {
  if (!stateDoor || typeof stateDoor.list !== "function") return [];
  let rows = [];
  try {
    rows = (await stateDoor.list()) || [];
  } catch {
    return [];
  }
  const out = [];
  for (const row of rows) {
    if (!row || typeof row.key !== "string" || !row.key.startsWith(KEY_PREFIX)) continue;
    const val = row.value;
    if (!val || typeof val !== "object" || !val.id || !val.kind || !val.data) continue;
    out.push({
      id: String(val.id),
      kind: String(val.kind),
      title: String(val.title || ""),
      data: String(val.data),
      sizeBytes: Number(val.sizeBytes) || approxBytesOfBase64(val.data),
      updatedAt: String(val.updatedAt || ""),
    });
  }
  out.sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
  return out;
}

export async function saveStandaloneRecord(stateDoor, record) {
  const clean = {
    id: String(record.id || makeFileId()),
    kind: String(record.kind),
    title: String(record.title || "").trim() || "Untitled",
    data: String(record.data || ""),
    sizeBytes: approxBytesOfBase64(record.data),
    updatedAt: new Date().toISOString(),
  };
  const serialized = JSON.stringify(clean);
  if (stateDoor && typeof stateDoor.set === "function" && utf8Length(serialized) <= MAX_JSON_BYTES) {
    try {
      await stateDoor.set(KEY_PREFIX + clean.id, clean);
      return { record: clean, persisted: true };
    } catch {
      // Quota exceeded or transient failure: fall back to window.name
    }
  }
  writeTransientRecord(clean);
  return { record: clean, persisted: false };
}

export async function deleteSavedFile(stateDoor, id) {
  if (!stateDoor || typeof stateDoor.delete !== "function") return;
  try {
    await stateDoor.delete(KEY_PREFIX + id);
  } catch {
    // already absent
  }
}

export function navigateFrame(target) {
  if (typeof location !== "undefined") {
    location.replace(target);
  }
}

export async function openStandaloneDraft(expectedKind) {
  if (typeof location === "undefined") return null;
  const params = new URLSearchParams(location.search || "");
  const fileId = params.get("file") || "";
  const transientId = params.get("transient") || "";
  if (!fileId && !transientId) return null;

  const stateDoor = await getStateDoor();
  let record = null;
  let persisted = false;

  if (fileId && stateDoor && typeof stateDoor.get === "function") {
    try {
      const entry = await stateDoor.get(KEY_PREFIX + fileId);
      if (entry && entry.value && typeof entry.value === "object" && entry.value.data) {
        record = entry.value;
        persisted = true;
      }
    } catch {
      // check transient fallback
    }
  }
  if (!record) {
    record = readTransientRecord(fileId || transientId);
    persisted = false;
  }
  if (!record || (expectedKind && record.kind !== expectedKind)) {
    return null;
  }

  return {
    id: record.id,
    kind: record.kind,
    title: record.title || "",
    data: record.data,
    persisted,
    async save(nextBase64, nextTitle) {
      const updated = {
        ...record,
        title: nextTitle != null ? String(nextTitle).trim() || record.title : record.title,
        data: nextBase64 != null ? nextBase64 : record.data,
      };
      const res = await saveStandaloneRecord(stateDoor, updated);
      record = res.record;
      persisted = res.persisted;
      return res;
    },
    back() {
      navigateFrame("office.html");
    },
  };
}
