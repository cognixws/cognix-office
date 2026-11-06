// The page's own language (ADR-0230 door A page): the browser's language is
// the default, a saved choice wins, and the number-format locale follows it.
// Univer's UI strings come from its locale bundles; these are the strings the
// page draws itself (the bar, the save state, the conflict line, failures).

export const LANGS = ["pt-BR", "en-US"];

// resolveLang answers the language for a saved choice and a browser tag: a
// known saved choice wins; otherwise the browser's language, Portuguese for
// any pt-* tag and English for everything else.
export function resolveLang(saved, browser) {
  if (LANGS.includes(saved)) return saved;
  return String(browser || "").toLowerCase().startsWith("pt") ? "pt-BR" : "en-US";
}

// localeOf and numfmtOf are the engine's spellings of the same choice.
export const localeOf = (lang) => (lang === "pt-BR" ? "ptBR" : "enUS");
export const numfmtOf = (lang) => (lang === "pt-BR" ? "pt_BR" : "en_US");

export const STRINGS = {
  "pt-BR": {
    language: "Idioma",
    title: "Planilha",
    saved: (n) => "Salvo · v" + n,
    saving: "Salvando…",
    unsaved: "Não salvo · tente de novo",
    readOnly: "Somente leitura",
    conflict: (n) => "Um agente publicou a versão " + n + " enquanto você editava. Sua edição está aqui.",
    conflictSave: (n) => "Não salvo · v" + n + " chegou",
    open: "Abrir a nova versão",
    keep: "Manter a minha",
    retry: "Recarregar",
    notStandalone: "Esta página desenha uma versão de artefato. Abra pelo artefato, não sozinha.",
    unreadable: "Esta planilha não pôde ser lida. O arquivo continua nas versões do artefato.",
    notWorkbook: "Este arquivo não é uma planilha que o COGNIX[WS] consiga desenhar.",
    noSheet: "Esta planilha não tem nenhuma aba.",
    undrawable: "Esta planilha não pôde ser desenhada. O arquivo continua nas versões do artefato.",
  },
  "en-US": {
    language: "Language",
    title: "Spreadsheet",
    saved: (n) => "Saved · v" + n,
    saving: "Saving…",
    unsaved: "Not saved · try again",
    readOnly: "Read-only",
    conflict: (n) => "An agent published version " + n + " while you were editing. Your edit is here.",
    conflictSave: (n) => "Not saved · v" + n + " arrived",
    open: "Open the new version",
    keep: "Keep mine",
    retry: "Retry",
    notStandalone: "This page draws an artifact version. Open it from the artifact, not on its own.",
    unreadable: "This workbook could not be read. Its file is still in the artifact's versions.",
    notWorkbook: "This file is not a workbook COGNIX[WS] can draw.",
    noSheet: "This workbook holds no sheet.",
    undrawable: "This workbook could not be drawn. Its file is still in the artifact's versions.",
  },
};

export const stringsFor = (lang) => STRINGS[lang] || STRINGS["en-US"];
