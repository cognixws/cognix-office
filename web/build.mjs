import { build } from "vite";

const entries = [
  { entry: "src/main.js", name: "OfficeSheet", base: "office" },
  { entry: "src/doc-main.js", name: "OfficeDoc", base: "doc" },
  { entry: "src/deck-main.js", name: "OfficeDeck", base: "deck" },
];

for (const item of entries) {
  await build({
    configFile: false,
    base: "./",
    define: { "process.env.NODE_ENV": JSON.stringify("production") },
    build: {
      lib: {
        entry: item.entry,
        name: item.name,
        formats: ["iife"],
        fileName: () => `${item.base}.js`,
        cssFileName: item.base,
      },
      cssCodeSplit: false,
      outDir: "../ui",
      emptyOutDir: false,
      minify: true,
      assetsInlineLimit: 0,
      target: "es2020",
    },
  });
}
