import { defineConfig } from "vite";
export default defineConfig({
  base: "./",
  define: { "process.env.NODE_ENV": JSON.stringify("production") },
  build: {
    lib: { entry: "src/main.js", name: "OfficeSheet", formats: ["iife"], fileName: () => "office.js", cssFileName: "office" },
    cssCodeSplit: false,
    outDir: "../ui",
    emptyOutDir: false,
    minify: true,
    assetsInlineLimit: 0,
    target: "es2020",
  },
});
