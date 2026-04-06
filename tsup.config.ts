import { defineConfig } from "tsup";

export default defineConfig([
  {
    entry: { office: "src/index.ts" },
    format: ["esm"],
    dts: false,
    clean: true,
    outDir: "dist",
    outExtension: () => ({ js: ".esm.js" }),
  },
  {
    entry: { office: "src/browser.ts" },
    format: ["iife"],
    outDir: "dist",
    outExtension: () => ({ js: ".js" }),
  },
]);
