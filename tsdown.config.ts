import { defineConfig } from "tsdown";

export default defineConfig([
  {
    entry: { office: "src/index.ts" },
    format: ["esm"],
    dts: { tsconfig: "./tsconfig.build.json" },
    clean: true,
    outDir: "dist",
  },
  {
    entry: { office: "src/browser.ts" },
    format: ["iife"],
    outDir: "dist",
  },
]);
