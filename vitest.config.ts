import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    exclude: [".references/**", ".claude/**", "node_modules/**"],
    passWithNoTests: true,
    typecheck: {
      tsconfig: "./tsconfig.test-d.json",
    },
  },
});
