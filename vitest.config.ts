import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["tests/unit/**/*.test.ts"],
    exclude: [".references/**", ".claude/**", "node_modules/**"],
    passWithNoTests: true,
    typecheck: {
      tsconfig: "./tsconfig.test-d.json",
      include: ["tests/unit/**/*.test-d.ts"],
    },
  },
});
