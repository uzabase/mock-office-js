import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: ".",
  testMatch: "*.e2e.test.ts",
  webServer: {
    command: "npx webpack serve --mode development",
    cwd: "./fixture",
    url: "http://localhost:3000/taskpane.html",
    reuseExistingServer: !process.env.CI,
  },
  use: {
    baseURL: "http://localhost:3000",
  },
});
