import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: ".",
  testMatch: "*.e2e.test.ts",
  webServer: {
    command: "npx webpack serve --mode development",
    cwd: "./fixture",
    url: "https://localhost:3000/taskpane.html",
    ignoreHTTPSErrors: true,
    reuseExistingServer: !process.env.CI,
  },
  use: {
    baseURL: "https://localhost:3000",
    ignoreHTTPSErrors: true,
  },
});
