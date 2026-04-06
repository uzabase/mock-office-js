# E2E Testing Design for mock-office-js

## Goal

Verify that mock-office-js works correctly in a real browser environment, loaded via `<script type="module">` in an Excel Add-in application structure. No GUI testing — the focus is on confirming the mock API functions properly when used as a drop-in replacement for Office.js.

## Test Directory Restructuring

Split `tests/` into `tests/unit/` and `tests/e2e/`:

```
tests/
├── unit/                          # Existing tests moved here
│   ├── address.test.ts
│   ├── cell-storage.test.ts
│   ├── custom-functions-mock.test.ts
│   ├── custom-functions.test-d.ts
│   ├── excel-mock.test.ts
│   ├── formula-evaluator.test.ts
│   ├── formula-parser.test.ts
│   ├── integration.test.ts
│   ├── range.test.ts
│   ├── range.test-d.ts
│   ├── request-context.test.ts
│   ├── workbook.test-d.ts
│   └── workbook.test.ts
└── e2e/
    ├── fixture/                   # yo office generated Excel Custom Functions Add-in
    │   ├── manifest.xml
    │   ├── src/
    │   │   ├── taskpane/
    │   │   │   ├── taskpane.html
    │   │   │   └── taskpane.ts
    │   │   └── functions/
    │   │       ├── functions.ts
    │   │       └── functions.json
    │   ├── webpack.config.js
    │   └── package.json
    ├── excel-mock.e2e.test.ts
    └── playwright.config.ts
```

## Fixture App

### Approach

1. Generate an Excel Custom Functions Add-in using `yo office` template
2. Modify the webpack config to resolve `mock-office-js` to `../../../src/index.ts` (avoids dependency on built `dist/`)
3. Replace the Office.js CDN `<script>` tag in HTML with a `<script>` tag that loads mock-office-js
4. Confirm E2E tests pass, then trim unnecessary generated files

### Why yo office

- Produces a real Add-in structure with manifest.xml, functions.json, and function implementations
- Safer to start from a working template and remove what's not needed than to hand-craft and miss something
- webpack config and other Office-specific setup come pre-configured

## Mock Global Wiring

The fixture app's entry point must instantiate `ExcelMock` and expose it as browser globals, replacing the real Office.js:

```ts
// In the fixture's bootstrap code (e.g., taskpane.ts)
import { ExcelMock } from "mock-office-js";

const mock = new ExcelMock();
(window as any).Excel = mock.excel;
(window as any).CustomFunctions = mock.customFunctions;
// Expose mock instance for test-side setup (e.g., setCell, reset)
(window as any).__mock__ = mock;
```

The `__mock__` global allows E2E tests to call helper methods like `mock.setCell()` and `mock.reset()` via `page.evaluate()`.

## E2E Test Architecture

### Flow

1. Playwright `webServer` option auto-starts the fixture app's dev server
2. Playwright opens a browser and navigates to taskpane.html
3. Tests use `page.evaluate()` to interact with the global `Excel` / `CustomFunctions` / `__mock__` objects in the browser
4. Assertions verify mock behavior

### Test Scenarios

- `Excel.run()` — get and set cell values via RequestContext
- `CustomFunctions.associate()` — register custom functions and evaluate formulas
- Worksheet operations — add worksheets, switch active worksheet
- Load/sync pattern — verify properties require load + sync before access

### Example Test

```ts
test("Excel.run can read cell values set via mock", async ({ page }) => {
  await page.goto("/taskpane.html");

  const value = await page.evaluate(async () => {
    const mock = (window as any).__mock__;
    await mock.setCell("Sheet1", "A1", { value: 42 });

    let result: number[][] = [];
    await (window as any).Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("A1");
      range.load("values");
      await context.sync();
      result = range.values;
    });
    return result;
  });

  expect(value).toEqual([[42]]);
});
```

## Configuration Changes

### vitest.config.ts

- Change test include path to `tests/unit/` only

### playwright.config.ts (new, in `tests/e2e/`)

- `webServer`: start fixture app's dev server
- `testDir`: `tests/e2e/`
- `testMatch`: `*.e2e.test.ts`

### Fixture webpack config

- Add `resolve.alias` to point `mock-office-js` to source TypeScript (`../../../src/index.ts`)
- Note: the alias target is outside the fixture directory. The TypeScript loader's `include`/`configFile` settings may need adjustment to process files outside the fixture root. Exact config adjustments determined at implementation time based on generated template.

### Fixture dependency installation

- The fixture has its own `package.json` with its own dependencies (webpack, loaders, dev server, etc.)
- Add a `pretest:e2e` script to the root `package.json` that runs `npm install` in the fixture directory
- Commit the fixture's `package-lock.json` for reproducible installs

### package.json

- Add `@playwright/test` to devDependencies
- Add `test:e2e` script: `playwright test --config tests/e2e/playwright.config.ts`
- Add `pretest:e2e` script: `cd tests/e2e/fixture && npm install`

### .gitignore

- Add `tests/e2e/fixture/node_modules/`
- Add `test-results/`
- Add `playwright-report/`

### npm scripts (updated)

- `test` → `vitest run` (target: `tests/unit/`)
- `test:e2e` → `playwright test` (with `pretest:e2e` auto-running fixture install)
- `test:typecheck` → unchanged
