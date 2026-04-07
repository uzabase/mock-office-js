# Global API Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Redesign mock-office-js so that `import "mock-office-js"` automatically sets up `Excel`, `Office`, `CustomFunctions`, and `MockOfficeJs` as globals, matching real office.js loading behavior.

**Architecture:** Replace `ExcelMock` class with a `createMockEnvironment()` factory function in `src/setup.ts` that constructs shared state and returns all global objects. `src/index.ts` becomes a side-effect-only module that calls the factory and assigns globals. Test helpers live under `MockOfficeJs.excel.*`.

**Tech Stack:** TypeScript, tsdown (ESM + IIFE), vitest, Playwright

**Spec:** `docs/superpowers/specs/2026-04-08-global-api-redesign.md`

---

## File Structure

| File | Action | Responsibility |
|---|---|---|
| `src/setup.ts` | Create | Factory function `createMockEnvironment()` — constructs shared state, returns `{ excel, office, customFunctions, mockOfficeJs }` |
| `src/globals.d.ts` | Create | `declare global` for `MockOfficeJs` type |
| `src/index.ts` | Rewrite | Side-effect-only — imports factory, assigns globals |
| `src/browser.ts` | Delete | No longer needed (index.ts handles both ESM and IIFE) |
| `src/excel-mock.ts` | Delete | Replaced by `setup.ts` |
| `tsdown.config.ts` | Modify | Both ESM and IIFE use `src/index.ts` as entry |
| `package.json` | Modify | Add `peerDependencies`, update build script |
| `tests/unit/excel-mock.test.ts` | Rewrite → rename to `tests/unit/mock-office-js.test.ts` | Test `MockOfficeJs.excel.*` and `MockOfficeJs.reset()` |
| `tests/unit/integration.test.ts` | Rewrite | Use globals instead of `ExcelMock` instance |
| `tests/e2e/excel-mock.e2e.test.ts` | Rewrite → rename to `tests/e2e/mock-office-js.e2e.test.ts` | Migrate from `window.__mock__` to `window.MockOfficeJs` |

Core layer files (`cell-storage.ts`, `custom-functions-mock.ts`, `request-context.ts`, `range.ts`, `workbook.ts`, `worksheet.ts`, `worksheet-collection.ts`, `formula-parser.ts`, `formula-evaluator.ts`, `address.ts`) are **unchanged**.

Type test files (`tests/unit/range.test-d.ts`, `tests/unit/workbook.test-d.ts`, `tests/unit/custom-functions.test-d.ts`) are **unchanged** — they import from core modules, not `ExcelMock`.

---

### Task 1: Create `setup.ts` with `createMockEnvironment()`

**Files:**
- Create: `src/setup.ts`
- Test: `tests/unit/setup.test.ts`

This task builds the factory function that replaces `ExcelMock`. The factory constructs shared state and returns four objects for global assignment. All logic is extracted from the existing `ExcelMock` class (`src/excel-mock.ts`).

- [ ] **Step 1: Write failing test for `createMockEnvironment` return shape**

Create `tests/unit/setup.test.ts`:

```ts
import { describe, expect, test } from "vitest";
import { createMockEnvironment } from "../../src/setup.js";

describe("createMockEnvironment", () => {
  test("returns excel, office, customFunctions, and mockOfficeJs", () => {
    const env = createMockEnvironment();
    expect(env.excel).toBeDefined();
    expect(env.excel.run).toBeTypeOf("function");
    expect(env.office).toBeDefined();
    expect(env.office.onReady).toBeTypeOf("function");
    expect(env.customFunctions).toBeDefined();
    expect(env.customFunctions.associate).toBeTypeOf("function");
    expect(env.mockOfficeJs).toBeDefined();
    expect(env.mockOfficeJs.excel).toBeDefined();
    expect(env.mockOfficeJs.reset).toBeTypeOf("function");
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/setup.test.ts`
Expected: FAIL — `Cannot find module '../../src/setup.js'`

- [ ] **Step 3: Write minimal `setup.ts`**

Create `src/setup.ts`. Extract all logic from `ExcelMock` class into the factory function:

```ts
import { CellStorage, CellState } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { MockWorksheetCollection } from "./worksheet-collection.js";
import { MockRequestContext } from "./request-context.js";
import { FormulaEvaluator } from "./formula-evaluator.js";
import {
  parseAddress,
  cellAddressFromPosition,
  parseCellAddress,
} from "./address.js";
import { MockRange } from "./range.js";

export function createMockEnvironment() {
  const storage = new CellStorage();
  const customFunctions = new MockCustomFunctions();
  const dummyLoads: MockRange[] = [];
  const worksheets = new MockWorksheetCollection(storage, dummyLoads);
  const evaluator = new FormulaEvaluator(storage, customFunctions);

  let selectedSheet: string | undefined;
  let selectedAddress: string | undefined;

  const excel = {
    run: async <T>(
      callback: (context: MockRequestContext) => Promise<T>,
    ): Promise<T> => {
      const context = new MockRequestContext(
        storage,
        customFunctions,
        worksheets,
      );
      if (selectedSheet && selectedAddress) {
        context.workbook.setSelectedRange(selectedSheet, selectedAddress);
      }
      return await callback(context);
    },
  };

  const office = {
    onReady: (cb?: (info: { host: string; platform: string }) => void) => {
      const info = { host: "Excel", platform: "Web" };
      if (cb) cb(info);
      return Promise.resolve(info);
    },
    actions: {
      associate: () => {},
    },
  };

  const mockOfficeJs = {
    excel: {
      setCell(
        sheet: string,
        address: string,
        value: unknown,
      ): void | Promise<void> {
        if (typeof value === "object" && value !== null && "formula" in value) {
          return evaluator.evaluateAndStore(
            sheet,
            address,
            (value as { formula: string }).formula,
          );
        }
        storage.setValue(sheet, address, value);
      },

      getCell(sheet: string, address: string): CellState {
        return storage.getCell(sheet, address);
      },

      setCells(
        sheet: string,
        startAddress: string,
        values: unknown[][],
      ): void {
        const start = parseCellAddress(startAddress);
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const addr = cellAddressFromPosition(start.row + r, start.col + c);
            storage.setValue(sheet, addr, values[r][c]);
          }
        }
      },

      getCells(sheet: string, rangeAddress: string): CellState[][] {
        const range = parseAddress(rangeAddress);
        const rows: CellState[][] = [];
        for (let r = range.startRow; r <= range.endRow; r++) {
          const row: CellState[] = [];
          for (let c = range.startCol; c <= range.endCol; c++) {
            row.push(storage.getCell(sheet, cellAddressFromPosition(r, c)));
          }
          rows.push(row);
        }
        return rows;
      },

      setSelectedRange(sheet: string, address: string): void {
        selectedSheet = sheet;
        selectedAddress = address;
      },

      setActiveWorksheet(sheet: string): void {
        worksheets.setActiveWorksheet(sheet);
      },

      addWorksheet(name: string): void {
        worksheets.add(name);
      },

      async loadFunctionsMetadata(url: string): Promise<void> {
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(
            `Failed to fetch functions metadata: ${response.status}`,
          );
        }
        const metadata = await response.json();
        customFunctions.loadMetadata(metadata);
      },
    },

    reset(): void {
      storage.clearAll();
      customFunctions.reset();
      worksheets.reset();
      selectedSheet = undefined;
      selectedAddress = undefined;
    },
  };

  return { excel, office, customFunctions, mockOfficeJs };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/unit/setup.test.ts`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/setup.ts tests/unit/setup.test.ts
git commit -m "feat: add createMockEnvironment factory function"
```

---

### Task 2: Test `MockOfficeJs.excel` helpers via `createMockEnvironment`

**Files:**
- Modify: `tests/unit/setup.test.ts`

Add tests for all `mockOfficeJs.excel.*` methods. These mirror the existing `ExcelMock` tests but use the factory function instead.

- [ ] **Step 1: Write failing tests for `mockOfficeJs.excel` helpers**

Append to `tests/unit/setup.test.ts`:

```ts
describe("mockOfficeJs.excel", () => {
  test("setCell and getCell with plain value", () => {
    const { mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(42);
  });

  test("setCell with formula evaluates custom function", async () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("uninitialized cell returns empty strings", () => {
    const { mockOfficeJs } = createMockEnvironment();
    const cell = mockOfficeJs.excel.getCell("Sheet1", "Z99");
    expect(cell.value).toBe("");
    expect(cell.formula).toBe("");
  });

  test("setCells writes multiple cells at once", () => {
    const { mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCells("Sheet1", "A1", [[1, 2], [3, 4]]);
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(1);
    expect(mockOfficeJs.excel.getCell("Sheet1", "B1").value).toBe(2);
    expect(mockOfficeJs.excel.getCell("Sheet1", "A2").value).toBe(3);
    expect(mockOfficeJs.excel.getCell("Sheet1", "B2").value).toBe(4);
  });

  test("getCells reads a range of cells", () => {
    const { mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCells("Sheet1", "A1", [[1, 2], [3, 4]]);
    const cells = mockOfficeJs.excel.getCells("Sheet1", "A1:B2");
    expect(cells[0][0].value).toBe(1);
    expect(cells[0][1].value).toBe(2);
    expect(cells[1][0].value).toBe(3);
    expect(cells[1][1].value).toBe(4);
  });

  test("setSelectedRange is used by Excel.run", async () => {
    const { excel, customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("DOUBLE", (n: number) => n * 2);
    mockOfficeJs.excel.setCell("Sheet1", "A1", 5);
    mockOfficeJs.excel.setSelectedRange("Sheet1", "B1");
    await excel.run(async (context: any) => {
      const selected = context.workbook.getSelectedRange();
      selected.formulas = [["=DOUBLE(5)"]];
      await context.sync();
    });
    expect(mockOfficeJs.excel.getCell("Sheet1", "B1").value).toBe(10);
  });

  test("addWorksheet persists across Excel.run calls", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.addWorksheet("Sheet2");
    await excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getItem("Sheet2");
      expect(sheet).toBeDefined();
    });
  });

  test("setActiveWorksheet changes active worksheet", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.addWorksheet("Sheet2");
    mockOfficeJs.excel.setActiveWorksheet("Sheet2");
    await excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      expect(sheet.name).toBe("Sheet2");
    });
  });

  test("spill collision returns #SPILL!", async () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("MATRIX", () => [[1, 2], [3, 4]]);
    mockOfficeJs.excel.setCell("Sheet1", "C2", 999);
    await mockOfficeJs.excel.setCell("Sheet1", "B2", { formula: "=MATRIX()" });
    expect(mockOfficeJs.excel.getCell("Sheet1", "B2").value).toBe("#SPILL!");
  });

  test("function name lookup is case-insensitive", async () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=add(1, 2)" });
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(3);
  });
});
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `npx vitest run tests/unit/setup.test.ts`
Expected: PASS (implementation already exists from Task 1)

- [ ] **Step 3: Commit**

```bash
git add tests/unit/setup.test.ts
git commit -m "test: add mockOfficeJs.excel helper tests"
```

---

### Task 3: Test `Excel.run` shared state and `reset()`

**Files:**
- Modify: `tests/unit/setup.test.ts`

- [ ] **Step 1: Write tests for shared state and reset**

Append to `tests/unit/setup.test.ts`:

```ts
describe("Excel.run shared state", () => {
  test("Excel.run reads data set via mockOfficeJs.excel.setCell", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    await excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });

  test("Excel.run formula write evaluates custom function", async () => {
    const { excel, customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("DOUBLE", (n: number) => n * 2);
    await excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.formulas = [["=DOUBLE(21)"]];
      await context.sync();
    });
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(42);
  });
});

describe("Office.onReady", () => {
  test("calls callback with host and platform info", async () => {
    const { office } = createMockEnvironment();
    let receivedInfo: any;
    await office.onReady((info) => {
      receivedInfo = info;
    });
    expect(receivedInfo).toEqual({ host: "Excel", platform: "Web" });
  });

  test("returns a promise with host and platform info", async () => {
    const { office } = createMockEnvironment();
    const info = await office.onReady();
    expect(info).toEqual({ host: "Excel", platform: "Web" });
  });
});

describe("mockOfficeJs.reset", () => {
  test("clears all state", () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("ADD", (a: number, b: number) => a + b);
    mockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    mockOfficeJs.excel.addWorksheet("Sheet2");
    mockOfficeJs.reset();
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe("");
    expect(customFunctions.getFunction("ADD")).toBeUndefined();
  });

  test("clears selected range", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setSelectedRange("Sheet1", "B1");
    mockOfficeJs.reset();
    // After reset, getSelectedRange should throw (no selected range set)
    await excel.run(async (context: any) => {
      expect(() => context.workbook.getSelectedRange()).toThrow();
    });
  });
});
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `npx vitest run tests/unit/setup.test.ts`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add tests/unit/setup.test.ts
git commit -m "test: add shared state, Office.onReady, and reset tests"
```

---

### Task 4: Test `loadFunctionsMetadata`

**Files:**
- Modify: `tests/unit/setup.test.ts`

- [ ] **Step 1: Write tests for loadFunctionsMetadata**

Append to `tests/unit/setup.test.ts`:

Add `vi` and `afterEach` to the existing vitest import at the top of the file: `import { describe, expect, test, vi, afterEach } from "vitest";`

Append the following describe block:

```ts
describe("mockOfficeJs.excel.loadFunctionsMetadata", () => {
  afterEach(() => vi.unstubAllGlobals());

  test("fetches metadata from URL", async () => {
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({ functions: [] }),
    });
    vi.stubGlobal("fetch", fetchMock);

    const { mockOfficeJs } = createMockEnvironment();
    await mockOfficeJs.excel.loadFunctionsMetadata("/my/functions.json");
    expect(fetchMock).toHaveBeenCalledWith("/my/functions.json");
  });

  test("throws on fetch failure", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({ ok: false, status: 404 }),
    );

    const { mockOfficeJs } = createMockEnvironment();
    await expect(
      mockOfficeJs.excel.loadFunctionsMetadata("/missing.json"),
    ).rejects.toThrow("Failed to fetch functions metadata: 404");
  });

  test("metadata survives reset", async () => {
    const metadata = {
      functions: [
        { id: "ADD3", name: "ADD3", parameters: [{ name: "a" }, { name: "b" }, { name: "c" }] },
      ],
    };
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        json: () => Promise.resolve(metadata),
      }),
    );

    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    await mockOfficeJs.excel.loadFunctionsMetadata("/functions.json");
    customFunctions.associate("ADD3", () => 0);
    mockOfficeJs.reset();

    expect(customFunctions.getFunction("ADD3")).toBeUndefined();

    let receivedArgs: unknown[] = [];
    customFunctions.associate("ADD3", (...args: unknown[]) => {
      receivedArgs = args;
      return 0;
    });
    await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD3(1, 2)" });
    // Third parameter padded with null (metadata says 3 params, only 2 provided)
    expect(receivedArgs[2]).toBeNull();
  });

  test("loadFunctionsMetadata enables invocation context argument", async () => {
    const metadata = {
      functions: [
        {
          id: "ADD3",
          name: "ADD3",
          parameters: [
            { name: "a", type: "number" },
            { name: "b", type: "number" },
            { name: "c", type: "number", optional: true },
          ],
        },
      ],
    };
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: true,
        json: () => Promise.resolve(metadata),
      }),
    );

    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    await mockOfficeJs.excel.loadFunctionsMetadata("/functions.json");

    let receivedArgs: unknown[] = [];
    customFunctions.associate("ADD3", (...args: unknown[]) => {
      receivedArgs = args;
      return 0;
    });
    await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD3(1, 2)" });
    expect(receivedArgs[3]).toEqual({
      address: "Sheet1!A1",
      functionName: "ADD3",
    });
  });
});
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `npx vitest run tests/unit/setup.test.ts`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add tests/unit/setup.test.ts
git commit -m "test: add loadFunctionsMetadata tests"
```

---

### Task 5: Create `globals.d.ts` with `declare global`

**Files:**
- Create: `src/globals.d.ts`

- [ ] **Step 1: Create the global type declaration**

Create `src/globals.d.ts`:

```ts
import type { CellState } from "./cell-storage.js";

declare global {
  var MockOfficeJs: {
    excel: {
      setCell(sheet: string, address: string, value: unknown): void | Promise<void>;
      getCell(sheet: string, address: string): CellState;
      setCells(sheet: string, startAddress: string, values: unknown[][]): void;
      getCells(sheet: string, rangeAddress: string): CellState[][];
      setSelectedRange(sheet: string, address: string): void;
      setActiveWorksheet(sheet: string): void;
      addWorksheet(name: string): void;
      loadFunctionsMetadata(url: string): Promise<void>;
    };
    reset(): void;
  };
}
```

- [ ] **Step 2: Verify typecheck passes**

Run: `npx tsc --noEmit`
Expected: PASS (no type errors)

- [ ] **Step 3: Commit**

```bash
git add src/globals.d.ts
git commit -m "feat: add declare global for MockOfficeJs type"
```

---

### Task 6: Rewrite `index.ts` as side-effect-only entry point

**Files:**
- Rewrite: `src/index.ts`

- [ ] **Step 1: Rewrite `index.ts`**

Replace entire contents of `src/index.ts`:

```ts
import { createMockEnvironment } from "./setup.js";

const env = createMockEnvironment();

globalThis.Excel = env.excel as any;
globalThis.Office = env.office as any;
globalThis.CustomFunctions = env.customFunctions as any;
globalThis.MockOfficeJs = env.mockOfficeJs;
```

The `as any` casts are needed because the mock types don't fully match the `@types/office-js` global types — the mock implements a subset.

- [ ] **Step 2: Verify typecheck passes**

Run: `npx tsc --noEmit`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add src/index.ts
git commit -m "feat: rewrite index.ts as side-effect-only entry point"
```

---

### Task 7: Delete `browser.ts` and `excel-mock.ts`, update build config

**Files:**
- Delete: `src/browser.ts`
- Delete: `src/excel-mock.ts`
- Modify: `tsdown.config.ts`
- Modify: `package.json`

- [ ] **Step 1: Delete `browser.ts`**

```bash
git rm src/browser.ts
```

- [ ] **Step 2: Delete `excel-mock.ts`**

```bash
git rm src/excel-mock.ts
```

- [ ] **Step 3: Update `tsdown.config.ts`**

Replace with:

```ts
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
    entry: { office: "src/index.ts" },
    format: ["iife"],
    outDir: "dist",
  },
]);
```

- [ ] **Step 4: Update `package.json`**

Add `peerDependencies` and keep `@types/office-js` in `devDependencies`:

```json
{
  "peerDependencies": {
    "@types/office-js": "^1.0.0"
  }
}
```

The build script stays the same: `"build": "tsdown && mv dist/office.iife.js dist/office.js"`.

- [ ] **Step 5: Verify build succeeds**

Run: `npm run build`
Expected: Builds successfully, produces `dist/office.mjs`, `dist/office.d.mts`, and `dist/office.js`

- [ ] **Step 6: Verify typecheck passes**

Run: `npx tsc --noEmit`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "refactor: remove browser.ts and excel-mock.ts, unify build entry point"
```

---

### Task 8: Rewrite unit tests to use globals

**Files:**
- Delete: `tests/unit/excel-mock.test.ts`
- Create: `tests/unit/mock-office-js.test.ts`
- Rewrite: `tests/unit/integration.test.ts`

The old `excel-mock.test.ts` tests are now covered by `setup.test.ts` (Tasks 1-4). The new `mock-office-js.test.ts` tests that `import "../../src/index.js"` correctly sets up globals. `integration.test.ts` is rewritten to use globals.

- [ ] **Step 1: Delete old test file**

```bash
git rm tests/unit/excel-mock.test.ts
```

- [ ] **Step 2: Create `mock-office-js.test.ts` to verify global setup**

Create `tests/unit/mock-office-js.test.ts`:

```ts
import { describe, expect, test } from "vitest";
import "../../src/index.js";

describe("mock-office-js global setup", () => {
  test("Excel global is defined with run method", () => {
    expect(Excel).toBeDefined();
    expect(Excel.run).toBeTypeOf("function");
  });

  test("Office global is defined with onReady method", () => {
    expect(Office).toBeDefined();
    expect(Office.onReady).toBeTypeOf("function");
  });

  test("CustomFunctions global is defined with associate method", () => {
    expect(CustomFunctions).toBeDefined();
    expect(CustomFunctions.associate).toBeTypeOf("function");
  });

  test("MockOfficeJs global is defined with excel and reset", () => {
    expect(MockOfficeJs).toBeDefined();
    expect(MockOfficeJs.excel).toBeDefined();
    expect(MockOfficeJs.reset).toBeTypeOf("function");
  });

  test("globals share state", async () => {
    MockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
    MockOfficeJs.reset();
  });
});
```

- [ ] **Step 3: Rewrite `integration.test.ts`**

Replace entire contents of `tests/unit/integration.test.ts`:

```ts
import { describe, expect, test, afterEach } from "vitest";
import "../../src/index.js";

describe("E2E integration", () => {
  afterEach(() => MockOfficeJs.reset());

  test("full flow: register function, write formula via Excel.run, verify value", async () => {
    CustomFunctions.associate("TRIPLE", (n: number) => n * 3);
    MockOfficeJs.excel.setCell("Sheet1", "A1", 7);
    MockOfficeJs.excel.setSelectedRange("Sheet1", "B1");

    await Excel.run(async (context: any) => {
      const source = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      source.load("values");
      await context.sync();
      const val = source.values[0][0];
      const selected = context.workbook.getSelectedRange();
      selected.formulas = [[`=TRIPLE(${val})`]];
      await context.sync();
    });

    expect(MockOfficeJs.excel.getCell("Sheet1", "B1").value).toBe(21);
    expect(MockOfficeJs.excel.getCell("Sheet1", "B1").formula).toBe("=TRIPLE(7)");

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("B1");
      range.load(["values", "formulas"]);
      await context.sync();
      expect(range.values).toEqual([[21]]);
      expect(range.formulas).toEqual([["=TRIPLE(7)"]]);
    });
  });

  test("spill flow: function returns 2D array, verify all cells", async () => {
    CustomFunctions.associate("TABLE", () => [
      ["Name", "Score"],
      ["Alice", 95],
      ["Bob", 87],
    ]);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=TABLE()" });

    expect(MockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe("Name");
    expect(MockOfficeJs.excel.getCell("Sheet1", "B1").value).toBe("Score");
    expect(MockOfficeJs.excel.getCell("Sheet1", "A2").value).toBe("Alice");
    expect(MockOfficeJs.excel.getCell("Sheet1", "B2").value).toBe(95);
    expect(MockOfficeJs.excel.getCell("Sheet1", "A3").value).toBe("Bob");
    expect(MockOfficeJs.excel.getCell("Sheet1", "B3").value).toBe(87);
  });

  test("load/sync enforcement catches missing load", async () => {
    MockOfficeJs.excel.setCell("Sheet1", "A1", 42);

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      expect(() => range.values).toThrow();
      range.load("values");
      expect(() => range.values).toThrow();
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });

  test("multiple Excel.run calls share cell state", async () => {
    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.values = [[100]];
      await context.sync();
    });

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[100]]);
    });
  });

  test("reset isolates test cases", () => {
    MockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    MockOfficeJs.reset();
    expect(MockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe("");
  });
});
```

- [ ] **Step 4: Run all unit tests**

Run: `npx vitest run`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/unit/mock-office-js.test.ts tests/unit/integration.test.ts
git commit -m "test: rewrite unit tests to use global API"
```

---

### Task 9: Rewrite E2E tests

**Files:**
- Delete: `tests/e2e/excel-mock.e2e.test.ts`
- Create: `tests/e2e/mock-office-js.e2e.test.ts`

- [ ] **Step 1: Delete old E2E test file and create new one**

```bash
git rm tests/e2e/excel-mock.e2e.test.ts
```

Create `tests/e2e/mock-office-js.e2e.test.ts`:

```ts
import { test, expect } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const mockBundle = fs.readFileSync(
  path.join(__dirname, "../../dist/office.js"),
  "utf-8"
);

test.beforeEach(async ({ page }) => {
  await page.route("**/appsforoffice.microsoft.com/**", async (route) => {
    await route.fulfill({
      contentType: "application/javascript",
      body: mockBundle,
    });
  });

  await page.goto("/taskpane.html");
  await page.waitForFunction(() => (window as any).MockOfficeJs !== undefined);
});

test("Excel.run can read cell values set via mock", async ({ page }) => {
  const value = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    MockOfficeJs.excel.setCell("Sheet1", "A1", 42);

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

test("CustomFunctions.associate registers functions and formulas evaluate", async ({
  page,
}) => {
  const value = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(2, 3)" });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(value).toBe(5);
});

test("worksheet operations: add and switch active worksheet", async ({
  page,
}) => {
  const sheetName = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    MockOfficeJs.excel.addWorksheet("Sheet2");
    MockOfficeJs.excel.setActiveWorksheet("Sheet2");

    let name = "";
    await (window as any).Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      name = sheet.name;
    });
    return name;
  });

  expect(sheetName).toBe("Sheet2");
});

test("accessing range properties without load/sync throws an error", async ({
  page,
}) => {
  const errorMessage = await page.evaluate(async () => {
    let error = "";
    await (window as any).Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("A1");
      try {
        const _values = range.values;
      } catch (e: any) {
        error = e.message;
      }
    });
    return error;
  });

  expect(errorMessage).toBeTruthy();
});

test("MockOfficeJs.reset() clears cell values", async ({ page }) => {
  await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    MockOfficeJs.excel.setCell("Sheet1", "A1", 99);
  });

  await page.evaluate(() => (window as any).MockOfficeJs.reset());

  const value = await page.evaluate(async () => {
    let result: any[][] = [];
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

  expect(value).toEqual([[""]]);
});
```

- [ ] **Step 2: Run E2E tests**

Run: `npm run test:e2e`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git rm tests/e2e/excel-mock.e2e.test.ts
git add tests/e2e/mock-office-js.e2e.test.ts
git commit -m "test: rewrite E2E tests to use MockOfficeJs global"
```

---

### Task 10: Final verification

- [ ] **Step 1: Run full test suite**

Run: `npm test`
Expected: All unit tests, type tests, and E2E tests pass.

- [ ] **Step 2: Run typecheck**

Run: `npx tsc --noEmit`
Expected: PASS

- [ ] **Step 3: Verify build output**

Run: `npm run build && ls -la dist/`
Expected: `office.mjs`, `office.d.mts`, `office.js` all present.

- [ ] **Step 4: Inspect built IIFE to confirm globals are set**

Run: `grep -c "globalThis" dist/office.js`
Expected: Non-zero (confirms globalThis assignments are in the bundle).

- [ ] **Step 5: Verify `declare global` appears in built DTS**

Run: `grep "MockOfficeJs" dist/office.d.mts`
Expected: Contains `declare global` block with `MockOfficeJs`. If missing, the fix is to add `import "./globals.js"` to `index.ts` so tsdown includes the global augmentation in the DTS output, or inline the `declare global` block into `setup.ts`.

- [ ] **Step 6: Commit any remaining changes**

If any fixups were needed, commit them now.
