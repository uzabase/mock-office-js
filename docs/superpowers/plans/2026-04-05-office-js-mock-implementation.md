# mock-office-js Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an in-memory mock of the Excel JavaScript API and CustomFunctions runtime for E2E-style testing of Excel Add-ins.

**Architecture:** Two-layer design — a Core layer that reproduces the Excel JavaScript API (Excel.run, RequestContext, Workbook, Worksheet, Range, CustomFunctions) backed by an in-memory CellStorage, and a Wrapper layer (ExcelMock) that provides convenience methods (setCell, getCell, reset) for test setup and assertions. Formula evaluation is limited to registered custom functions; native Excel functions are out of scope.

**Tech Stack:** TypeScript, Vitest (runtime + type tests), `@types/office-js` and `@types/custom-functions-runtime` (devDependencies for type conformance testing)

**Spec:** `docs/superpowers/specs/2026-04-05-mock-office-js-design.md`

**Key architectural note:** ExcelMock owns shared CellStorage + MockWorksheetCollection, passed to each MockRequestContext. FormulaEvaluator is a shared utility used by both the wrapper (ExcelMock.setCell) and the core (MockRequestContext.sync). This avoids duplicating formula evaluation logic.

---

## Task 1: Project Setup

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `tsconfig.test-d.json`
- Create: `vitest.config.ts`
- Modify: `.gitignore`

- [ ] **Step 1: Initialize package.json**

```bash
cd /Users/ot07/Development/npm-packages/mock-office-js
npm init -y
```

Then update `package.json`:

```json
{
  "name": "mock-office-js",
  "version": "0.1.0",
  "description": "In-memory mock of Excel JavaScript API for testing Excel Add-ins",
  "main": "dist/index.js",
  "types": "dist/index.d.ts",
  "scripts": {
    "build": "tsc",
    "test": "vitest run",
    "test:watch": "vitest",
    "test:typecheck": "vitest typecheck"
  },
  "license": "MIT"
}
```

- [ ] **Step 2: Install dependencies**

```bash
npm install -D typescript vitest @types/office-js @types/custom-functions-runtime
```

- [ ] **Step 3: Create tsconfig.json**

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ES2020",
    "moduleResolution": "bundler",
    "lib": ["ES2020"],
    "declaration": true,
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true
  },
  "include": ["src/**/*.ts"]
}
```

- [ ] **Step 4: Create tsconfig.test-d.json**

Type tests need `@types/office-js` and `@types/custom-functions-runtime` for conformance checking. Production code must NOT see these types.

```json
{
  "extends": "./tsconfig.json",
  "include": ["src/**/*.ts", "tests/**/*.test-d.ts"],
  "compilerOptions": {
    "rootDir": ".",
    "types": ["office-js", "custom-functions-runtime"],
    "noEmit": true
  }
}
```

- [ ] **Step 5: Create vitest.config.ts**

```typescript
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    typecheck: {
      tsconfig: "./tsconfig.test-d.json",
    },
  },
});
```

- [ ] **Step 6: Update .gitignore**

Append to existing `.gitignore`:

```
node_modules/
dist/
```

- [ ] **Step 7: Create empty entry point and verify setup**

Create `src/index.ts`:

```typescript
export {};
```

Run:

```bash
npx vitest run
```

Expected: Vitest runs with no tests found (no error).

- [ ] **Step 8: Commit**

```bash
git add package.json tsconfig.json tsconfig.test-d.json vitest.config.ts .gitignore src/index.ts package-lock.json
git commit -m "chore: initialize project with TypeScript and Vitest"
```

---

## Task 2: Address Utilities

Address parsing is used by CellStorage, MockRange, and MockWorksheet. Build it first as a pure utility with no dependencies.

**Files:**
- Create: `src/address.ts`
- Create: `tests/address.test.ts`

- [ ] **Step 1: Write failing tests for column letter ↔ number conversion**

Create `tests/address.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { columnLetterToIndex, indexToColumnLetter } from "../src/address";

describe("columnLetterToIndex", () => {
  test("converts single letter columns", () => {
    expect(columnLetterToIndex("A")).toBe(0);
    expect(columnLetterToIndex("B")).toBe(1);
    expect(columnLetterToIndex("Z")).toBe(25);
  });

  test("converts multi-letter columns", () => {
    expect(columnLetterToIndex("AA")).toBe(26);
    expect(columnLetterToIndex("AB")).toBe(27);
    expect(columnLetterToIndex("AZ")).toBe(51);
  });

  test("is case-insensitive", () => {
    expect(columnLetterToIndex("a")).toBe(0);
    expect(columnLetterToIndex("aa")).toBe(26);
  });
});

describe("indexToColumnLetter", () => {
  test("converts single digit indices", () => {
    expect(indexToColumnLetter(0)).toBe("A");
    expect(indexToColumnLetter(1)).toBe("B");
    expect(indexToColumnLetter(25)).toBe("Z");
  });

  test("converts multi-letter indices", () => {
    expect(indexToColumnLetter(26)).toBe("AA");
    expect(indexToColumnLetter(27)).toBe("AB");
    expect(indexToColumnLetter(51)).toBe("AZ");
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/address.test.ts
```

Expected: FAIL — module not found.

- [ ] **Step 3: Implement column conversion functions**

Create `src/address.ts`:

```typescript
export function columnLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function indexToColumnLetter(index: number): string {
  let letter = "";
  let n = index + 1;
  while (n > 0) {
    n--;
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26);
  }
  return letter;
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/address.test.ts
```

Expected: PASS

- [ ] **Step 5: Write failing tests for address parsing**

Add to `tests/address.test.ts`:

```typescript
import { parseAddress, parseCellAddress, resolveRangeAddresses } from "../src/address";

describe("parseCellAddress", () => {
  test("parses simple cell address", () => {
    expect(parseCellAddress("A1")).toEqual({ row: 0, col: 0 });
    expect(parseCellAddress("B3")).toEqual({ row: 2, col: 1 });
    expect(parseCellAddress("AA10")).toEqual({ row: 9, col: 26 });
  });

  test("handles absolute references by stripping $", () => {
    expect(parseCellAddress("$A$1")).toEqual({ row: 0, col: 0 });
    expect(parseCellAddress("$B3")).toEqual({ row: 2, col: 1 });
  });
});

describe("parseAddress", () => {
  test("parses single cell address", () => {
    expect(parseAddress("A1")).toEqual({
      startRow: 0, startCol: 0, endRow: 0, endCol: 0,
    });
  });

  test("parses range address", () => {
    expect(parseAddress("A1:C2")).toEqual({
      startRow: 0, startCol: 0, endRow: 1, endCol: 2,
    });
  });

  test("parses sheet-qualified address", () => {
    expect(parseAddress("Sheet1!A1:B2")).toEqual({
      startRow: 0, startCol: 0, endRow: 1, endCol: 1,
    });
  });

  test("parses quoted sheet name with spaces", () => {
    expect(parseAddress("'My Sheet'!A1")).toEqual({
      startRow: 0, startCol: 0, endRow: 0, endCol: 0,
    });
  });
});

describe("resolveRangeAddresses", () => {
  test("resolves single cell to one address", () => {
    expect(resolveRangeAddresses("A1")).toEqual(["A1"]);
  });

  test("resolves range to all cell addresses", () => {
    const result = resolveRangeAddresses("A1:B2");
    expect(result).toEqual(["A1", "B1", "A2", "B2"]);
  });

  test("resolves range addresses in row-major order", () => {
    const result = resolveRangeAddresses("B2:C3");
    expect(result).toEqual(["B2", "C2", "B3", "C3"]);
  });
});
```

- [ ] **Step 6: Run test to verify it fails**

```bash
npx vitest run tests/address.test.ts
```

Expected: FAIL — functions not exported.

- [ ] **Step 7: Implement address parsing functions**

Add to `src/address.ts`:

```typescript
export interface CellPosition {
  row: number;
  col: number;
}

export interface RangePosition {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function parseCellAddress(address: string): CellPosition {
  const cleaned = address.replace(/\$/g, "");
  const match = cleaned.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }
  return {
    col: columnLetterToIndex(match[1]),
    row: parseInt(match[2], 10) - 1,
  };
}

export function parseAddress(address: string): RangePosition {
  // Strip sheet qualifier (e.g., "Sheet1!" or "'My Sheet'!")
  let cellPart = address;
  const sheetSepIndex = address.lastIndexOf("!");
  if (sheetSepIndex !== -1) {
    cellPart = address.substring(sheetSepIndex + 1);
  }

  const parts = cellPart.split(":");
  const start = parseCellAddress(parts[0]);

  if (parts.length === 1) {
    return {
      startRow: start.row,
      startCol: start.col,
      endRow: start.row,
      endCol: start.col,
    };
  }

  const end = parseCellAddress(parts[1]);
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

export function cellAddressFromPosition(row: number, col: number): string {
  return `${indexToColumnLetter(col)}${row + 1}`;
}

export function resolveRangeAddresses(address: string): string[] {
  const range = parseAddress(address);
  const addresses: string[] = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      addresses.push(cellAddressFromPosition(row, col));
    }
  }
  return addresses;
}
```

- [ ] **Step 8: Run test to verify it passes**

```bash
npx vitest run tests/address.test.ts
```

Expected: PASS

- [ ] **Step 9: Commit**

```bash
git add src/address.ts tests/address.test.ts
git commit -m "feat: add address parsing utilities"
```

---

## Task 3: CellStorage

**Files:**
- Create: `src/cell-storage.ts`
- Create: `tests/cell-storage.test.ts`

- [ ] **Step 1: Write failing tests for basic get/set**

Create `tests/cell-storage.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { CellStorage } from "../src/cell-storage";

describe("CellStorage", () => {
  test("returns empty string for uninitialized cell", () => {
    const storage = new CellStorage();
    const cell = storage.getCell("Sheet1", "A1");
    expect(cell.value).toBe("");
    expect(cell.formula).toBe("");
  });

  test("stores and retrieves a value", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 42);
    expect(storage.getCell("Sheet1", "A1").value).toBe(42);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("");
  });

  test("stores and retrieves a formula with value", () => {
    const storage = new CellStorage();
    storage.setFormula("Sheet1", "A1", "=ADD(1,2)", 3);
    const cell = storage.getCell("Sheet1", "A1");
    expect(cell.value).toBe(3);
    expect(cell.formula).toBe("=ADD(1,2)");
  });

  test("setValue overwrites previous formula", () => {
    const storage = new CellStorage();
    storage.setFormula("Sheet1", "A1", "=ADD(1,2)", 3);
    storage.setValue("Sheet1", "A1", "hello");
    const cell = storage.getCell("Sheet1", "A1");
    expect(cell.value).toBe("hello");
    expect(cell.formula).toBe("");
  });

  test("different sheets are independent", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet2", "A1", 2);
    expect(storage.getCell("Sheet1", "A1").value).toBe(1);
    expect(storage.getCell("Sheet2", "A1").value).toBe(2);
  });

  test("clear removes cell content", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 42);
    storage.clear("Sheet1", "A1");
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
  });

  test("clearSheet removes all cells in a sheet", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet1", "B1", 2);
    storage.clearSheet("Sheet1");
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
    expect(storage.getCell("Sheet1", "B1").value).toBe("");
  });

  test("clearAll removes all cells across all sheets", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet2", "A1", 2);
    storage.clearAll();
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
    expect(storage.getCell("Sheet2", "A1").value).toBe("");
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/cell-storage.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement CellStorage**

Create `src/cell-storage.ts`:

```typescript
export interface CellState {
  value: unknown;
  formula: string;
  spilledFrom?: string;
}

const EMPTY_CELL: CellState = { value: "", formula: "" };

export class CellStorage {
  private sheets = new Map<string, Map<string, CellState>>();

  getCell(sheet: string, address: string): CellState {
    return this.sheets.get(sheet)?.get(address) ?? { ...EMPTY_CELL };
  }

  setValue(sheet: string, address: string, value: unknown): void {
    this.ensureSheet(sheet).set(address, { value, formula: "" });
  }

  setFormula(sheet: string, address: string, formula: string, value: unknown): void {
    this.ensureSheet(sheet).set(address, { value, formula });
  }

  clear(sheet: string, address: string): void {
    this.sheets.get(sheet)?.delete(address);
  }

  clearSheet(sheet: string): void {
    this.sheets.delete(sheet);
  }

  clearAll(): void {
    this.sheets.clear();
  }

  private ensureSheet(sheet: string): Map<string, CellState> {
    let sheetMap = this.sheets.get(sheet);
    if (!sheetMap) {
      sheetMap = new Map();
      this.sheets.set(sheet, sheetMap);
    }
    return sheetMap;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/cell-storage.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/cell-storage.ts tests/cell-storage.test.ts
git commit -m "feat: add CellStorage with basic get/set/clear"
```

---

## Task 4: CellStorage Spill Support

**Files:**
- Modify: `src/cell-storage.ts`
- Modify: `tests/cell-storage.test.ts`

- [ ] **Step 1: Write failing tests for spill behavior**

Add to `tests/cell-storage.test.ts`:

```typescript
describe("spill", () => {
  test("spills 2D array to adjacent cells", () => {
    const storage = new CellStorage();
    storage.setFormulaWithSpill("Sheet1", "B2", "=MATRIX()", [
      [1, 2],
      [3, 4],
    ]);

    expect(storage.getCell("Sheet1", "B2")).toEqual({ value: 1, formula: "=MATRIX()" });
    expect(storage.getCell("Sheet1", "C2")).toEqual({ value: 2, formula: "", spilledFrom: "B2" });
    expect(storage.getCell("Sheet1", "B3")).toEqual({ value: 3, formula: "", spilledFrom: "B2" });
    expect(storage.getCell("Sheet1", "C3")).toEqual({ value: 4, formula: "", spilledFrom: "B2" });
  });

  test("clearing spill origin clears all spilled cells", () => {
    const storage = new CellStorage();
    storage.setFormulaWithSpill("Sheet1", "B2", "=MATRIX()", [
      [1, 2],
      [3, 4],
    ]);
    storage.clear("Sheet1", "B2");

    expect(storage.getCell("Sheet1", "B2").value).toBe("");
    expect(storage.getCell("Sheet1", "C2").value).toBe("");
    expect(storage.getCell("Sheet1", "B3").value).toBe("");
    expect(storage.getCell("Sheet1", "C3").value).toBe("");
  });

  test("spill collision returns #SPILL! error", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "C2", 999);
    storage.setFormulaWithSpill("Sheet1", "B2", "=MATRIX()", [
      [1, 2],
      [3, 4],
    ]);

    expect(storage.getCell("Sheet1", "B2").value).toBe("#SPILL!");
    expect(storage.getCell("Sheet1", "B2").formula).toBe("=MATRIX()");
    expect(storage.getCell("Sheet1", "C2").value).toBe(999); // unchanged
  });

  test("overwriting spilled non-origin cell causes #SPILL! on origin", () => {
    const storage = new CellStorage();
    storage.setFormulaWithSpill("Sheet1", "B2", "=MATRIX()", [
      [1, 2],
      [3, 4],
    ]);
    storage.setValue("Sheet1", "C2", 999);

    expect(storage.getCell("Sheet1", "B2").value).toBe("#SPILL!");
    expect(storage.getCell("Sheet1", "B2").formula).toBe("=MATRIX()");
    expect(storage.getCell("Sheet1", "B3").value).toBe("");
    expect(storage.getCell("Sheet1", "C3").value).toBe("");
  });

  test("spill with 1D result (single row)", () => {
    const storage = new CellStorage();
    storage.setFormulaWithSpill("Sheet1", "A1", "=ROW()", [
      [10, 20, 30],
    ]);

    expect(storage.getCell("Sheet1", "A1").value).toBe(10);
    expect(storage.getCell("Sheet1", "B1")).toEqual({ value: 20, formula: "", spilledFrom: "A1" });
    expect(storage.getCell("Sheet1", "C1")).toEqual({ value: 30, formula: "", spilledFrom: "A1" });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/cell-storage.test.ts
```

Expected: FAIL — `setFormulaWithSpill` not defined.

- [ ] **Step 3: Implement spill support in CellStorage**

Add imports to `src/cell-storage.ts`:

```typescript
import { cellAddressFromPosition, parseCellAddress } from "./address";
```

Add methods to `CellStorage` class:

```typescript
  setFormulaWithSpill(
    sheet: string,
    address: string,
    formula: string,
    resultArray: unknown[][],
  ): void {
    const origin = parseCellAddress(address);
    const sheetMap = this.ensureSheet(sheet);

    // Check for spill collision
    for (let r = 0; r < resultArray.length; r++) {
      for (let c = 0; c < resultArray[r].length; c++) {
        if (r === 0 && c === 0) continue; // skip origin
        const targetAddr = cellAddressFromPosition(origin.row + r, origin.col + c);
        const existing = sheetMap.get(targetAddr);
        if (existing && existing.value !== "" && !existing.spilledFrom) {
          // Collision — set #SPILL! on origin, don't spill
          sheetMap.set(address, { value: "#SPILL!", formula });
          return;
        }
      }
    }

    // No collision — write spill
    for (let r = 0; r < resultArray.length; r++) {
      for (let c = 0; c < resultArray[r].length; c++) {
        const targetAddr = cellAddressFromPosition(origin.row + r, origin.col + c);
        if (r === 0 && c === 0) {
          sheetMap.set(targetAddr, { value: resultArray[r][c], formula });
        } else {
          sheetMap.set(targetAddr, { value: resultArray[r][c], formula: "", spilledFrom: address });
        }
      }
    }
  }
```

Also update `setValue` to handle overwriting a spilled cell:

```typescript
  setValue(sheet: string, address: string, value: unknown): void {
    const sheetMap = this.ensureSheet(sheet);
    const existing = sheetMap.get(address);

    // If overwriting a spilled (non-origin) cell, invalidate the origin
    if (existing?.spilledFrom) {
      this.invalidateSpillOrigin(sheet, existing.spilledFrom);
    }

    sheetMap.set(address, { value, formula: "" });
  }

  private invalidateSpillOrigin(sheet: string, originAddress: string): void {
    const sheetMap = this.sheets.get(sheet);
    if (!sheetMap) return;

    const origin = sheetMap.get(originAddress);
    if (!origin) return;

    // Set #SPILL! on origin
    sheetMap.set(originAddress, { value: "#SPILL!", formula: origin.formula });

    // Remove all spilled cells for this origin
    for (const [addr, cell] of sheetMap) {
      if (cell.spilledFrom === originAddress) {
        sheetMap.delete(addr);
      }
    }
  }
```

Also update `clear` to handle clearing a spill origin:

```typescript
  clear(sheet: string, address: string): void {
    const sheetMap = this.sheets.get(sheet);
    if (!sheetMap) return;

    // If clearing a spill origin, remove all spilled cells
    for (const [addr, cell] of sheetMap) {
      if (cell.spilledFrom === address) {
        sheetMap.delete(addr);
      }
    }

    sheetMap.delete(address);
  }
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/cell-storage.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/cell-storage.ts tests/cell-storage.test.ts
git commit -m "feat: add spill support with collision detection to CellStorage"
```

---

## Task 5: Formula Evaluator

Shared formula evaluation logic used by both MockRequestContext (core) and ExcelMock (wrapper). Avoids duplicating evaluation code.

**Files:**
- Create: `src/formula-evaluator.ts`
- Create: `tests/formula-evaluator.test.ts`

- [ ] **Step 1: Write failing tests**

Create `tests/formula-evaluator.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { FormulaEvaluator } from "../src/formula-evaluator";
import { CellStorage } from "../src/cell-storage";
import { MockCustomFunctions } from "../src/custom-functions-mock";

describe("FormulaEvaluator", () => {
  function createEvaluator() {
    const storage = new CellStorage();
    const cf = new MockCustomFunctions();
    const evaluator = new FormulaEvaluator(storage, cf);
    return { evaluator, storage, cf };
  }

  test("evaluates registered function and stores result", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("ADD", (a: number, b: number) => a + b);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=ADD(1, 2)");
    expect(storage.getCell("Sheet1", "A1").value).toBe(3);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("=ADD(1, 2)");
  });

  test("stores #NAME? for unregistered function", async () => {
    const { evaluator, storage } = createEvaluator();
    await evaluator.evaluateAndStore("Sheet1", "A1", "=UNKNOWN(1)");
    expect(storage.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

  test("handles async functions", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("ASYNC_ADD", async (a: number, b: number) => a + b);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=ASYNC_ADD(10, 20)");
    expect(storage.getCell("Sheet1", "A1").value).toBe(30);
  });

  test("passes Invocation with address and functionName", async () => {
    const { evaluator, cf } = createEvaluator();
    let receivedInvocation: unknown;
    cf.associate("CAPTURE", (...args: unknown[]) => {
      receivedInvocation = args[args.length - 1];
      return 0;
    });
    await evaluator.evaluateAndStore("Sheet1", "B3", "=CAPTURE()");
    expect(receivedInvocation).toEqual({
      address: "Sheet1!B3",
      functionName: "CAPTURE",
    });
  });

  test("spills 2D array result to adjacent cells", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("MATRIX", () => [[1, 2], [3, 4]]);
    await evaluator.evaluateAndStore("Sheet1", "B2", "=MATRIX()");
    expect(storage.getCell("Sheet1", "B2").value).toBe(1);
    expect(storage.getCell("Sheet1", "C2").value).toBe(2);
    expect(storage.getCell("Sheet1", "B3").value).toBe(3);
    expect(storage.getCell("Sheet1", "C3").value).toBe(4);
  });

  test("stores #VALUE! when function throws", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("FAIL", () => { throw new Error("boom"); });
    await evaluator.evaluateAndStore("Sheet1", "A1", "=FAIL()");
    expect(storage.getCell("Sheet1", "A1").value).toBe("#VALUE!");
  });

  test("non-formula string stores as plain value", async () => {
    const { evaluator, storage } = createEvaluator();
    await evaluator.evaluateAndStore("Sheet1", "A1", "hello");
    expect(storage.getCell("Sheet1", "A1").value).toBe("hello");
    expect(storage.getCell("Sheet1", "A1").formula).toBe("");
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/formula-evaluator.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement FormulaEvaluator**

Create `src/formula-evaluator.ts`:

```typescript
import { CellStorage } from "./cell-storage";
import { MockCustomFunctions } from "./custom-functions-mock";
import { parseFormula } from "./formula-parser";

export class FormulaEvaluator {
  constructor(
    private _storage: CellStorage,
    private _customFunctions: MockCustomFunctions,
  ) {}

  async evaluateAndStore(sheet: string, address: string, formulaStr: string): Promise<void> {
    const parsed = parseFormula(formulaStr);

    if (!parsed) {
      // Not a formula — store as plain value
      this._storage.setValue(sheet, address, formulaStr);
      return;
    }

    const fn = this._customFunctions.getFunction(parsed.functionName);
    if (!fn) {
      this._storage.setFormula(sheet, address, formulaStr, "#NAME?");
      return;
    }

    const invocation = {
      address: `${sheet}!${address}`,
      functionName: parsed.functionName.toUpperCase(),
    };

    try {
      const result = await fn(...parsed.args, invocation);
      if (Array.isArray(result) && Array.isArray(result[0])) {
        this._storage.setFormulaWithSpill(sheet, address, formulaStr, result);
      } else {
        this._storage.setFormula(sheet, address, formulaStr, result);
      }
    } catch {
      this._storage.setFormula(sheet, address, formulaStr, "#VALUE!");
    }
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/formula-evaluator.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/formula-evaluator.ts tests/formula-evaluator.test.ts
git commit -m "feat: add FormulaEvaluator for shared formula evaluation logic"
```

---

## Task 6: Formula Parser (was Task 5 — renumbered)

**Files:**
- Create: `src/formula-parser.ts`
- Create: `tests/formula-parser.test.ts`

- [ ] **Step 1: Write failing tests**

Create `tests/formula-parser.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { parseFormula } from "../src/formula-parser";

describe("parseFormula", () => {
  test("parses function with no arguments", () => {
    expect(parseFormula("=NOW()")).toEqual({
      functionName: "NOW",
      args: [],
    });
  });

  test("parses function with string argument", () => {
    expect(parseFormula('=GETPRICE("AAPL")')).toEqual({
      functionName: "GETPRICE",
      args: ["AAPL"],
    });
  });

  test("parses function with number arguments", () => {
    expect(parseFormula("=ADD(1, 2)")).toEqual({
      functionName: "ADD",
      args: [1, 2],
    });
  });

  test("parses function with negative number", () => {
    expect(parseFormula("=ADD(-1, 2.5)")).toEqual({
      functionName: "ADD",
      args: [-1, 2.5],
    });
  });

  test("parses function with boolean arguments", () => {
    expect(parseFormula("=CHECK(TRUE, FALSE)")).toEqual({
      functionName: "CHECK",
      args: [true, false],
    });
  });

  test("parses function with mixed argument types", () => {
    expect(parseFormula('=FUNC("hello", 42, TRUE)')).toEqual({
      functionName: "FUNC",
      args: ["hello", 42, true],
    });
  });

  test("parses namespace-prefixed function name", () => {
    expect(parseFormula('=CONTOSO.GETPRICE("AAPL")')).toEqual({
      functionName: "CONTOSO.GETPRICE",
      args: ["AAPL"],
    });
  });

  test("returns null for non-formula string", () => {
    expect(parseFormula("hello")).toBeNull();
  });

  test("returns null for native Excel functions with cell references", () => {
    expect(parseFormula("=SUM(A1:A5)")).toEqual({
      functionName: "SUM",
      args: ["A1:A5"],
    });
  });

  test("parses string with escaped quotes", () => {
    expect(parseFormula('=FUNC("he said ""hi""")')).toEqual({
      functionName: "FUNC",
      args: ['he said "hi"'],
    });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/formula-parser.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement formula parser**

Create `src/formula-parser.ts`:

```typescript
export interface ParsedFormula {
  functionName: string;
  args: unknown[];
}

export function parseFormula(formula: string): ParsedFormula | null {
  if (!formula.startsWith("=")) return null;

  const content = formula.substring(1);
  const parenOpen = content.indexOf("(");
  if (parenOpen === -1) return null;

  const functionName = content.substring(0, parenOpen).trim();
  if (!functionName) return null;

  const parenClose = content.lastIndexOf(")");
  if (parenClose === -1) return null;

  const argsString = content.substring(parenOpen + 1, parenClose).trim();
  const args = argsString === "" ? [] : parseArgs(argsString);

  return { functionName, args };
}

function parseArgs(argsString: string): unknown[] {
  const args: unknown[] = [];
  let current = "";
  let inString = false;
  let i = 0;

  while (i < argsString.length) {
    const char = argsString[i];

    if (inString) {
      if (char === '"') {
        // Check for escaped quote ("")
        if (i + 1 < argsString.length && argsString[i + 1] === '"') {
          current += '"';
          i += 2;
          continue;
        }
        inString = false;
        i++;
        continue;
      }
      current += char;
      i++;
      continue;
    }

    if (char === '"') {
      inString = true;
      i++;
      continue;
    }

    if (char === ",") {
      args.push(parseArgValue(current.trim()));
      current = "";
      i++;
      continue;
    }

    current += char;
    i++;
  }

  if (current.trim() !== "") {
    args.push(parseArgValue(current.trim()));
  }

  return args;
}

function parseArgValue(value: string): unknown {
  if (value.startsWith('"') && value.endsWith('"')) {
    return value.slice(1, -1);
  }

  const upper = value.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;

  const num = Number(value);
  if (!isNaN(num) && value !== "") return num;

  // Unresolved token (e.g., cell reference) — return as string
  return value;
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/formula-parser.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/formula-parser.ts tests/formula-parser.test.ts
git commit -m "feat: add formula parser for custom function formulas"
```

---

## Task 6: MockCustomFunctions

**Files:**
- Create: `src/custom-functions-mock.ts`
- Create: `tests/custom-functions-mock.test.ts`

- [ ] **Step 1: Write failing tests**

Create `tests/custom-functions-mock.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { MockCustomFunctions } from "../src/custom-functions-mock";

describe("MockCustomFunctions", () => {
  test("associate registers a single function", () => {
    const cf = new MockCustomFunctions();
    const fn = (a: number, b: number) => a + b;
    cf.associate("ADD", fn);
    expect(cf.getFunction("ADD")).toBe(fn);
  });

  test("associate registers multiple functions via object", () => {
    const cf = new MockCustomFunctions();
    const add = (a: number, b: number) => a + b;
    const mul = (a: number, b: number) => a * b;
    cf.associate({ ADD: add, MUL: mul });
    expect(cf.getFunction("ADD")).toBe(add);
    expect(cf.getFunction("MUL")).toBe(mul);
  });

  test("function lookup is case-insensitive", () => {
    const cf = new MockCustomFunctions();
    const fn = () => 1;
    cf.associate("GETPRICE", fn);
    expect(cf.getFunction("getprice")).toBe(fn);
    expect(cf.getFunction("GetPrice")).toBe(fn);
  });

  test("getFunction returns undefined for unregistered function", () => {
    const cf = new MockCustomFunctions();
    expect(cf.getFunction("UNKNOWN")).toBeUndefined();
  });

  test("namespace-prefixed function names work", () => {
    const cf = new MockCustomFunctions();
    const fn = () => 42;
    cf.associate("CONTOSO.ADD", fn);
    expect(cf.getFunction("CONTOSO.ADD")).toBe(fn);
    expect(cf.getFunction("contoso.add")).toBe(fn);
  });

  test("Error class stores code and message", () => {
    const cf = new MockCustomFunctions();
    const error = new cf.Error(cf.ErrorCode.invalidValue, "bad input");
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toBe("bad input");
  });

  test("ErrorCode enum has correct values", () => {
    const cf = new MockCustomFunctions();
    expect(cf.ErrorCode.invalidValue).toBe("#VALUE!");
    expect(cf.ErrorCode.notAvailable).toBe("#N/A");
    expect(cf.ErrorCode.divisionByZero).toBe("#DIV/0!");
    expect(cf.ErrorCode.invalidNumber).toBe("#NUM!");
    expect(cf.ErrorCode.nullReference).toBe("#NULL!");
    expect(cf.ErrorCode.invalidName).toBe("#NAME?");
    expect(cf.ErrorCode.invalidReference).toBe("#REF!");
  });

  test("reset clears all registered functions", () => {
    const cf = new MockCustomFunctions();
    cf.associate("ADD", () => 0);
    cf.reset();
    expect(cf.getFunction("ADD")).toBeUndefined();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/custom-functions-mock.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement MockCustomFunctions**

Create `src/custom-functions-mock.ts`:

```typescript
export class MockCustomFunctions {
  private registry = new Map<string, Function>();

  associate(idOrMappings: string | Record<string, Function>, fn?: Function): void {
    if (typeof idOrMappings === "string") {
      this.registry.set(idOrMappings.toUpperCase(), fn!);
    } else {
      for (const [id, func] of Object.entries(idOrMappings)) {
        this.registry.set(id.toUpperCase(), func);
      }
    }
  }

  getFunction(id: string): Function | undefined {
    return this.registry.get(id.toUpperCase());
  }

  reset(): void {
    this.registry.clear();
  }

  Error = MockCustomFunctionsError;

  ErrorCode = {
    invalidValue: "#VALUE!" as const,
    notAvailable: "#N/A" as const,
    divisionByZero: "#DIV/0!" as const,
    invalidNumber: "#NUM!" as const,
    nullReference: "#NULL!" as const,
    invalidName: "#NAME?" as const,
    invalidReference: "#REF!" as const,
  };
}

class MockCustomFunctionsError {
  code: string;
  message?: string;

  constructor(code: string, message?: string) {
    this.code = code;
    this.message = message;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/custom-functions-mock.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/custom-functions-mock.ts tests/custom-functions-mock.test.ts
git commit -m "feat: add MockCustomFunctions with associate, Error, and ErrorCode"
```

---

## Task 7: MockRange with load/sync Constraints

**Files:**
- Create: `src/range.ts`
- Create: `tests/range.test.ts`

- [ ] **Step 1: Write failing tests for load/sync constraints and property access**

Create `tests/range.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { MockRange } from "../src/range";
import { CellStorage } from "../src/cell-storage";

describe("MockRange", () => {
  function createRange(address: string, storage?: CellStorage): {
    range: MockRange;
    storage: CellStorage;
    sync: () => Promise<void>;
  } {
    const s = storage ?? new CellStorage();
    const pendingLoads: MockRange[] = [];
    const range = new MockRange("Sheet1", address, s, pendingLoads);
    const sync = async () => {
      for (const r of pendingLoads) r.resolveLoads(s);
      pendingLoads.length = 0;
    };
    return { range, storage: s, sync };
  }

  test("accessing values without load throws", () => {
    const { range } = createRange("A1");
    expect(() => range.values).toThrow();
  });

  test("accessing values after load but before sync throws", () => {
    const { range } = createRange("A1");
    range.load("values");
    expect(() => range.values).toThrow();
  });

  test("accessing values after load and sync returns value", async () => {
    const { range, storage, sync } = createRange("A1");
    storage.setValue("Sheet1", "A1", 42);
    range.load("values");
    await sync();
    expect(range.values).toEqual([[42]]);
  });

  test("load accepts comma-separated string", async () => {
    const { range, storage, sync } = createRange("A1");
    storage.setValue("Sheet1", "A1", 42);
    range.load("values, formulas");
    await sync();
    expect(range.values).toEqual([[42]]);
    expect(range.formulas).toEqual([[42]]); // no formula, returns value
  });

  test("load accepts array of strings", async () => {
    const { range, storage, sync } = createRange("A1");
    storage.setValue("Sheet1", "A1", 42);
    range.load(["values", "address"]);
    await sync();
    expect(range.values).toEqual([[42]]);
    expect(range.address).toBe("Sheet1!A1");
  });

  test("address property returns sheet-qualified address", async () => {
    const { range, sync } = createRange("B3:D5");
    range.load("address");
    await sync();
    expect(range.address).toBe("Sheet1!B3:D5");
  });

  test("rowCount and columnCount for range", async () => {
    const { range, sync } = createRange("A1:C3");
    range.load(["rowCount", "columnCount"]);
    await sync();
    expect(range.rowCount).toBe(3);
    expect(range.columnCount).toBe(3);
  });

  test("rowIndex and columnIndex for range", async () => {
    const { range, sync } = createRange("C2");
    range.load(["rowIndex", "columnIndex"]);
    await sync();
    expect(range.rowIndex).toBe(1);
    expect(range.columnIndex).toBe(2);
  });

  test("values for multi-cell range", async () => {
    const { range, storage, sync } = createRange("A1:B2");
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet1", "B1", 2);
    storage.setValue("Sheet1", "A2", 3);
    storage.setValue("Sheet1", "B2", 4);
    range.load("values");
    await sync();
    expect(range.values).toEqual([[1, 2], [3, 4]]);
  });

  test("writing values does not require load", () => {
    const { range } = createRange("A1");
    expect(() => { range.values = [[42]]; }).not.toThrow();
  });

  test("getCell returns a MockRange for a single cell", async () => {
    const { range, storage, sync } = createRange("A1:C3");
    storage.setValue("Sheet1", "B2", 99);
    const cell = range.getCell(1, 1); // row 1, col 1 within range → B2
    cell.load("values");
    await sync();
    expect(cell.values).toEqual([[99]]);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/range.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement MockRange**

Create `src/range.ts`:

```typescript
import { parseAddress, cellAddressFromPosition, indexToColumnLetter } from "./address";
import { CellStorage } from "./cell-storage";

type WriteOperation = {
  type: "values" | "formulas" | "clear";
  data: unknown[][];
};

export class MockRange {
  private _sheetName: string;
  private _address: string;
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];
  private _loadedProperties = new Set<string>();
  private _syncedProperties = new Set<string>();
  private _cachedValues: Record<string, unknown> = {};
  private _writeQueue: WriteOperation[] = [];
  private _startRow: number;
  private _startCol: number;
  private _endRow: number;
  private _endCol: number;

  constructor(
    sheetName: string,
    address: string,
    storage: CellStorage,
    pendingLoads: MockRange[],
  ) {
    this._sheetName = sheetName;
    this._address = address;
    this._storage = storage;
    this._pendingLoads = pendingLoads;

    const parsed = parseAddress(address);
    this._startRow = parsed.startRow;
    this._startCol = parsed.startCol;
    this._endRow = parsed.endRow;
    this._endCol = parsed.endCol;
  }

  load(properties: string | string[]): MockRange {
    const props = typeof properties === "string"
      ? properties.split(",").map((p) => p.trim())
      : properties;

    for (const prop of props) {
      this._loadedProperties.add(prop);
    }
    this._pendingLoads.push(this);
    return this; // Real API returns Range for chaining
  }

  resolveLoads(storage: CellStorage): void {
    for (const prop of this._loadedProperties) {
      this._syncedProperties.add(prop);
      switch (prop) {
        case "values":
          this._cachedValues.values = this.readValues(storage);
          break;
        case "formulas":
          this._cachedValues.formulas = this.readFormulas(storage);
          break;
        case "address":
          this._cachedValues.address = this.computeAddress();
          break;
        case "rowCount":
          this._cachedValues.rowCount = this._endRow - this._startRow + 1;
          break;
        case "columnCount":
          this._cachedValues.columnCount = this._endCol - this._startCol + 1;
          break;
        case "rowIndex":
          this._cachedValues.rowIndex = this._startRow;
          break;
        case "columnIndex":
          this._cachedValues.columnIndex = this._startCol;
          break;
        case "text":
          this._cachedValues.text = this.readText(storage);
          break;
        case "numberFormat":
          this._cachedValues.numberFormat = this.readNumberFormat(storage);
          break;
        case "hasSpill":
          this._cachedValues.hasSpill = this.readHasSpill(storage);
          break;
      }
    }
    this._loadedProperties.clear();
  }

  private requireLoaded(prop: string): void {
    if (!this._syncedProperties.has(prop)) {
      throw new Error(
        `Property '${prop}' was not loaded. Call 'range.load("${prop}")' and 'context.sync()' first.`,
      );
    }
  }

  get values(): unknown[][] {
    this.requireLoaded("values");
    return this._cachedValues.values as unknown[][];
  }

  set values(data: unknown[][]) {
    this._writeQueue.push({ type: "values", data });
  }

  get formulas(): unknown[][] {
    this.requireLoaded("formulas");
    return this._cachedValues.formulas as unknown[][];
  }

  set formulas(data: unknown[][]) {
    this._writeQueue.push({ type: "formulas", data });
  }

  get address(): string {
    this.requireLoaded("address");
    return this._cachedValues.address as string;
  }

  get rowCount(): number {
    this.requireLoaded("rowCount");
    return this._cachedValues.rowCount as number;
  }

  get columnCount(): number {
    this.requireLoaded("columnCount");
    return this._cachedValues.columnCount as number;
  }

  get rowIndex(): number {
    this.requireLoaded("rowIndex");
    return this._cachedValues.rowIndex as number;
  }

  get columnIndex(): number {
    this.requireLoaded("columnIndex");
    return this._cachedValues.columnIndex as number;
  }

  get text(): string[][] {
    this.requireLoaded("text");
    return this._cachedValues.text as string[][];
  }

  get numberFormat(): unknown[][] {
    this.requireLoaded("numberFormat");
    return this._cachedValues.numberFormat as unknown[][];
  }

  set numberFormat(_data: unknown[][]) {
    // numberFormat write — no-op in mock (format storage not implemented)
  }

  get hasSpill(): boolean {
    this.requireLoaded("hasSpill");
    return this._cachedValues.hasSpill as boolean;
  }

  getCell(row: number, column: number): MockRange {
    const cellAddress = cellAddressFromPosition(
      this._startRow + row,
      this._startCol + column,
    );
    return new MockRange(this._sheetName, cellAddress, this._storage, this._pendingLoads);
  }

  clear(_applyTo?: string): void {
    // Queue clear as a write operation — executed on sync()
    this._writeQueue.push({ type: "clear" as any, data: [] });
    this._pendingLoads.push(this); // register for sync processing
  }

  executeClear(): void {
    for (let r = this._startRow; r <= this._endRow; r++) {
      for (let c = this._startCol; c <= this._endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        this._storage.clear(this._sheetName, addr);
      }
    }
  }

  getWriteQueue(): WriteOperation[] {
    return this._writeQueue;
  }

  clearWriteQueue(): void {
    this._writeQueue = [];
  }

  getSheetName(): string {
    return this._sheetName;
  }

  getStartRow(): number {
    return this._startRow;
  }

  getStartCol(): number {
    return this._startCol;
  }

  private readValues(storage: CellStorage): unknown[][] {
    const rows: unknown[][] = [];
    for (let r = this._startRow; r <= this._endRow; r++) {
      const row: unknown[] = [];
      for (let c = this._startCol; c <= this._endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        row.push(storage.getCell(this._sheetName, addr).value);
      }
      rows.push(row);
    }
    return rows;
  }

  private readFormulas(storage: CellStorage): unknown[][] {
    const rows: unknown[][] = [];
    for (let r = this._startRow; r <= this._endRow; r++) {
      const row: unknown[] = [];
      for (let c = this._startCol; c <= this._endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        const cell = storage.getCell(this._sheetName, addr);
        // Real Excel: formulas returns the formula if present, otherwise the value
        row.push(cell.formula || cell.value);
      }
      rows.push(row);
    }
    return rows;
  }

  private readText(storage: CellStorage): string[][] {
    const rows: string[][] = [];
    for (let r = this._startRow; r <= this._endRow; r++) {
      const row: string[] = [];
      for (let c = this._startCol; c <= this._endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        row.push(String(storage.getCell(this._sheetName, addr).value));
      }
      rows.push(row);
    }
    return rows;
  }

  private readNumberFormat(_storage: CellStorage): unknown[][] {
    // Default number format — simplified
    const rows: unknown[][] = [];
    for (let r = this._startRow; r <= this._endRow; r++) {
      const row: unknown[] = [];
      for (let c = this._startCol; c <= this._endCol; c++) {
        row.push("General");
      }
      rows.push(row);
    }
    return rows;
  }

  private readHasSpill(storage: CellStorage): boolean {
    for (let r = this._startRow; r <= this._endRow; r++) {
      for (let c = this._startCol; c <= this._endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        const cell = storage.getCell(this._sheetName, addr);
        if (cell.spilledFrom) return true;
      }
    }
    return false;
  }

  private computeAddress(): string {
    const start = `${indexToColumnLetter(this._startCol)}${this._startRow + 1}`;
    if (this._startRow === this._endRow && this._startCol === this._endCol) {
      return `${this._sheetName}!${start}`;
    }
    const end = `${indexToColumnLetter(this._endCol)}${this._endRow + 1}`;
    return `${this._sheetName}!${start}:${end}`;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/range.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/range.ts tests/range.test.ts
git commit -m "feat: add MockRange with load/sync constraints and property access"
```

---

## Task 8: MockWorksheet, MockWorksheetCollection, MockWorkbook

**Files:**
- Create: `src/worksheet.ts`
- Create: `src/worksheet-collection.ts`
- Create: `src/workbook.ts`
- Create: `tests/workbook.test.ts`

- [ ] **Step 1: Write failing tests**

Create `tests/workbook.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { MockWorkbook } from "../src/workbook";
import { MockWorksheetCollection } from "../src/worksheet-collection";
import { CellStorage } from "../src/cell-storage";
import { MockRange } from "../src/range";

describe("MockWorkbook / MockWorksheetCollection / MockWorksheet", () => {
  function createWorkbook() {
    const storage = new CellStorage();
    const pendingLoads: MockRange[] = [];
    const worksheets = new MockWorksheetCollection(storage, pendingLoads);
    const workbook = new MockWorkbook(storage, pendingLoads, worksheets);
    return { workbook, storage, pendingLoads };
  }

  test("default state has Sheet1 as active worksheet", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getActiveWorksheet();
    expect(sheet).toBeDefined();
  });

  test("getRange returns a MockRange", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1");
    expect(range).toBeInstanceOf(MockRange);
  });

  test("getItem returns worksheet by name", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getItem("Sheet1");
    expect(sheet).toBeDefined();
  });

  test("getItem throws for unknown sheet", () => {
    const { workbook } = createWorkbook();
    expect(() => workbook.worksheets.getItem("Unknown")).toThrow();
  });

  test("add creates a new worksheet", () => {
    const { workbook } = createWorkbook();
    workbook.worksheets.add("Sheet2");
    const sheet = workbook.worksheets.getItem("Sheet2");
    expect(sheet).toBeDefined();
  });

  test("getSelectedRange returns the set selected range", () => {
    const { workbook } = createWorkbook();
    workbook.setSelectedRange("Sheet1", "B2");
    const range = workbook.getSelectedRange();
    expect(range).toBeInstanceOf(MockRange);
  });

  test("getSelectedRange throws when no selection is set", () => {
    const { workbook } = createWorkbook();
    expect(() => workbook.getSelectedRange()).toThrow();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/workbook.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement MockWorksheet**

Create `src/worksheet.ts`:

```typescript
import { CellStorage } from "./cell-storage";
import { MockRange } from "./range";

export class MockWorksheet {
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];

  // Properties matching real Excel.Worksheet — no load/sync required for name/id
  // since they are always available on worksheet objects in real Excel
  readonly name: string;
  readonly id: string;

  constructor(name: string, id: string, storage: CellStorage, pendingLoads: MockRange[]) {
    this.name = name;
    this.id = id;
    this._storage = storage;
    this._pendingLoads = pendingLoads;
  }

  getRange(address: string): MockRange {
    return new MockRange(this.name, address, this._storage, this._pendingLoads);
  }
}
```

- [ ] **Step 4: Implement MockWorksheetCollection**

Create `src/worksheet-collection.ts`:

```typescript
import { CellStorage } from "./cell-storage";
import { MockWorksheet } from "./worksheet";
import { MockRange } from "./range";

export class MockWorksheetCollection {
  private _worksheets = new Map<string, MockWorksheet>();
  private _activeWorksheetName: string;
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];
  private _nextId = 1;

  constructor(storage: CellStorage, pendingLoads: MockRange[]) {
    this._storage = storage;
    this._pendingLoads = pendingLoads;
    this.add("Sheet1");
    this._activeWorksheetName = "Sheet1";
  }

  getActiveWorksheet(): MockWorksheet {
    return this._worksheets.get(this._activeWorksheetName)!;
  }

  setActiveWorksheet(name: string): void {
    if (!this._worksheets.has(name)) {
      throw new Error(`Worksheet '${name}' not found.`);
    }
    this._activeWorksheetName = name;
  }

  getItem(name: string): MockWorksheet {
    const sheet = this._worksheets.get(name);
    if (!sheet) {
      throw new Error(`Worksheet '${name}' not found.`);
    }
    return sheet;
  }

  add(name: string): MockWorksheet {
    const id = `{${String(this._nextId++).padStart(8, "0")}}`;
    const sheet = new MockWorksheet(name, id, this._storage, this._pendingLoads);
    this._worksheets.set(name, sheet);
    return sheet;
  }

  reset(): void {
    this._worksheets.clear();
    this._nextId = 1;
    this.add("Sheet1");
    this._activeWorksheetName = "Sheet1";
  }
}
```

- [ ] **Step 5: Implement MockWorkbook**

Create `src/workbook.ts`:

MockWorkbook accepts an existing MockWorksheetCollection so that state is shared across Excel.run calls. The ExcelMock wrapper owns the collection and passes it to each MockRequestContext.

```typescript
import { CellStorage } from "./cell-storage";
import { MockWorksheetCollection } from "./worksheet-collection";
import { MockRange } from "./range";

export class MockWorkbook {
  readonly worksheets: MockWorksheetCollection;
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];
  private _selectedSheet?: string;
  private _selectedAddress?: string;

  constructor(
    storage: CellStorage,
    pendingLoads: MockRange[],
    worksheets: MockWorksheetCollection,
  ) {
    this._storage = storage;
    this._pendingLoads = pendingLoads;
    this.worksheets = worksheets;
  }

  getSelectedRange(): MockRange {
    if (!this._selectedSheet || !this._selectedAddress) {
      throw new Error("No range is currently selected. Call setSelectedRange() first.");
    }
    return new MockRange(this._selectedSheet, this._selectedAddress, this._storage, this._pendingLoads);
  }

  setSelectedRange(sheet: string, address: string): void {
    this._selectedSheet = sheet;
    this._selectedAddress = address;
  }

  resetSelection(): void {
    this._selectedSheet = undefined;
    this._selectedAddress = undefined;
  }
}
```

- [ ] **Step 6: Run test to verify it passes**

```bash
npx vitest run tests/workbook.test.ts
```

Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add src/worksheet.ts src/worksheet-collection.ts src/workbook.ts tests/workbook.test.ts
git commit -m "feat: add MockWorksheet, MockWorksheetCollection, and MockWorkbook"
```

---

## Task 9: MockRequestContext and Excel.run with Write Queue

**Files:**
- Create: `src/request-context.ts`
- Create: `tests/request-context.test.ts`

- [ ] **Step 1: Write failing tests**

Create `tests/request-context.test.ts`:

```typescript
import { describe, expect, test } from "vitest";
import { MockRequestContext } from "../src/request-context";
import { CellStorage } from "../src/cell-storage";
import { MockCustomFunctions } from "../src/custom-functions-mock";
import { MockWorksheetCollection } from "../src/worksheet-collection";
import { MockRange } from "../src/range";

describe("MockRequestContext", () => {
  function createContext() {
    const storage = new CellStorage();
    const cf = new MockCustomFunctions();
    const pendingLoads: MockRange[] = [];
    const worksheets = new MockWorksheetCollection(storage, pendingLoads);
    const context = new MockRequestContext(storage, cf, worksheets);
    return { context, storage, cf };
  }

  test("context.workbook is accessible", () => {
    const { context } = createContext();
    expect(context.workbook).toBeDefined();
    expect(context.workbook.worksheets).toBeDefined();
  });

  test("sync resolves pending loads", async () => {
    const { context, storage } = createContext();
    storage.setValue("Sheet1", "A1", 42);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.load("values");
    await context.sync();
    expect(range.values).toEqual([[42]]);
  });

  test("sync executes queued value writes", async () => {
    const { context, storage } = createContext();
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.values = [[99]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe(99);
  });

  test("sync executes queued formula writes and evaluates custom functions", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("ADD", (a: number, b: number) => a + b);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=ADD(1, 2)"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe(3);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("=ADD(1, 2)");
  });

  test("sync evaluates async custom functions", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("ASYNC_ADD", async (a: number, b: number) => a + b);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=ASYNC_ADD(10, 20)"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe(30);
  });

  test("unregistered function formula sets #NAME? value", async () => {
    const { context, storage } = createContext();
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=UNKNOWN(1)"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

  test("formula evaluation passes Invocation with address and functionName", async () => {
    const { context, cf } = createContext();
    let receivedInvocation: unknown;
    cf.associate("CAPTURE", (...args: unknown[]) => {
      receivedInvocation = args[args.length - 1];
      return 0;
    });
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("B3");
    range.formulas = [["=CAPTURE()"]];
    await context.sync();
    expect(receivedInvocation).toEqual({
      address: "Sheet1!B3",
      functionName: "CAPTURE",
    });
  });

  test("spilling formula writes to multiple cells", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("MATRIX", () => [[1, 2], [3, 4]]);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("B2");
    range.formulas = [["=MATRIX()"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "B2").value).toBe(1);
    expect(storage.getCell("Sheet1", "C2").value).toBe(2);
    expect(storage.getCell("Sheet1", "B3").value).toBe(3);
    expect(storage.getCell("Sheet1", "C3").value).toBe(4);
  });

  test("write then read in same run works after sync", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("DOUBLE", (n: number) => n * 2);
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const writeRange = sheet.getRange("A1");
    writeRange.formulas = [["=DOUBLE(5)"]];
    await context.sync();

    const readRange = sheet.getRange("A1");
    readRange.load("values");
    await context.sync();
    expect(readRange.values).toEqual([[10]]);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/request-context.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement MockRequestContext**

Create `src/request-context.ts`:

MockRequestContext accepts a shared MockWorksheetCollection (owned by ExcelMock) so worksheet state persists across Excel.run calls. Uses FormulaEvaluator for formula processing (shared with ExcelMock wrapper — no duplication).

```typescript
import { CellStorage } from "./cell-storage";
import { MockCustomFunctions } from "./custom-functions-mock";
import { MockWorksheetCollection } from "./worksheet-collection";
import { MockWorkbook } from "./workbook";
import { MockRange } from "./range";
import { FormulaEvaluator } from "./formula-evaluator";
import { cellAddressFromPosition } from "./address";

export class MockRequestContext {
  readonly workbook: MockWorkbook;
  private _storage: CellStorage;
  private _evaluator: FormulaEvaluator;
  private _pendingLoads: MockRange[] = [];

  constructor(
    storage: CellStorage,
    customFunctions: MockCustomFunctions,
    worksheets: MockWorksheetCollection,
  ) {
    this._storage = storage;
    this._evaluator = new FormulaEvaluator(storage, customFunctions);
    this.workbook = new MockWorkbook(storage, this._pendingLoads, worksheets);
  }

  async sync(): Promise<void> {
    // Process writes first, then resolve loads
    const processed = new Set<MockRange>();
    for (const range of [...this._pendingLoads]) {
      if (processed.has(range)) continue;
      processed.add(range);
      await this.processRangeWrites(range);
    }

    for (const range of this._pendingLoads) {
      range.resolveLoads(this._storage);
    }
    this._pendingLoads.length = 0;
  }

  private async processRangeWrites(range: MockRange): Promise<void> {
    const writes = range.getWriteQueue();
    for (const write of writes) {
      if (write.type === "values") {
        await this.writeValues(range, write.data);
      } else if (write.type === "formulas") {
        await this.writeFormulas(range, write.data);
      } else if (write.type === "clear") {
        range.executeClear();
      }
    }
    range.clearWriteQueue();
  }

  private async writeValues(range: MockRange, data: unknown[][]): Promise<void> {
    const sheetName = range.getSheetName();
    const startRow = range.getStartRow();
    const startCol = range.getStartCol();

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const addr = cellAddressFromPosition(startRow + r, startCol + c);
        this._storage.setValue(sheetName, addr, data[r][c]);
      }
    }
  }

  private async writeFormulas(range: MockRange, data: unknown[][]): Promise<void> {
    const sheetName = range.getSheetName();
    const startRow = range.getStartRow();
    const startCol = range.getStartCol();

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const addr = cellAddressFromPosition(startRow + r, startCol + c);
        await this._evaluator.evaluateAndStore(sheetName, addr, String(data[r][c]));
      }
    }
  }
}
```

Also update `MockRange` so that setting values/formulas registers the range for write processing. In `src/range.ts`, update the setters:

```typescript
  set values(data: unknown[][]) {
    this._writeQueue.push({ type: "values", data });
    this._pendingLoads.push(this); // register for sync processing
  }

  set formulas(data: unknown[][]) {
    this._writeQueue.push({ type: "formulas", data });
    this._pendingLoads.push(this); // register for sync processing
  }
```

- [ ] **Step 4: Run test to verify it passes**

```bash
npx vitest run tests/request-context.test.ts
```

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/request-context.ts tests/request-context.test.ts src/range.ts
git commit -m "feat: add MockRequestContext with sync, write queue, and formula evaluation"
```

---

## Task 10: ExcelMock Wrapper and Excel.run

**Files:**
- Create: `src/excel-mock.ts`
- Modify: `src/index.ts`
- Create: `tests/excel-mock.test.ts`

- [ ] **Step 1: Write failing tests**

Create `tests/excel-mock.test.ts`:

```typescript
import { describe, expect, test, vi, afterEach, afterAll, beforeAll } from "vitest";
import { ExcelMock } from "../src/excel-mock";

describe("ExcelMock", () => {
  const mock = new ExcelMock();

  beforeAll(() => {
    vi.stubGlobal("Excel", mock.excel);
    vi.stubGlobal("CustomFunctions", mock.customFunctions);
  });

  afterEach(() => mock.reset());
  afterAll(() => vi.unstubAllGlobals());

  test("registered function returns correct value via setCell", async () => {
    mock.customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mock.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("setCell with plain value", () => {
    mock.setCell("Sheet1", "A1", 42);
    expect(mock.getCell("Sheet1", "A1").value).toBe(42);
  });

  test("uninitialized cell returns empty strings", () => {
    const cell = mock.getCell("Sheet1", "Z99");
    expect(cell.value).toBe("");
    expect(cell.formula).toBe("");
  });

  test("Excel.run works with context", async () => {
    mock.setCell("Sheet1", "A1", 42);

    await mock.excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });

  test("Excel.run formula write evaluates custom function", async () => {
    mock.customFunctions.associate("DOUBLE", (n: number) => n * 2);

    await mock.excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.formulas = [["=DOUBLE(21)"]];
      await context.sync();
    });

    expect(mock.getCell("Sheet1", "A1").value).toBe(42);
  });

  test("task pane E2E: read cell, write formula to selected range", async () => {
    mock.customFunctions.associate("DOUBLE", (n: number) => n * 2);
    mock.setCell("Sheet1", "A1", 5);
    mock.setSelectedRange("Sheet1", "B1");

    await mock.excel.run(async (context: any) => {
      const source = context.workbook.worksheets
        .getActiveWorksheet().getRange("A1");
      source.load("values");
      await context.sync();

      const selected = context.workbook.getSelectedRange();
      selected.formulas = [["=DOUBLE(5)"]];
      await context.sync();
    });

    expect(mock.getCell("Sheet1", "B1").value).toBe(10);
    expect(mock.getCell("Sheet1", "B1").formula).toBe("=DOUBLE(5)");
  });

  test("spill collision returns #SPILL!", async () => {
    mock.customFunctions.associate("MATRIX", () => [[1, 2], [3, 4]]);
    mock.setCell("Sheet1", "C2", 999);
    await mock.setCell("Sheet1", "B2", { formula: "=MATRIX()" });
    expect(mock.getCell("Sheet1", "B2").value).toBe("#SPILL!");
  });

  test("function name lookup is case-insensitive", async () => {
    mock.customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mock.setCell("Sheet1", "A1", { formula: "=add(1, 2)" });
    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("reset clears all state", async () => {
    mock.customFunctions.associate("ADD", (a: number, b: number) => a + b);
    mock.setCell("Sheet1", "A1", 42);
    mock.addWorksheet("Sheet2");

    mock.reset();

    expect(mock.getCell("Sheet1", "A1").value).toBe("");
    expect(mock.customFunctions.getFunction("ADD")).toBeUndefined();
    expect(() => mock.excel.run(async (context: any) => {
      context.workbook.worksheets.getItem("Sheet2");
    })).rejects.toThrow();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run tests/excel-mock.test.ts
```

Expected: FAIL

- [ ] **Step 3: Implement ExcelMock**

Create `src/excel-mock.ts`:

ExcelMock owns the shared CellStorage, MockWorksheetCollection, and MockCustomFunctions. Each Excel.run call gets a new MockRequestContext but shares the same worksheet and cell state. Uses FormulaEvaluator for formula processing (same as MockRequestContext — no duplication).

```typescript
import { CellStorage, CellState } from "./cell-storage";
import { MockCustomFunctions } from "./custom-functions-mock";
import { MockWorksheetCollection } from "./worksheet-collection";
import { MockRequestContext } from "./request-context";
import { FormulaEvaluator } from "./formula-evaluator";
import { parseAddress, cellAddressFromPosition, parseCellAddress } from "./address";
import { MockRange } from "./range";

export class ExcelMock {
  private _storage = new CellStorage();
  private _worksheets: MockWorksheetCollection;
  private _evaluator: FormulaEvaluator;
  readonly customFunctions = new MockCustomFunctions();

  private _selectedSheet?: string;
  private _selectedAddress?: string;

  constructor() {
    const dummyLoads: MockRange[] = [];
    this._worksheets = new MockWorksheetCollection(this._storage, dummyLoads);
    this._evaluator = new FormulaEvaluator(this._storage, this.customFunctions);
  }

  readonly excel = {
    run: async <T>(callback: (context: MockRequestContext) => Promise<T>): Promise<T> => {
      const context = new MockRequestContext(
        this._storage,
        this.customFunctions,
        this._worksheets,
      );
      if (this._selectedSheet && this._selectedAddress) {
        context.workbook.setSelectedRange(this._selectedSheet, this._selectedAddress);
      }
      return await callback(context);
    },
  };

  setCell(sheet: string, address: string, value: unknown): void | Promise<void> {
    if (typeof value === "object" && value !== null && "formula" in value) {
      return this._evaluator.evaluateAndStore(
        sheet,
        address,
        (value as { formula: string }).formula,
      );
    }
    this._storage.setValue(sheet, address, value);
  }

  setCells(sheet: string, startAddress: string, values: unknown[][]): void {
    const start = parseCellAddress(startAddress);
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const addr = cellAddressFromPosition(start.row + r, start.col + c);
        this._storage.setValue(sheet, addr, values[r][c]);
      }
    }
  }

  getCell(sheet: string, address: string): CellState {
    return this._storage.getCell(sheet, address);
  }

  getCells(sheet: string, rangeAddress: string): CellState[][] {
    const range = parseAddress(rangeAddress);
    const rows: CellState[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: CellState[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(this._storage.getCell(sheet, cellAddressFromPosition(r, c)));
      }
      rows.push(row);
    }
    return rows;
  }

  setSelectedRange(sheet: string, address: string): void {
    this._selectedSheet = sheet;
    this._selectedAddress = address;
  }

  setActiveWorksheet(sheet: string): void {
    this._worksheets.setActiveWorksheet(sheet);
  }

  addWorksheet(name: string): void {
    this._worksheets.add(name);
  }

  reset(): void {
    this._storage.clearAll();
    this.customFunctions.reset();
    this._worksheets.reset();
    this._selectedSheet = undefined;
    this._selectedAddress = undefined;
  }
}
```

- [ ] **Step 4: Update src/index.ts**

```typescript
export { ExcelMock } from "./excel-mock";
export type { CellState } from "./cell-storage";
```

- [ ] **Step 5: Run test to verify it passes**

```bash
npx vitest run tests/excel-mock.test.ts
```

Expected: PASS

- [ ] **Step 6: Commit**

```bash
git add src/excel-mock.ts src/index.ts tests/excel-mock.test.ts
git commit -m "feat: add ExcelMock wrapper with Excel.run and convenience API"
```

---

## Task 11: Type Conformance Tests

**Files:**
- Create: `tests/range.test-d.ts`
- Create: `tests/workbook.test-d.ts`
- Create: `tests/custom-functions.test-d.ts`

- [ ] **Step 1: Create type test for MockRange**

Create `tests/range.test-d.ts`:

```typescript
import { expectTypeOf } from "vitest";
import { MockRange } from "../src/range";

declare const mockRange: MockRange;

type ImplementedRange = Pick<
  Excel.Range,
  | "values"
  | "formulas"
  | "address"
  | "rowCount"
  | "columnCount"
  | "columnIndex"
  | "rowIndex"
  | "text"
  | "numberFormat"
  | "hasSpill"
  | "getCell"
  | "clear"
>;

expectTypeOf(mockRange).toMatchTypeOf<ImplementedRange>();
```

- [ ] **Step 2: Create type test for MockWorkbook**

Create `tests/workbook.test-d.ts`:

```typescript
import { expectTypeOf } from "vitest";
import { MockWorkbook } from "../src/workbook";
import { MockWorksheetCollection } from "../src/worksheet-collection";
import { MockWorksheet } from "../src/worksheet";

declare const mockWorkbook: MockWorkbook;
declare const mockCollection: MockWorksheetCollection;
declare const mockWorksheet: MockWorksheet;

type ImplementedWorkbook = Pick<Excel.Workbook, "worksheets" | "getSelectedRange">;
expectTypeOf(mockWorkbook).toMatchTypeOf<ImplementedWorkbook>();

type ImplementedCollection = Pick<Excel.WorksheetCollection, "getActiveWorksheet" | "getItem" | "add">;
expectTypeOf(mockCollection).toMatchTypeOf<ImplementedCollection>();

type ImplementedWorksheet = Pick<Excel.Worksheet, "getRange" | "name" | "id">;
expectTypeOf(mockWorksheet).toMatchTypeOf<ImplementedWorksheet>();
```

- [ ] **Step 3: Create type test for MockCustomFunctions**

Create `tests/custom-functions.test-d.ts`:

```typescript
import { expectTypeOf } from "vitest";
import { MockCustomFunctions } from "../src/custom-functions-mock";

declare const cf: MockCustomFunctions;

// associate should accept the same signatures as CustomFunctions.associate
expectTypeOf(cf.associate).toBeCallableWith("ADD", (a: number) => a);
expectTypeOf(cf.associate).toBeCallableWith({ ADD: (a: number) => a });

// Error class
const error = new cf.Error(cf.ErrorCode.invalidValue, "msg");
expectTypeOf(error.code).toBeString();
expectTypeOf(error.message).toEqualTypeOf<string | undefined>();
```

- [ ] **Step 4: Run type tests**

```bash
npx vitest typecheck
```

Expected: PASS. If there are type mismatches, fix the mock types to match the real Excel API.

- [ ] **Step 5: Commit**

```bash
git add tests/range.test-d.ts tests/workbook.test-d.ts tests/custom-functions.test-d.ts
git commit -m "test: add type conformance tests against @types/office-js"
```

---

## Task 12: Integration Test — Full E2E Flow

**Files:**
- Create: `tests/integration.test.ts`

- [ ] **Step 1: Write E2E integration test**

Create `tests/integration.test.ts`:

```typescript
import { describe, expect, test, vi, afterEach, afterAll, beforeAll } from "vitest";
import { ExcelMock } from "../src/excel-mock";

describe("E2E integration", () => {
  const mock = new ExcelMock();

  beforeAll(() => {
    vi.stubGlobal("Excel", mock.excel);
    vi.stubGlobal("CustomFunctions", mock.customFunctions);
  });

  afterEach(() => mock.reset());
  afterAll(() => vi.unstubAllGlobals());

  test("full flow: register function, write formula via Excel.run, verify value", async () => {
    // 1. Register custom function
    CustomFunctions.associate("TRIPLE", (n: number) => n * 3);

    // 2. Set up initial cell data
    mock.setCell("Sheet1", "A1", 7);
    mock.setSelectedRange("Sheet1", "B1");

    // 3. Simulate task pane action: read A1, write formula to selected cell
    await Excel.run(async (context: any) => {
      const source = context.workbook.worksheets
        .getActiveWorksheet().getRange("A1");
      source.load("values");
      await context.sync();

      const val = source.values[0][0];
      const selected = context.workbook.getSelectedRange();
      selected.formulas = [[`=TRIPLE(${val})`]];
      await context.sync();
    });

    // 4. Verify via wrapper
    expect(mock.getCell("Sheet1", "B1").value).toBe(21);
    expect(mock.getCell("Sheet1", "B1").formula).toBe("=TRIPLE(7)");

    // 5. Verify via Excel.run
    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet().getRange("B1");
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

    await mock.setCell("Sheet1", "A1", { formula: "=TABLE()" });

    expect(mock.getCell("Sheet1", "A1").value).toBe("Name");
    expect(mock.getCell("Sheet1", "B1").value).toBe("Score");
    expect(mock.getCell("Sheet1", "A2").value).toBe("Alice");
    expect(mock.getCell("Sheet1", "B2").value).toBe(95);
    expect(mock.getCell("Sheet1", "A3").value).toBe("Bob");
    expect(mock.getCell("Sheet1", "B3").value).toBe(87);
  });

  test("load/sync enforcement catches missing load", async () => {
    mock.setCell("Sheet1", "A1", 42);

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet().getRange("A1");

      expect(() => range.values).toThrow();

      range.load("values");
      expect(() => range.values).toThrow();

      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });
});
```

- [ ] **Step 2: Run all tests**

```bash
npx vitest run
```

Expected: ALL PASS

- [ ] **Step 3: Commit**

```bash
git add tests/integration.test.ts
git commit -m "test: add E2E integration tests for full add-in workflow"
```

---

## Task 13: Final Cleanup

- [ ] **Step 1: Run all tests and typecheck together**

```bash
npx vitest run && npx vitest typecheck
```

Expected: ALL PASS

- [ ] **Step 2: Build**

```bash
npx tsc
```

Expected: No errors.

- [ ] **Step 3: Verify exports**

Check that `dist/index.js` and `dist/index.d.ts` exist and export `ExcelMock` and `CellState`.

- [ ] **Step 4: Commit any fixes**

```bash
git add -A
git commit -m "chore: final cleanup and build verification"
```
