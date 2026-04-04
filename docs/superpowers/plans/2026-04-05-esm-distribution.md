# ESM Distribution & Build Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Migrate the build configuration to produce standards-compliant ESM output that works via npm import and CDN `<script type="module">`.

**Architecture:** Change `tsconfig.json` to `module: "nodenext"` (enforces `.js` extensions), modernize `package.json` with `type: "module"` and `exports` field, then update all relative imports to use `.js` extensions.

**Tech Stack:** TypeScript 6, tsc (no bundler)

**Spec:** `docs/superpowers/specs/2026-04-05-office-js-mock-design.md` (Distribution & Build section)

---

### Task 1: Update `package.json`

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Update package.json**

```jsonc
{
  "name": "office-js-mock",
  "version": "0.1.0",
  "description": "In-memory mock of Excel JavaScript API for testing Excel Add-ins",
  "type": "module",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "default": "./dist/index.js"
    }
  },
  "files": ["dist"],
  "scripts": {
    "build": "tsc",
    "test": "vitest run",
    "test:watch": "vitest",
    "test:typecheck": "vitest --typecheck --run"
  },
  "license": "MIT",
  "devDependencies": {
    "@types/custom-functions-runtime": "^1.6.11",
    "@types/office-js": "^1.0.587",
    "typescript": "^6.0.2",
    "vitest": "^4.1.2"
  }
}
```

Changes from current:
- Add `"type": "module"`
- Add `"exports"` with `types` first (TS recommended ordering)
- Add `"files": ["dist"]`
- Remove `"main"` and `"types"` (superseded by `exports`)

- [ ] **Step 2: Commit**

```bash
git add package.json
git commit -m "chore: add type module, exports, and files to package.json"
```

---

### Task 2: Update `tsconfig.json`

**Files:**
- Modify: `tsconfig.json`

- [ ] **Step 1: Update tsconfig.json**

```jsonc
{
  "compilerOptions": {
    "target": "esnext",
    "module": "nodenext",
    "moduleResolution": "nodenext",
    "lib": ["esnext"],
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

Changes from current:
- `target`: `ES2020` → `esnext`
- `module`: `ES2020` → `nodenext`
- `moduleResolution`: `bundler` → `nodenext`
- `lib`: `["ES2020"]` → `["esnext"]`
- No `exclude` needed (tests are in `tests/`, outside `include`)

- [ ] **Step 2: Verify build fails (confirms nodenext enforcement)**

Run: `npm run build`
Expected: FAIL — errors about missing `.js` extensions on relative imports.

- [ ] **Step 3: Commit**

```bash
git add tsconfig.json
git commit -m "chore: migrate tsconfig to nodenext with esnext target"
```

---

### Task 3: Add `.js` extensions to source imports

**Files (9 files in `src/`):**
- Modify: `src/index.ts`
- Modify: `src/excel-mock.ts`
- Modify: `src/request-context.ts`
- Modify: `src/formula-evaluator.ts`
- Modify: `src/workbook.ts`
- Modify: `src/worksheet-collection.ts`
- Modify: `src/worksheet.ts`
- Modify: `src/cell-storage.ts`
- Modify: `src/range.ts`

Every relative import in these files needs `.js` appended. The full list:

- [ ] **Step 1: Update `src/index.ts`**

```ts
export { ExcelMock } from "./excel-mock.js";
export type { CellState } from "./cell-storage.js";
```

- [ ] **Step 2: Update `src/excel-mock.ts`**

```ts
import { CellStorage, CellState } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { MockWorksheetCollection } from "./worksheet-collection.js";
import { MockRequestContext } from "./request-context.js";
import { FormulaEvaluator } from "./formula-evaluator.js";
import { parseAddress, cellAddressFromPosition, parseCellAddress } from "./address.js";
import { MockRange } from "./range.js";
```

- [ ] **Step 3: Update `src/request-context.ts`**

```ts
import { CellStorage } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { MockWorksheetCollection } from "./worksheet-collection.js";
import { MockWorkbook } from "./workbook.js";
import { MockRange } from "./range.js";
import { FormulaEvaluator } from "./formula-evaluator.js";
import { cellAddressFromPosition } from "./address.js";
```

- [ ] **Step 4: Update `src/formula-evaluator.ts`**

```ts
import { CellStorage } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { parseFormula } from "./formula-parser.js";
```

- [ ] **Step 5: Update `src/workbook.ts`**

```ts
import { CellStorage } from "./cell-storage.js";
import { MockWorksheetCollection } from "./worksheet-collection.js";
import { MockRange } from "./range.js";
```

- [ ] **Step 6: Update `src/worksheet-collection.ts`**

```ts
import { CellStorage } from "./cell-storage.js";
import { MockWorksheet } from "./worksheet.js";
import { MockRange } from "./range.js";
```

- [ ] **Step 7: Update `src/worksheet.ts`**

```ts
import { CellStorage } from "./cell-storage.js";
import { MockRange } from "./range.js";
```

- [ ] **Step 8: Update `src/cell-storage.ts`**

```ts
import { cellAddressFromPosition, parseCellAddress } from "./address.js";
```

- [ ] **Step 9: Update `src/range.ts`**

```ts
import {
  parseAddress,
  cellAddressFromPosition,
  indexToColumnLetter,
  type RangePosition,
} from "./address.js";
import type { CellStorage } from "./cell-storage.js";
```

- [ ] **Step 10: Verify build succeeds**

Run: `npm run build`
Expected: PASS — no errors, `dist/` contains `.js` and `.d.ts` files.

- [ ] **Step 11: Commit**

```bash
git add src/
git commit -m "chore: add .js extensions to all source imports for nodenext"
```

---

### Task 4: Add `.js` extensions to test imports

**Files (13 files in `tests/`):**
- Modify: `tests/address.test.ts`
- Modify: `tests/integration.test.ts`
- Modify: `tests/custom-functions-mock.test.ts`
- Modify: `tests/formula-parser.test.ts`
- Modify: `tests/custom-functions.test-d.ts`
- Modify: `tests/request-context.test.ts`
- Modify: `tests/range.test-d.ts`
- Modify: `tests/workbook.test.ts`
- Modify: `tests/workbook.test-d.ts`
- Modify: `tests/range.test.ts`
- Modify: `tests/cell-storage.test.ts`
- Modify: `tests/formula-evaluator.test.ts`
- Modify: `tests/excel-mock.test.ts`

- [ ] **Step 1: Update `tests/address.test.ts`**

```ts
import {
  columnLetterToIndex,
  indexToColumnLetter,
  parseAddress,
  parseCellAddress,
  resolveRangeAddresses,
} from "../src/address.js";
```

- [ ] **Step 2: Update `tests/integration.test.ts`**

```ts
import { ExcelMock } from "../src/excel-mock.js";
```

- [ ] **Step 3: Update `tests/custom-functions-mock.test.ts`**

```ts
import { MockCustomFunctions } from "../src/custom-functions-mock.js";
```

- [ ] **Step 4: Update `tests/formula-parser.test.ts`**

```ts
import { parseFormula } from "../src/formula-parser.js";
```

- [ ] **Step 5: Update `tests/custom-functions.test-d.ts`**

```ts
import { MockCustomFunctions } from "../src/custom-functions-mock.js";
```

- [ ] **Step 6: Update `tests/request-context.test.ts`**

```ts
import { MockRequestContext } from "../src/request-context.js";
import { CellStorage } from "../src/cell-storage.js";
import { MockCustomFunctions } from "../src/custom-functions-mock.js";
import { MockWorksheetCollection } from "../src/worksheet-collection.js";
import { MockRange } from "../src/range.js";
```

- [ ] **Step 7: Update `tests/range.test-d.ts`**

```ts
import { MockRange } from "../src/range.js";
```

- [ ] **Step 8: Update `tests/workbook.test.ts`**

```ts
import { MockWorkbook } from "../src/workbook.js";
import { MockWorksheetCollection } from "../src/worksheet-collection.js";
import { CellStorage } from "../src/cell-storage.js";
import { MockRange } from "../src/range.js";
```

- [ ] **Step 9: Update `tests/workbook.test-d.ts`**

```ts
import { MockWorkbook } from "../src/workbook.js";
import { MockWorksheetCollection } from "../src/worksheet-collection.js";
import { MockWorksheet } from "../src/worksheet.js";
import { MockRange } from "../src/range.js";
```

- [ ] **Step 10: Update `tests/range.test.ts`**

```ts
import { MockRange } from "../src/range.js";
import { CellStorage } from "../src/cell-storage.js";
```

- [ ] **Step 11: Update `tests/cell-storage.test.ts`**

```ts
import { CellStorage } from "../src/cell-storage.js";
```

- [ ] **Step 12: Update `tests/formula-evaluator.test.ts`**

```ts
import { FormulaEvaluator } from "../src/formula-evaluator.js";
import { CellStorage } from "../src/cell-storage.js";
import { MockCustomFunctions } from "../src/custom-functions-mock.js";
```

- [ ] **Step 13: Update `tests/excel-mock.test.ts`**

```ts
import { ExcelMock } from "../src/excel-mock.js";
```

- [ ] **Step 14: Run all tests**

Run: `npm test`
Expected: All tests pass.

- [ ] **Step 15: Run type check tests**

Run: `npm run test:typecheck`
Expected: All type tests pass.

- [ ] **Step 16: Commit**

```bash
git add tests/
git commit -m "chore: add .js extensions to all test imports for nodenext"
```

---

### Task 5: Final verification

- [ ] **Step 1: Clean build**

```bash
rm -rf dist && npm run build
```

Expected: PASS — `dist/` populated with `.js` and `.d.ts` files.

- [ ] **Step 2: Verify dist output has .js extensions in imports**

```bash
head -5 dist/index.js
```

Expected: Imports with `.js` extensions (e.g., `from "./excel-mock.js"`).

- [ ] **Step 3: Verify all tests pass**

Run: `npm test`
Expected: All tests pass.

- [ ] **Step 4: Verify type tests pass**

Run: `npm run test:typecheck`
Expected: All type tests pass.
