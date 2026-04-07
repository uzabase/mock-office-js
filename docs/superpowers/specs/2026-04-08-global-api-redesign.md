# Global API Redesign

## Summary

Redesign mock-office-js to match real office.js loading behavior: a side-effect-only import that sets up globals, instead of requiring manual instantiation and global assignment.

## Motivation

Real office.js is loaded via `<script src="...office.js">` and automatically sets `window.Excel`, `window.Office`, `window.CustomFunctions` as globals. The mock should follow the same pattern for consistency. Currently, consumers must instantiate `ExcelMock` and manually assign globals, which diverges from the real API's usage model.

## Design

### Entry Point

`import "mock-office-js"` sets the following globals:

| Global | Content | Real office.js equivalent |
|---|---|---|
| `Excel` | `{ run: (cb) => Promise }` | Real `Excel` |
| `Office` | `{ onReady: (cb?) => void, actions: { associate: () => {} } }` | Real `Office` |
| `CustomFunctions` | `associate()`, `Error`, `ErrorCode` | Real `CustomFunctions` |
| `MockOfficeJs` | Test helpers (mock-specific) | No equivalent |

```ts
// src/index.ts (side-effect import)
import { createMockEnvironment } from "./setup.js";

const env = createMockEnvironment();
globalThis.Excel = env.excel;
globalThis.Office = env.office;
globalThis.CustomFunctions = env.customFunctions;
globalThis.MockOfficeJs = env.mockOfficeJs;
```

### Factory Function (`setup.ts`)

Constructs shared state and returns all global objects:

```ts
export function createMockEnvironment() {
  const storage = new CellStorage();
  const customFunctions = new MockCustomFunctions();
  const worksheets = new MockWorksheetCollection(storage, []);
  const evaluator = new FormulaEvaluator(storage, customFunctions);

  const excel = { run: async (cb) => { ... } };
  const office = { onReady: (cb?) => { ... }, actions: { ... } };

  const mockOfficeJs = {
    excel: {
      setCell, getCell, setCells, getCells,
      setSelectedRange, setActiveWorksheet, addWorksheet,
      loadFunctionsMetadata,
    },
    reset: () => { /* clear storage, customFunctions, worksheets */ },
  };

  return { excel, office, customFunctions, mockOfficeJs };
}
```

All objects share the same underlying state (CellStorage, MockCustomFunctions, MockWorksheetCollection). `MockOfficeJs.excel.setCell()` writes to the same storage that `Excel.run()` reads from.

### `MockOfficeJs` API Structure

```ts
MockOfficeJs.excel.setCell(sheet, address, value)        // Set a single cell
MockOfficeJs.excel.getCell(sheet, address)                // Get a single cell
MockOfficeJs.excel.setCells(sheet, startAddress, values)  // Set a range of cells
MockOfficeJs.excel.getCells(sheet, rangeAddress)          // Get a range of cells
MockOfficeJs.excel.setSelectedRange(sheet, address)       // Set selected range
MockOfficeJs.excel.setActiveWorksheet(sheet)              // Set active worksheet
MockOfficeJs.excel.addWorksheet(name)                     // Add a worksheet
MockOfficeJs.excel.loadFunctionsMetadata(url)             // Load custom functions metadata (Excel-specific)
MockOfficeJs.reset()                                       // Reset all state across all hosts
```

Excel-specific helpers are nested under `MockOfficeJs.excel` because:
- Custom Functions are Excel-only (manifest.xml defines them inside `<Host xsi:type="Workbook">`)
- `setCell`, `getCell`, etc. are Excel-specific operations
- Future host mocks (Word, PowerPoint) would get their own namespace (e.g., `MockOfficeJs.word`)

`reset()` is at the top level because it clears all shared state across all hosts.

### Type Definitions

`MockOfficeJs` is the only `declare global` from this package. Types for `Excel`, `Office`, `CustomFunctions` come from `@types/office-js` (declared as a peer dependency).

```ts
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

### Peer Dependencies

```json
{
  "peerDependencies": {
    "@types/office-js": "^1.0.0"
  }
}
```

This ensures consumers have the global types for `Excel`, `Office`, `CustomFunctions` without mock-office-js re-declaring them (which would cause type conflicts).

### Build Changes

- **`browser.ts` removed** — `index.ts` handles global setup for both ESM and IIFE builds
- **`tsdown.config.ts`** — Both ESM and IIFE use `index.ts` as entry point
- **`ExcelMock` class removed** — Replaced by `createMockEnvironment()` factory function

### Package Exports

Structure remains the same, but consumption changes:

```ts
// Before
import { ExcelMock } from "mock-office-js";
const mock = new ExcelMock();
globalThis.Excel = mock.excel;

// After
import "mock-office-js";
// Excel, Office, CustomFunctions, MockOfficeJs are now global
```

### Impact on Existing Code

| Change | Detail |
|---|---|
| `ExcelMock` class | Removed. Replaced by `createMockEnvironment()` |
| `src/index.ts` | Named exports → side-effect only (global setup) |
| `src/browser.ts` | Removed |
| `package.json` exports | Structure maintained, usage changes |
| `tsdown.config.ts` | IIFE entry point unified to `index.ts` |
| `declare global` | `MockOfficeJs` only. `Excel` etc. from `@types/office-js` |
| `peerDependencies` | `@types/office-js` added |
| Existing tests | Rewrite to use global `MockOfficeJs` instead of `ExcelMock` |

This is a breaking change, acceptable at version 0.0.x.
