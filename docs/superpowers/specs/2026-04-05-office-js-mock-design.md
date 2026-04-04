# office-js-mock Design Spec

## Overview

A mock library for testing Excel Add-ins without the Excel host. Provides an in-memory implementation of the Excel JavaScript API and CustomFunctions runtime, enabling E2E-style testing of add-in code.

## Goals

- Mock `Excel.run()`, `RequestContext`, `Workbook`, `Worksheet`, `Range` with in-memory cell storage
- Mock `CustomFunctions.associate()` and evaluate registered functions when formulas are set
- Reproduce `load()` / `sync()` constraints (error on unloaded property access)
- Support spilling (2D array results expanding to multiple cells)
- Provide a convenience wrapper (ExcelMock) for test setup and assertions

## Non-Goals

- Excel calculation engine (native functions like `=SUM()`, `=VLOOKUP()`)
- Event handling
- Streaming / cancelable custom functions (extensible, but not in initial scope)
- Cell reference resolution in formula arguments (extensible, but not in initial scope)
- UI rendering
- External API mocking (test-side responsibility, e.g., MSW)

## Architecture

### Two-Layer Design

```
ExcelMock (Wrapper - convenience API for tests)
  └── Core (Excel API mock implementation)
        ├── MockExcel.run()
        ├── MockRequestContext
        ├── MockWorkbook
        ├── MockWorksheetCollection
        ├── MockWorksheet
        ├── MockRange
        ├── MockCustomFunctions.associate()
        ├── FormulaParser
        └── CellStorage (in-memory state)
```

- **Core**: Reproduces the Excel JavaScript API. Can be used directly via `mock.excel.run(...)`.
- **Wrapper**: Provides `setCell()`, `getCell()`, `reset()` etc. for concise test code. Delegates to core internally.

### Project Structure

```
office-js-mock/
├── src/
│   ├── index.ts                  # Public API (exports ExcelMock)
│   ├── excel-mock.ts             # ExcelMock wrapper class
│   ├── custom-functions-mock.ts  # MockCustomFunctions
│   ├── request-context.ts        # MockRequestContext
│   ├── workbook.ts               # MockWorkbook
│   ├── worksheet.ts              # MockWorksheet
│   ├── worksheet-collection.ts   # MockWorksheetCollection
│   ├── range.ts                  # MockRange
│   ├── cell-storage.ts           # In-memory cell storage
│   └── formula-parser.ts         # Formula parsing and evaluation
└── tests/
    └── ...
```

## Core Design

### CellStorage

Central store for all cell state across all worksheets.

```typescript
interface CellState {
  value: unknown;           // Displayed value
  formula?: string;         // Formula string (if any)
  spilledFrom?: string;     // Address of the spill origin (for spilled cells)
}

// Structure: Map<sheetName, Map<address, CellState>>
```

**Write behavior:**

- `range.values = [[...]]`: Stores values directly (no formula).
- `range.formulas = [[...]]`: Parses the formula, looks up the registered custom function, calls it, and stores both formula and computed value. If the function is not registered, stores `"#NAME?"` as value.

**Spill behavior:**

When a custom function returns a 2D array:
- Origin cell: `{ formula, value: array[0][0] }`
- Spilled cells: `{ value: array[r][c], spilledFrom: originAddress }`
- Overwriting or clearing the origin clears all spilled cells.

### MockRange

Supports the following properties and methods:

**Properties (require `load()` before read):**

| Property | Type | Read | Write | Description |
|---|---|---|---|---|
| `values` | `any[][]` | load required | direct | Cell values |
| `formulas` | `any[][]` | load required | direct (evaluates custom functions) | Formulas |
| `address` | `string` | load required | readonly | `"Sheet1!A1:B2"` format |
| `rowCount` | `number` | load required | readonly | Row count |
| `columnCount` | `number` | load required | readonly | Column count |
| `hasSpill` | `boolean` | load required | readonly | Whether the range has spill |

**Methods:**

| Method | Description |
|---|---|
| `load(properties)` | Schedules property loading |
| `getCell(row, column)` | Returns a new MockRange for a specific cell |

**Address parsing:**

`getRange("A1")`, `getRange("A1:C2")` are parsed to determine the corresponding cells in CellStorage.

### load() / sync() Constraints

Strictly reproduces Office.js behavior:
- Accessing a property without `load()` → throws error
- Accessing a property after `load()` but before `sync()` → throws error
- After `load()` + `sync()` → returns the value from CellStorage

### MockWorkbook / MockWorksheetCollection / MockWorksheet

**Object chain:**
```
context.workbook.worksheets.getActiveWorksheet().getRange("A1")
```

**MockWorkbook:**
- `worksheets`: Returns MockWorksheetCollection

**MockWorksheetCollection:**
- `getActiveWorksheet()`: Returns the active worksheet
- `getItem(name)`: Returns worksheet by name
- `add(name)`: Adds a new worksheet

**MockWorksheet:**
- `getRange(address)`: Returns MockRange
- `getSelectedRange()`: Returns the currently selected range
- `name`: Sheet name (load required)

**MockRequestContext:**
- `workbook`: Returns MockWorkbook
- `sync()`: Resolves all pending load requests

**Excel.run:**
```typescript
mock.excel.run = async (callback) => {
  const context = new MockRequestContext(cellStorage, functionRegistry);
  return await callback(context);
};
```

**Default state:** One worksheet named "Sheet1", set as active (matches real Excel behavior).

### MockCustomFunctions

**Methods:**
- `associate(id, fn)`: Registers a single function
- `associate(mappings)`: Registers multiple functions via `{ id: fn }` object

**Function registry:** `Map<string, Function>` — used by CellStorage during formula evaluation.

**Invocation:** When evaluating a custom function, the mock constructs an `Invocation` object and passes it as the last argument:
```typescript
const invocation = { address: "Sheet1!B1" };
const result = await registeredFn(...parsedArgs, invocation);
```

Extensible for future `StreamingInvocation` and `CancelableInvocation` support.

### Formula Parser

Parses custom function formulas and evaluates them.

**Supported argument types (initial scope):**

| Type | Example | Parsed as |
|---|---|---|
| String literal | `"AAPL"` | `"AAPL"` |
| Number | `2024` | `2024` |
| Boolean | `TRUE`, `FALSE` | `true`, `false` |

**Unsupported (extensible via ArgumentResolver):**

| Type | Example | Required processing |
|---|---|---|
| Single cell ref | `A1` | Lookup value from CellStorage |
| Range ref | `A1:A3` | Lookup array from CellStorage |
| Cross-sheet ref | `Sheet2!A1` | Sheet name + address resolution |

**Extension point:**
```typescript
interface ArgumentResolver {
  resolve(token: string, context: { cellStorage: CellStorage; sheet: string }): unknown;
}
```

**Unresolvable formulas:**
- Unregistered function name → value is `"#NAME?"` (matches real Excel behavior)
- Parse failure → value is `"#NAME?"`

## Wrapper API (ExcelMock)

```typescript
class ExcelMock {
  // Mock objects for global setup
  readonly excel: MockExcel;
  readonly customFunctions: MockCustomFunctions;

  // Setup
  setCell(sheet: string, address: string, value: unknown): void;
  setCells(sheet: string, startAddress: string, values: unknown[][]): void;
  setSelectedRange(sheet: string, address: string): void;
  setActiveWorksheet(sheet: string): void;
  addWorksheet(name: string): void;

  // Assertion
  getCell(sheet: string, address: string): CellState;
  getCells(sheet: string, rangeAddress: string): CellState[][];

  // Reset
  reset(): void;
}
```

## Usage Example

```typescript
import { ExcelMock } from "office-js-mock";

describe("custom function add-in", () => {
  const mock = new ExcelMock();

  beforeAll(() => {
    vi.stubGlobal("Excel", mock.excel);
    vi.stubGlobal("CustomFunctions", mock.customFunctions);
  });

  afterEach(() => mock.reset());
  afterAll(() => vi.unstubAllGlobals());

  test("registered function returns correct value", async () => {
    CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

    mock.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });

    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("function returning 2D array spills to adjacent cells", async () => {
    CustomFunctions.associate("MATRIX", () => [
      [1, 2],
      [3, 4],
    ]);

    mock.setCell("Sheet1", "B2", { formula: "=MATRIX()" });

    expect(mock.getCell("Sheet1", "B2").value).toBe(1);
    expect(mock.getCell("Sheet1", "C2").value).toBe(2);
    expect(mock.getCell("Sheet1", "B3").value).toBe(3);
    expect(mock.getCell("Sheet1", "C3").value).toBe(4);
  });

  test("task pane writes formula to selected cell via Excel.run", async () => {
    CustomFunctions.associate("DOUBLE", (n: number) => n * 2);
    mock.setCell("Sheet1", "A1", 5);
    mock.setSelectedRange("Sheet1", "B1");

    // Production code: reads A1, writes formula to selected cell
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const source = sheet.getRange("A1");
      source.load("values");
      await context.sync();

      const selected = sheet.getSelectedRange();
      selected.formulas = [['=DOUBLE(5)']];
      await context.sync();
    });

    expect(mock.getCell("Sheet1", "B1").value).toBe(10);
    expect(mock.getCell("Sheet1", "B1").formula).toBe("=DOUBLE(5)");
  });

  test("unregistered function returns #NAME? error", async () => {
    mock.setCell("Sheet1", "A1", { formula: "=UNKNOWN(1)" });

    expect(mock.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

  test("load/sync constraint is enforced", async () => {
    mock.setCell("Sheet1", "A1", 42);

    await Excel.run(async (context) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet().getRange("A1");

      // Accessing without load throws
      expect(() => range.values).toThrow();

      range.load("values");

      // Accessing after load but before sync throws
      expect(() => range.values).toThrow();

      await context.sync();

      // After sync, value is available
      expect(range.values).toEqual([[42]]);
    });
  });
});
```

## Reference Files

- `@types/office-js`: `.references/@types/office-js/index.d.ts`
  - `Excel.Range`: line 38064
  - `Excel.Worksheet`: line 36824
  - `Excel.WorksheetCollection`: line 37445
  - `Excel.Workbook`: line 36295
  - `Excel.RequestContext`: line 33103
- `@types/custom-functions-runtime`: `.references/@types/custom-functions-runtime/index.d.ts`
- Existing mock library for reference: `.references/office-addin-mock/`
