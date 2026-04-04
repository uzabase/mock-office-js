# office-js-mock Design Spec

## Overview

A mock library for testing Excel Add-ins without the Excel host. Provides an in-memory implementation of the Excel JavaScript API and CustomFunctions runtime, enabling E2E-style testing of add-in code.

## Principle

**Always match real Excel behavior.** When in doubt, check `@types/office-js` and replicate the real API surface, method signatures, property types, and error behavior.

## Goals

- Mock `Excel.run()`, `RequestContext`, `Workbook`, `Worksheet`, `Range` with in-memory cell storage
- Mock `CustomFunctions.associate()`, `CustomFunctions.Error`, and `CustomFunctions.ErrorCode`
- Evaluate registered custom functions when formulas are set via `range.formulas`
- Reproduce `load()` / `sync()` constraints (error on unloaded property access)
- Support spilling (2D array results expanding to multiple cells) with `#SPILL!` error on collision
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
        ├── MockCustomFunctions (associate, Error, ErrorCode)
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
│   ├── custom-functions-mock.ts  # MockCustomFunctions (associate, Error, ErrorCode)
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
  value: unknown;           // Displayed value ("" for empty cells, matching real Excel)
  formula?: string;         // Formula string (if any; "" for empty cells, matching real Excel)
  spilledFrom?: string;     // Address of the spill origin (for spilled cells)
}

// Structure: Map<sheetName, Map<address, CellState>>
```

**Default cell state:** Uninitialized cells return `{ value: "", formula: "" }` matching real Excel behavior where `Range.values` returns `""` and `Range.formulas` returns `""` for empty cells.

**Write behavior:**

Writes are queued and executed on `context.sync()`, matching real Excel's batch execution model.

- `range.values = [[...]]`: Queues storing values directly (no formula).
- `range.formulas = [[...]]`: Queues formula write. On sync, parses the formula, looks up the registered custom function, calls it, and stores both formula and computed value. If the function is not registered, stores `"#NAME?"` as value.

**Spill behavior:**

When a custom function returns a 2D array:
- Origin cell: `{ formula, value: array[0][0] }`
- Spilled cells: `{ value: array[r][c], spilledFrom: originAddress }`
- Overwriting or clearing the origin clears all spilled cells.

**Spill collision:** If a spilling formula would write into cells that already contain data, the origin cell gets `value: "#SPILL!"` and no spill occurs. This matches real Excel behavior.

**Overwriting a spilled (non-origin) cell:** If a value is written to a cell that is part of another cell's spill range, the spill origin cell gets `value: "#SPILL!"` and the remaining spill cells are cleared. This matches real Excel behavior.

### MockRange

Supports the following properties and methods, matching the real `Excel.Range` API.

**Properties (require `load()` before read):**

| Property | Type | Read | Write | Description |
|---|---|---|---|---|
| `values` | `any[][]` | load required | direct (queued for sync) | Cell values |
| `formulas` | `any[][]` | load required | direct (queued for sync, evaluates custom functions) | Formulas in A1-style notation |
| `address` | `string` | load required | readonly | `"Sheet1!A1:B2"` format |
| `rowCount` | `number` | load required | readonly | Row count |
| `columnCount` | `number` | load required | readonly | Column count |
| `columnIndex` | `number` | load required | readonly | Zero-indexed column of first cell |
| `rowIndex` | `number` | load required | readonly | Zero-indexed row of first cell |
| `text` | `string[][]` | load required | readonly | Text representation of values |
| `numberFormat` | `any[][]` | load required | direct | Number format codes |
| `hasSpill` | `boolean` | load required | readonly | True if all cells have spill border, false if none, null if mixed |

**Methods:**

| Method | Description |
|---|---|
| `load(properties)` | Schedules property loading. Accepts string (`"values"`, `"values, formulas"`), string array (`["values", "formulas"]`), or object notation. |
| `getCell(row, column)` | Returns a new MockRange for a specific cell within the range |
| `clear(applyTo?)` | Clears the range. `applyTo` accepts `"All"` (default), `"Formats"`, `"Contents"`, etc. Clearing a spill origin clears all spilled cells. |

**Address parsing:**

`getRange("A1")`, `getRange("A1:C2")` are parsed to determine the corresponding cells in CellStorage. Multi-letter columns (e.g., `"AA1"`) are supported. Sheet-qualified addresses (e.g., `"'Sheet 1'!A1"`) are supported for sheet names with spaces.

### load() / sync() Constraints

Strictly reproduces Office.js behavior:
- Accessing a property without `load()` → throws error
- Accessing a property after `load()` but before `sync()` → throws error
- After `load()` + `sync()` → returns the value from CellStorage
- Writing properties (e.g., `range.formulas = [...]`) does not require load/sync; writes are queued and executed on sync.

### MockWorkbook / MockWorksheetCollection / MockWorksheet

**Object chain (matches real Excel API):**
```
context.workbook.worksheets.getActiveWorksheet().getRange("A1")
context.workbook.getSelectedRange()
```

**MockWorkbook:**

| Method/Property | Description |
|---|---|
| `worksheets` | Returns MockWorksheetCollection |
| `getSelectedRange()` | Returns the currently selected range (this is a Workbook method in real Excel API, NOT Worksheet) |

**MockWorksheetCollection:**

| Method | Description |
|---|---|
| `getActiveWorksheet()` | Returns the active worksheet |
| `getItem(name)` | Returns worksheet by name |
| `add(name)` | Adds a new worksheet |

**MockWorksheet:**

| Method/Property | Description |
|---|---|
| `getRange(address)` | Returns MockRange for the given address |
| `name` | Sheet name (load required) |
| `id` | Sheet id (load required) |

**MockRequestContext:**

| Method/Property | Description |
|---|---|
| `workbook` | Returns MockWorkbook |
| `sync()` | Executes all queued writes, then resolves all pending load requests |

**Excel.run:**
```typescript
mock.excel.run = async (callback) => {
  const context = new MockRequestContext(cellStorage, functionRegistry);
  return await callback(context);
};
```

**Default state:** One worksheet named "Sheet1", set as active (matches real Excel new workbook behavior).

### MockCustomFunctions

**Methods and types:**
- `associate(id, fn)`: Registers a single function
- `associate(mappings)`: Registers multiple functions via `{ id: fn }` object
- `Error`: Class with `code` and optional `message` properties
- `ErrorCode`: Enum with `invalidValue`, `notAvailable`, `divisionByZero`, `invalidNumber`, `nullReference`, `invalidName`, `invalidReference`

**Function registry:** `Map<string, Function>` — function ID lookup is **case-insensitive**, matching real Excel behavior. `associate("ADD", fn)` is matched by formula `=add(1,2)`.

**Invocation:** When evaluating a custom function, the mock constructs an `Invocation` object and passes it as the last argument:
```typescript
const invocation: CustomFunctions.Invocation = {
  address: "Sheet1!B1",
  functionName: "GETPRICE",
};
const result = await registeredFn(...parsedArgs, invocation);
```

Extensible for future `StreamingInvocation` and `CancelableInvocation` support. The Invocation construction is isolated so additional properties can be added without changing the evaluation flow.

**Namespace-prefixed function names:** Custom functions are typically invoked as `=CONTOSO.ADD(1,2)` with a namespace prefix. The formula parser handles dot-separated names. `associate("CONTOSO.ADD", fn)` matches formula `=CONTOSO.ADD(1,2)`.

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

**Async evaluation:** Custom functions may be async (return a Promise). Formula evaluation uses `await` internally. The wrapper's `setCell` with formula is therefore `async` and returns `Promise<void>`.

## Wrapper API (ExcelMock)

```typescript
class ExcelMock {
  // Mock objects for global setup
  readonly excel: MockExcel;
  readonly customFunctions: MockCustomFunctions;

  // Setup
  setCell(sheet: string, address: string, value: unknown): Promise<void>;
  // When value is a primitive (string, number, boolean), stores it as a plain value.
  // When value is { formula: string }, parses and evaluates the formula.

  setCells(sheet: string, startAddress: string, values: unknown[][]): void;
  setSelectedRange(sheet: string, address: string): void;
  setActiveWorksheet(sheet: string): void;
  addWorksheet(name: string): void;

  // Assertion
  getCell(sheet: string, address: string): CellState;
  getCells(sheet: string, rangeAddress: string): CellState[][];

  // Reset
  reset(): void;
  // Clears all cell data across all sheets.
  // Removes added worksheets and restores the default "Sheet1".
  // Clears the custom function registry.
  // Resets selected range to undefined.
  // Resets active worksheet to "Sheet1".
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

    await mock.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });

    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("function returning 2D array spills to adjacent cells", async () => {
    CustomFunctions.associate("MATRIX", () => [
      [1, 2],
      [3, 4],
    ]);

    await mock.setCell("Sheet1", "B2", { formula: "=MATRIX()" });

    expect(mock.getCell("Sheet1", "B2").value).toBe(1);
    expect(mock.getCell("Sheet1", "C2").value).toBe(2);
    expect(mock.getCell("Sheet1", "B3").value).toBe(3);
    expect(mock.getCell("Sheet1", "C3").value).toBe(4);
  });

  test("spill collision returns #SPILL! error", async () => {
    CustomFunctions.associate("MATRIX", () => [
      [1, 2],
      [3, 4],
    ]);

    mock.setCell("Sheet1", "C2", 999);  // obstacle
    await mock.setCell("Sheet1", "B2", { formula: "=MATRIX()" });

    expect(mock.getCell("Sheet1", "B2").value).toBe("#SPILL!");
  });

  test("task pane writes formula to selected cell via Excel.run", async () => {
    CustomFunctions.associate("DOUBLE", (n: number) => n * 2);
    mock.setCell("Sheet1", "A1", 5);
    mock.setSelectedRange("Sheet1", "B1");

    await Excel.run(async (context) => {
      const source = context.workbook.worksheets
        .getActiveWorksheet().getRange("A1");
      source.load("values");
      await context.sync();

      const selected = context.workbook.getSelectedRange();
      selected.formulas = [['=DOUBLE(5)']];
      await context.sync();
    });

    expect(mock.getCell("Sheet1", "B1").value).toBe(10);
    expect(mock.getCell("Sheet1", "B1").formula).toBe("=DOUBLE(5)");
  });

  test("unregistered function returns #NAME? error", async () => {
    await mock.setCell("Sheet1", "A1", { formula: "=UNKNOWN(1)" });

    expect(mock.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

  test("function name lookup is case-insensitive", async () => {
    CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

    await mock.setCell("Sheet1", "A1", { formula: "=add(1, 2)" });

    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
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

  test("range.clear() removes cell content and spilled cells", async () => {
    CustomFunctions.associate("MATRIX", () => [
      [1, 2],
      [3, 4],
    ]);

    await mock.setCell("Sheet1", "B2", { formula: "=MATRIX()" });

    await Excel.run(async (context) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet().getRange("B2");
      range.clear();
      await context.sync();
    });

    expect(mock.getCell("Sheet1", "B2").value).toBe("");
    expect(mock.getCell("Sheet1", "C2").value).toBe("");  // spill cleared
  });
});
```

## Reference Files

- `@types/office-js`: `.references/@types/office-js/index.d.ts`
  - `Excel.Range`: line 38064
  - `Excel.Worksheet`: line 36824
  - `Excel.WorksheetCollection`: line 37445
  - `Excel.Workbook`: line 36295 (`getSelectedRange()` at line 36630)
  - `Excel.RequestContext`: line 33103
  - `Range.clear()`: line 38415
  - `NameErrorCellValue` (#NAME?): line 31735
  - `SpillErrorCellValueSubType` (#SPILL!): line 31959
- `@types/custom-functions-runtime`: `.references/@types/custom-functions-runtime/index.d.ts`
  - `CustomFunctions.associate()`: line 13
  - `CustomFunctions.Invocation`: line 47 (`address`, `functionName`, `parameterAddresses`)
  - `CustomFunctions.Error`: line 30
  - `CustomFunctions.ErrorCode`: line 132
- Existing mock library for reference: `.references/office-addin-mock/`
