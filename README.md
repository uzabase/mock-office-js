# mock-office-js

In-memory mock of the Office JavaScript API for testing Office Add-ins.

> [!WARNING]
> This library is in early development. Only a subset of the Excel JavaScript API is currently supported. See [Limitations](#limitations) for details.

## Install

```
npm install mock-office-js
```

## Usage

Importing the package registers `Excel`, `Office`, `CustomFunctions`, and `MockOfficeJs` on `globalThis` as a side effect — just like the real Office.js runtime does in an Add-in.

```typescript
import "mock-office-js";

// Register a custom function
CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

// Set a formula and check the result
await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
MockOfficeJs.excel.getCell("Sheet1", "A1").value; // 3
```

### With Vitest / Jest

Add `mock-office-js` to your test setup so the globals are available in every test:

```typescript
// vitest.config.ts
export default defineConfig({
  test: {
    setupFiles: ["mock-office-js"],
  },
});
```

Then use `MockOfficeJs.reset()` to clean up between tests:

```typescript
afterEach(() => {
  MockOfficeJs.reset();
});
```

## API

### Globals

Importing `mock-office-js` registers the following globals:

| Global | Description |
|---|---|
| `Excel` | Mock of the `Excel` namespace. Supports `Excel.run(callback)`. |
| `Office` | Mock of the `Office` namespace. Supports `Office.onReady()`. |
| `CustomFunctions` | Mock of the `CustomFunctions` namespace. Supports `associate()`, `Error`, and `ErrorCode`. |
| `MockOfficeJs` | Test helper for setting up and inspecting mock state. |

### `MockOfficeJs`

| Method | Description |
|---|---|
| `excel.setCell(sheet, address, value)` | Sets a cell value. Pass `{ formula: "=..." }` to evaluate a custom function (returns `Promise<void>`). |
| `excel.setCells(sheet, startAddress, values)` | Sets a 2D array of values starting from the given address. |
| `excel.getCell(sheet, address)` | Returns the `CellState` for a cell. |
| `excel.getCells(sheet, rangeAddress)` | Returns a 2D array of `CellState` for a range. |
| `excel.setSelectedRange(sheet, address)` | Sets the selected range for `workbook.getSelectedRange()`. |
| `excel.setActiveWorksheet(sheet)` | Sets the active worksheet. |
| `excel.addWorksheet(name)` | Adds a new worksheet. |
| `excel.loadFunctionsMetadata(url)` | Fetches and loads custom functions metadata JSON from a URL. |
| `reset()` | Clears all cell data, worksheets, custom functions, and selections. |

## Limitations

- No native Excel functions (`=SUM()`, `=VLOOKUP()`, etc.) — only registered custom functions are evaluated
- No cell reference resolution in formula arguments (e.g., `=ADD(A1, B1)`)
- No streaming or cancelable custom functions
- No event handling

## License

MIT
