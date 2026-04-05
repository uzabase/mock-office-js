# mock-office-js

In-memory mock of the Office JavaScript API for testing Office Add-ins.

> [!WARNING]
> This library is in early development. Only a subset of the Excel JavaScript API is currently supported. See [Limitations](#limitations) for details.

## Install

```
npm install mock-office-js
```

## Usage

```typescript
import { ExcelMock } from "mock-office-js";

const mock = new ExcelMock();

globalThis.Excel = mock.excel;
globalThis.CustomFunctions = mock.customFunctions;

// Register a custom function
CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

// Set a formula and check the result
await mock.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
mock.getCell("Sheet1", "A1").value; // 3
```

## API

### `ExcelMock`

| Method | Description |
|---|---|
| `excel.run(callback)` | Mock of `Excel.run()`. Creates a `RequestContext` and passes it to the callback. |
| `customFunctions` | Mock of the `CustomFunctions` namespace. Supports `associate()`, `Error`, and `ErrorCode`. |
| `setCell(sheet, address, value)` | Sets a cell value. Pass `{ formula: "=..." }` to evaluate a custom function (returns `Promise<void>`). |
| `setCells(sheet, startAddress, values)` | Sets a 2D array of values starting from the given address. |
| `getCell(sheet, address)` | Returns the `CellState` for a cell. |
| `getCells(sheet, rangeAddress)` | Returns a 2D array of `CellState` for a range. |
| `setSelectedRange(sheet, address)` | Sets the selected range for `workbook.getSelectedRange()`. |
| `setActiveWorksheet(sheet)` | Sets the active worksheet. |
| `addWorksheet(name)` | Adds a new worksheet. |
| `reset()` | Clears all cell data, worksheets, custom functions, and selections. |

## Limitations

- No native Excel functions (`=SUM()`, `=VLOOKUP()`, etc.) — only registered custom functions are evaluated
- No cell reference resolution in formula arguments (e.g., `=ADD(A1, B1)`)
- No streaming or cancelable custom functions
- No event handling

## License

MIT
