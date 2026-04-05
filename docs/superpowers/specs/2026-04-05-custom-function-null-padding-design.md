# Custom Function Optional Argument Null Padding

## Problem

When a custom function is defined with 3 parameters but called with only 2 arguments (e.g., `=FUNC(1, 2)`), the `FormulaEvaluator` currently passes `fn(1, 2, invocation)`. This causes the `invocation` object to land in the wrong parameter position, leading to `#VALUE!` errors.

## How Real Excel Works

1. The native layer (C++) reads `functions.json` metadata and knows each function's parameter count
2. When a user omits optional parameters, the native layer pads missing arguments with `null`
3. The JS runtime (`office.js`) appends `InvocationContext` as the last element of `parameterValues`
4. The function is called via `.apply()` with the padded array

For `=FUNC(1, 2)` where FUNC has 3 params + invocation:
- Native layer builds: `[1, 2, null]`
- JS layer appends: `[1, 2, null, invocationContext]`
- Function receives: `fn(1, 2, null, invocationContext)`

Sources:
- [Options for Excel custom functions - Microsoft Learn](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-parameter-options)
- `.references/office-debug-js/excel-web-16.00.debug.js` lines 71488-71685 (CustomFunctionProxy)
- Manifest.xml uses `<Metadata>` tag with `resid="Functions.Metadata.Url"` to reference functions.json

## Design

### API

Replace direct constructor usage with a static factory method when metadata is needed:

```ts
// New: with metadata (fetches functions.json)
const mock = await ExcelMock.create({
  functionsMetadataUrl: "/functions.json"
});

// Existing: without metadata (unchanged, no padding)
const mock = new ExcelMock();
```

The property name `functionsMetadataUrl` follows the manifest.xml naming convention where `resid="Functions.Metadata.Url"` references the functions.json location.

### functions.json Format

Standard Office Add-in custom functions metadata:

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        { "name": "first", "type": "number" },
        { "name": "second", "type": "number" },
        { "name": "third", "type": "number", "optional": true }
      ]
    }
  ]
}
```

### Options Type

```ts
interface ExcelMockOptions {
  functionsMetadataUrl: string;
}
```

### Internal Flow

1. `ExcelMock.create()` uses `fetch()` to load the URL and parses the JSON
2. Parameter counts are stored in `MockCustomFunctions` (keyed by function ID, uppercased)
3. `FormulaEvaluator.evaluateAndStore()` queries the parameter count from metadata
4. Missing arguments are padded with `null`
5. `invocation` is appended to the padded array (matching real Excel's `parameterValues.push(invocationContext)`)
6. Function is called: `fn(...paddedArgs)` where `paddedArgs` already contains invocation at the end

### Changes Required

| File | Change |
|---|---|
| `excel-mock.ts` | Add `static async create(options)` factory method |
| `custom-functions-mock.ts` | Add metadata storage (parameter count per function) |
| `formula-evaluator.ts` | Add null-padding logic using metadata |
| `request-context.ts` | No change needed — already receives the shared `MockCustomFunctions` instance which holds the metadata |

### `reset()` Behavior

`reset()` preserves metadata. Metadata is structural configuration (analogous to functions.json existing on disk) and survives `reset()`, which only clears runtime state (cell values, registered functions).

### Out of Scope

- `"repeating": true` parameters (variable-length argument lists) — padding is not applicable

### Compatibility

- `new ExcelMock()` continues to work as before (no metadata, no padding)
- `ExcelMock.create()` is the new recommended path when custom functions with optional parameters are used
- `CustomFunctions.associate()` API is unchanged (same as real Office.js)
- No changes to production/application code required
