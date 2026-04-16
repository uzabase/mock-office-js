# Design: Require metadata for custom function invocation

## Problem

Currently, custom functions work even without loading `functions.json` metadata via `loadFunctionsMetadata()` or `loadMetadata()`. This diverges from real Office.js behavior, where:

1. Excel loads JSON metadata from the manifest automatically on add-in startup
2. `CustomFunctions.associate()` links a function ID to its JS implementation
3. Functions **not defined in the JSON metadata** are not registered in Excel — they return `#NAME?`

In mock-office-js, metadata is only used for null-padding optional parameters. Without metadata, functions still execute but with incorrect argument handling. This makes bugs hard to detect.

## Design

### Behavior change

**Formula evaluation**: When a formula like `=ADD(1, 2)` is evaluated, if the function `ADD` has been `associate()`d but has **no metadata loaded**, the cell should display `#NAME?` — matching real Excel behavior.

**`associate()` warning**: When `CustomFunctions.associate()` is called for a function ID that has no metadata loaded, emit a `console.warn` to help developers notice the issue early.

### Implementation

#### `FormulaEvaluator.evaluateAndStore()` (`src/formula-evaluator.ts`)

After finding the function via `getFunction()`, check `getParameterCount()`. If it returns `undefined` (no metadata), store `#NAME?` instead of invoking the function.

```typescript
const fn = this._customFunctions.getFunction(parsed.functionName);
if (!fn) {
  this._storage.setFormula(sheet, address, formulaStr, "#NAME?");
  return;
}
const paramCount = this._customFunctions.getParameterCount(parsed.functionName);
if (paramCount === undefined) {
  this._storage.setFormula(sheet, address, formulaStr, "#NAME?");
  return;
}
```

#### `MockCustomFunctions.associate()` (`src/custom-functions-mock.ts`)

Add a `console.warn` when associating a function that has no metadata:

```typescript
associate(idOrMappings: string | Record<string, Function>, fn?: Function): void {
  if (typeof idOrMappings === "string") {
    this.registry.set(idOrMappings.toUpperCase(), fn!);
    if (!this.parameterCounts.has(idOrMappings.toUpperCase())) {
      console.warn(
        `[mock-office-js] CustomFunctions.associate("${idOrMappings}"): no metadata loaded for this function. ` +
        `Call loadFunctionsMetadata() or loadMetadata() first. Without metadata, the function will return #NAME?.`
      );
    }
  } else {
    for (const [id, func] of Object.entries(idOrMappings)) {
      this.registry.set(id.toUpperCase(), func);
      if (!this.parameterCounts.has(id.toUpperCase())) {
        console.warn(
          `[mock-office-js] CustomFunctions.associate("${id}"): no metadata loaded for this function. ` +
          `Call loadFunctionsMetadata() or loadMetadata() first. Without metadata, the function will return #NAME?.`
        );
      }
    }
  }
}
```

#### Test updates

Existing tests that call `associate()` without `loadMetadata()` will need to add metadata setup. This affects:

- `tests/unit/formula-evaluator.test.ts`
- `tests/unit/integration.test.ts`
- `tests/unit/setup.test.ts`
- `tests/e2e/mock-office-js.e2e.test.ts`

A new test should verify that a function without metadata returns `#NAME?`.

### What stays the same

- `loadFunctionsMetadata(url)` API unchanged
- `loadMetadata()` API unchanged
- `reset()` still clears function registry but preserves metadata (existing behavior)

### Breaking change

This is a breaking change. Users who do not call `loadMetadata()` / `loadFunctionsMetadata()` will see all custom function formulas return `#NAME?`. Given the library is at v0.0.5 with limited adoption, this is acceptable.
