# Require Metadata for Custom Function Invocation — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Custom functions without loaded metadata return `#NAME?` instead of executing, matching real Office.js behavior.

**Architecture:** Two changes — (1) `FormulaEvaluator` checks for metadata before invoking a function, returning `#NAME?` if missing; (2) `MockCustomFunctions.associate()` emits `console.warn` when no metadata exists for the function ID. All existing tests that call `associate()` without `loadMetadata()` must be updated to load metadata first.

**Tech Stack:** TypeScript, Vitest, Playwright

---

### Task 1: Add metadata-required check to FormulaEvaluator

**Files:**
- Modify: `src/formula-evaluator.ts:17-21`
- Test: `tests/unit/formula-evaluator.test.ts`

- [ ] **Step 1: Write the failing test — function with no metadata returns #NAME?**

Add this test to `tests/unit/formula-evaluator.test.ts` inside the `describe("FormulaEvaluator")` block. This replaces the existing `"no padding when no metadata is loaded"` test (line 124–139) which asserts the OLD behavior (function executes without metadata). Delete that old test and add:

```typescript
test("returns #NAME? when function has no metadata", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  cf.associate("FUNC3", (...args: unknown[]) => 0);
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC3(1, 2)");
  expect(storage.getCell("Sheet1", "A1").value).toBe("#NAME?");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/formula-evaluator.test.ts -t "returns #NAME? when function has no metadata"`

Expected: FAIL — currently the function executes and returns `0`.

- [ ] **Step 3: Implement metadata check in FormulaEvaluator**

Edit `src/formula-evaluator.ts`. Move the `getParameterCount` call before the invocation block, and return `#NAME?` if metadata is missing. The full `evaluateAndStore` method should become:

```typescript
async evaluateAndStore(sheet: string, address: string, formulaStr: string): Promise<void> {
  const parsed = parseFormula(formulaStr);
  if (!parsed) {
    this._storage.setValue(sheet, address, formulaStr);
    return;
  }
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
  const invocation = {
    address: `${sheet}!${address}`,
    functionName: parsed.functionName.toUpperCase(),
  };
  try {
    const paddedArgs = [...parsed.args];
    while (paddedArgs.length < paramCount) {
      paddedArgs.push(null);
    }
    paddedArgs.push(invocation);
    const result = await fn(...paddedArgs);
    if (Array.isArray(result) && Array.isArray(result[0])) {
      this._storage.setFormulaWithSpill(sheet, address, formulaStr, result);
    } else {
      this._storage.setFormula(sheet, address, formulaStr, result);
    }
  } catch {
    this._storage.setFormula(sheet, address, formulaStr, "#VALUE!");
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/unit/formula-evaluator.test.ts -t "returns #NAME? when function has no metadata"`

Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/formula-evaluator.ts tests/unit/formula-evaluator.test.ts
git commit -m "feat: return #NAME? when custom function has no metadata loaded

Functions without JSON metadata now return #NAME? instead of executing,
matching real Office.js behavior where functions.json must define a function
before it can be invoked.

BREAKING CHANGE: Custom functions require loadMetadata() before invocation."
```

---

### Task 2: Add console.warn to associate() when metadata is missing

**Files:**
- Modify: `src/custom-functions-mock.ts:5-16`
- Test: `tests/unit/custom-functions-mock.test.ts`

- [ ] **Step 1: Write the failing test — warn on associate without metadata (single function)**

Add this test to `tests/unit/custom-functions-mock.test.ts`:

```typescript
test("associate warns when no metadata is loaded for the function", () => {
  const cf = new MockCustomFunctions();
  const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  cf.associate("ADD", () => 0);
  expect(warnSpy).toHaveBeenCalledOnce();
  expect(warnSpy.mock.calls[0][0]).toContain("ADD");
  expect(warnSpy.mock.calls[0][0]).toContain("no metadata");
  warnSpy.mockRestore();
});
```

Also add `vi` to the import at line 1:

```typescript
import { describe, expect, test, vi } from "vitest";
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/unit/custom-functions-mock.test.ts -t "associate warns when no metadata"`

Expected: FAIL — no `console.warn` is currently emitted.

- [ ] **Step 3: Write the failing test — warn on associate without metadata (object form)**

Add this test to `tests/unit/custom-functions-mock.test.ts`:

```typescript
test("associate warns for each function without metadata (object form)", () => {
  const cf = new MockCustomFunctions();
  const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  cf.associate({ ADD: () => 0, MUL: () => 0 });
  expect(warnSpy).toHaveBeenCalledTimes(2);
  warnSpy.mockRestore();
});
```

- [ ] **Step 4: Write the failing test — no warn when metadata is loaded**

Add this test to `tests/unit/custom-functions-mock.test.ts`:

```typescript
test("associate does not warn when metadata is loaded", () => {
  const cf = new MockCustomFunctions();
  cf.loadMetadata({ functions: [{ id: "ADD", parameters: [{ name: "a" }] }] });
  const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
  cf.associate("ADD", () => 0);
  expect(warnSpy).not.toHaveBeenCalled();
  warnSpy.mockRestore();
});
```

- [ ] **Step 5: Implement console.warn in associate()**

Edit `src/custom-functions-mock.ts`. Replace the `associate` method (lines 5–16) with:

```typescript
associate(
  idOrMappings: string | Record<string, Function>,
  fn?: Function,
): void {
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

- [ ] **Step 6: Run all custom-functions-mock tests**

Run: `npx vitest run tests/unit/custom-functions-mock.test.ts`

Expected: Some existing tests will fail because they call `associate()` without metadata and don't suppress the warning. We'll fix those in Task 3.

- [ ] **Step 7: Commit**

```bash
git add src/custom-functions-mock.ts tests/unit/custom-functions-mock.test.ts
git commit -m "feat: warn on CustomFunctions.associate() when metadata is not loaded"
```

---

### Task 3: Update existing tests — formula-evaluator.test.ts

Tests in this file that call `associate()` without `loadMetadata()` and expect the function to execute need metadata added. Tests that DON'T involve formula evaluation (only `associate`/`getFunction`) don't need changes.

**Files:**
- Modify: `tests/unit/formula-evaluator.test.ts`

- [ ] **Step 1: Add helper metadata loader to createEvaluator**

Update the `createEvaluator` helper to also return a convenience function. Add this right after the `createEvaluator` function definition:

```typescript
function withMetadata(cf: MockCustomFunctions, id: string, paramCount: number) {
  const parameters = Array.from({ length: paramCount }, (_, i) => ({ name: `p${i}` }));
  cf.loadMetadata({ functions: [{ id, parameters }] });
}
```

- [ ] **Step 2: Add metadata to each test that evaluates formulas without metadata**

The following tests need `withMetadata` calls added. For each test, add the call right after `createEvaluator()` and before `cf.associate(...)`:

**"evaluates registered function and stores result"** (line 14):
```typescript
test("evaluates registered function and stores result", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "ADD", 2);
  cf.associate("ADD", (a: number, b: number) => a + b);
  await evaluator.evaluateAndStore("Sheet1", "A1", "=ADD(1, 2)");
  expect(storage.getCell("Sheet1", "A1").value).toBe(3);
  expect(storage.getCell("Sheet1", "A1").formula).toBe("=ADD(1, 2)");
});
```

**"handles async functions"** (line 28):
```typescript
test("handles async functions", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "ASYNC_ADD", 2);
  cf.associate("ASYNC_ADD", async (a: number, b: number) => a + b);
  await evaluator.evaluateAndStore("Sheet1", "A1", "=ASYNC_ADD(10, 20)");
  expect(storage.getCell("Sheet1", "A1").value).toBe(30);
});
```

**"passes Invocation with address and functionName"** (line 35):
```typescript
test("passes Invocation with address and functionName", async () => {
  const { evaluator, cf } = createEvaluator();
  withMetadata(cf, "CAPTURE", 0);
  let receivedInvocation: unknown;
  cf.associate("CAPTURE", (...args: unknown[]) => {
    receivedInvocation = args[args.length - 1];
    return 0;
  });
  await evaluator.evaluateAndStore("Sheet1", "B3", "=CAPTURE()");
  expect(receivedInvocation).toEqual({ address: "Sheet1!B3", functionName: "CAPTURE" });
});
```

**"spills 2D array result to adjacent cells"** (line 46):
```typescript
test("spills 2D array result to adjacent cells", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "MATRIX", 0);
  cf.associate("MATRIX", () => [[1, 2], [3, 4]]);
  await evaluator.evaluateAndStore("Sheet1", "B2", "=MATRIX()");
  expect(storage.getCell("Sheet1", "B2").value).toBe(1);
  expect(storage.getCell("Sheet1", "C2").value).toBe(2);
  expect(storage.getCell("Sheet1", "B3").value).toBe(3);
  expect(storage.getCell("Sheet1", "C3").value).toBe(4);
});
```

**"stores #VALUE! when function throws"** (line 56):
```typescript
test("stores #VALUE! when function throws", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "FAIL", 0);
  cf.associate("FAIL", () => { throw new Error("boom"); });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FAIL()");
  expect(storage.getCell("Sheet1", "A1").value).toBe("#VALUE!");
});
```

**"lowercase function name resolves correctly"** (line 210):
```typescript
test("lowercase function name resolves correctly", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "ADD", 2);
  cf.associate("ADD", (a: number, b: number) => a + b);
  await evaluator.evaluateAndStore("Sheet1", "A1", "=add(1, 2)");
  expect(storage.getCell("Sheet1", "A1").value).toBe(3);
});
```

**"namespaced function evaluates correctly"** (line 217):
```typescript
test("namespaced function evaluates correctly", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "CONTOSO.ADD", 2);
  cf.associate("CONTOSO.ADD", (a: number, b: number) => a + b);
  await evaluator.evaluateAndStore("Sheet1", "A1", "=CONTOSO.ADD(1, 2)");
  expect(storage.getCell("Sheet1", "A1").value).toBe(3);
  expect(storage.getCell("Sheet1", "A1").formula).toBe("=CONTOSO.ADD(1, 2)");
});
```

**"string argument is passed to function as string"** (line 225):
```typescript
test("string argument is passed to function as string", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "ECHO", 1);
  let receivedArg: unknown;
  cf.associate("ECHO", (val: unknown) => {
    receivedArg = val;
    return val;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", '=ECHO("hello")');
  expect(receivedArg).toBe("hello");
  expect(typeof receivedArg).toBe("string");
});
```

**"boolean argument is passed to function as boolean"** (line 237):
```typescript
test("boolean argument is passed to function as boolean", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "ECHO", 1);
  let receivedArg: unknown;
  cf.associate("ECHO", (val: unknown) => {
    receivedArg = val;
    return val;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=ECHO(TRUE)");
  expect(receivedArg).toBe(true);
  expect(typeof receivedArg).toBe("boolean");
});
```

**"function returning string stores string value"** (line 249):
```typescript
test("function returning string stores string value", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "GREET", 1);
  cf.associate("GREET", (name: string) => "hello " + name);
  await evaluator.evaluateAndStore("Sheet1", "A1", '=GREET("world")');
  expect(storage.getCell("Sheet1", "A1").value).toBe("hello world");
});
```

**"async function that throws stores #VALUE!"** (line 256):
```typescript
test("async function that throws stores #VALUE!", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  withMetadata(cf, "ASYNC_FAIL", 0);
  cf.associate("ASYNC_FAIL", async () => { throw new Error("async boom"); });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=ASYNC_FAIL()");
  expect(storage.getCell("Sheet1", "A1").value).toBe("#VALUE!");
});
```

- [ ] **Step 3: Run all formula-evaluator tests**

Run: `npx vitest run tests/unit/formula-evaluator.test.ts`

Expected: ALL PASS

- [ ] **Step 4: Commit**

```bash
git add tests/unit/formula-evaluator.test.ts
git commit -m "test: add metadata to formula-evaluator tests for metadata-required behavior"
```

---

### Task 4: Update existing tests — setup.test.ts

**Files:**
- Modify: `tests/unit/setup.test.ts`

- [ ] **Step 1: Add metadata to tests that evaluate formulas**

The following tests in `setup.test.ts` call `associate()` and then evaluate formulas. They need metadata loaded first.

**"setCell with formula evaluates custom function"** (line 26):
```typescript
test("setCell with formula evaluates custom function", async () => {
  const { customFunctions, mockOfficeJs } = createMockEnvironment();
  customFunctions.loadMetadata({ functions: [{ id: "ADD", parameters: [{ name: "a" }, { name: "b" }] }] });
  customFunctions.associate("ADD", (a: number, b: number) => a + b);
  await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
  expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(3);
});
```

**"setSelectedRange is used by Excel.run"** (line 59):
```typescript
test("setSelectedRange is used by Excel.run", async () => {
  const { excel, customFunctions, mockOfficeJs } = createMockEnvironment();
  customFunctions.loadMetadata({ functions: [{ id: "DOUBLE", parameters: [{ name: "n" }] }] });
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
```

**"spill collision returns #SPILL!"** (line 93):
```typescript
test("spill collision returns #SPILL!", async () => {
  const { customFunctions, mockOfficeJs } = createMockEnvironment();
  customFunctions.loadMetadata({ functions: [{ id: "MATRIX", parameters: [] }] });
  customFunctions.associate("MATRIX", () => [[1, 2], [3, 4]]);
  mockOfficeJs.excel.setCell("Sheet1", "C2", 999);
  await mockOfficeJs.excel.setCell("Sheet1", "B2", { formula: "=MATRIX()" });
  expect(mockOfficeJs.excel.getCell("Sheet1", "B2").value).toBe("#SPILL!");
});
```

**"function name lookup is case-insensitive"** (line 101):
```typescript
test("function name lookup is case-insensitive", async () => {
  const { customFunctions, mockOfficeJs } = createMockEnvironment();
  customFunctions.loadMetadata({ functions: [{ id: "ADD", parameters: [{ name: "a" }, { name: "b" }] }] });
  customFunctions.associate("ADD", (a: number, b: number) => a + b);
  await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=add(1, 2)" });
  expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(3);
});
```

**"Excel.run formula write evaluates custom function"** (line 121):
```typescript
test("Excel.run formula write evaluates custom function", async () => {
  const { excel, customFunctions, mockOfficeJs } = createMockEnvironment();
  customFunctions.loadMetadata({ functions: [{ id: "DOUBLE", parameters: [{ name: "n" }] }] });
  customFunctions.associate("DOUBLE", (n: number) => n * 2);
  await excel.run(async (context: any) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=DOUBLE(21)"]];
    await context.sync();
  });
  expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(42);
});
```

- [ ] **Step 2: Run all setup tests**

Run: `npx vitest run tests/unit/setup.test.ts`

Expected: ALL PASS

- [ ] **Step 3: Commit**

```bash
git add tests/unit/setup.test.ts
git commit -m "test: add metadata to setup tests for metadata-required behavior"
```

---

### Task 5: Update existing tests — integration.test.ts

**Files:**
- Modify: `tests/unit/integration.test.ts`

- [ ] **Step 1: Add metadata to tests that evaluate formulas**

This file uses the global `CustomFunctions` and `MockOfficeJs` from `import "../../src/index.js"`. To load metadata, use `MockOfficeJs.excel` — but we need `loadMetadata()` which is on the `CustomFunctions` object. Since the global `CustomFunctions` is a `MockCustomFunctions` instance, we can call `(CustomFunctions as any).loadMetadata(...)`. Alternatively, we can call the internal `loadMetadata` through the global. Let's check what's exposed.

Looking at `src/index.ts`, `CustomFunctions` is set to `env.customFunctions` which is a `MockCustomFunctions` instance. So `CustomFunctions.loadMetadata` is available at runtime but not typed in the global declaration. We'll cast to `any`.

Update the following tests. Add a helper at the top of the describe block:

```typescript
function loadMetadata(id: string, paramCount: number) {
  const parameters = Array.from({ length: paramCount }, (_, i) => ({ name: `p${i}` }));
  (CustomFunctions as any).loadMetadata({ functions: [{ id, parameters }] });
}
```

**"full flow: register function, write formula via Excel.run, verify value"** (line 7):
Add `loadMetadata("TRIPLE", 1);` before `CustomFunctions.associate(...)`:
```typescript
test("full flow: register function, write formula via Excel.run, verify value", async () => {
  loadMetadata("TRIPLE", 1);
  CustomFunctions.associate("TRIPLE", (n: number) => n * 3);
  // ... rest unchanged
});
```

**"spill flow: function returns 2D array, verify all cells"** (line 34):
Add `loadMetadata("TABLE", 0);` before `CustomFunctions.associate(...)`:
```typescript
test("spill flow: function returns 2D array, verify all cells", async () => {
  loadMetadata("TABLE", 0);
  CustomFunctions.associate("TABLE", () => [
    // ... rest unchanged
  ]);
  // ...
});
```

**"namespaced custom function evaluates via setCell"** (line 85):
Add `loadMetadata("CONTOSO.ADD", 2);` before `CustomFunctions.associate(...)`:
```typescript
test("namespaced custom function evaluates via setCell", async () => {
  loadMetadata("CONTOSO.ADD", 2);
  CustomFunctions.associate("CONTOSO.ADD", (a: number, b: number) => a + b);
  // ... rest unchanged
});
```

**"mixed argument types are passed correctly to custom function"** (line 98):
Add `loadMetadata("MIXED", 3);` before `CustomFunctions.associate(...)`:
```typescript
test("mixed argument types are passed correctly to custom function", async () => {
  loadMetadata("MIXED", 3);
  let receivedArgs: unknown[] = [];
  CustomFunctions.associate("MIXED", (...args: unknown[]) => {
    // ... rest unchanged
  });
  // ...
});
```

- [ ] **Step 2: Run all integration tests**

Run: `npx vitest run tests/unit/integration.test.ts`

Expected: ALL PASS

- [ ] **Step 3: Commit**

```bash
git add tests/unit/integration.test.ts
git commit -m "test: add metadata to integration tests for metadata-required behavior"
```

---

### Task 6: Update E2E tests

**Files:**
- Modify: `tests/e2e/mock-office-js.e2e.test.ts`

- [ ] **Step 1: Add metadata loading to E2E tests that use custom functions**

E2E tests run in the browser via Playwright's `page.evaluate()`. The `CustomFunctions` global is a `MockCustomFunctions` instance, so `loadMetadata` is available. Add metadata loading inside each `page.evaluate` block before `associate()`.

**"CustomFunctions.associate registers functions and formulas evaluate"** (line 45):
```typescript
test("CustomFunctions.associate registers functions and formulas evaluate", async ({
  page,
}) => {
  const value = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.loadMetadata({ functions: [{ id: "ADD", parameters: [{ name: "a" }, { name: "b" }] }] });
    CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(2, 3)" });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(value).toBe(5);
});
```

**"quoted numeric string argument is preserved as string"** (line 103):
```typescript
test("quoted numeric string argument is preserved as string", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.loadMetadata({ functions: [{ id: "ECHO", parameters: [{ name: "val" }] }] });
    CustomFunctions.associate("ECHO", (val: any) => typeof val + ":" + val);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: '=ECHO("2023")' });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe("string:2023");
});
```

**"quoted string argument after comma has no leading space"** (line 119):
```typescript
test("quoted string argument after comma has no leading space", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.loadMetadata({ functions: [{ id: "JOIN", parameters: [{ name: "a" }, { name: "b" }] }] });
    CustomFunctions.associate("JOIN", (a: any, b: any) => a + ":" + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: '=JOIN(1, "hello")' });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe("1:hello");
});
```

**"throwing function produces #VALUE! error"** (line 145):
```typescript
test("throwing function produces #VALUE! error", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.loadMetadata({ functions: [{ id: "FAIL", parameters: [] }] });
    CustomFunctions.associate("FAIL", () => { throw new Error("boom"); });

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=FAIL()" });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe("#VALUE!");
});
```

**"function returning 2D array spills to adjacent cells"** (line 159):
```typescript
test("function returning 2D array spills to adjacent cells", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.loadMetadata({ functions: [{ id: "MATRIX", parameters: [] }] });
    CustomFunctions.associate("MATRIX", () => [[1, 2], [3, 4]]);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=MATRIX()" });
    return {
      a1: MockOfficeJs.excel.getCell("Sheet1", "A1").value,
      b1: MockOfficeJs.excel.getCell("Sheet1", "B1").value,
      a2: MockOfficeJs.excel.getCell("Sheet1", "A2").value,
      b2: MockOfficeJs.excel.getCell("Sheet1", "B2").value,
    };
  });

  expect(result).toEqual({ a1: 1, b1: 2, a2: 3, b2: 4 });
});
```

**"async custom function evaluates correctly"** (line 178):
```typescript
test("async custom function evaluates correctly", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.loadMetadata({ functions: [{ id: "ASYNC_ADD", parameters: [{ name: "a" }, { name: "b" }] }] });
    CustomFunctions.associate("ASYNC_ADD", async (a: number, b: number) => a + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ASYNC_ADD(10, 20)" });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe(30);
});
```

- [ ] **Step 2: Add a new E2E test — associate without metadata returns #NAME?**

Add this test after the "unregistered function formula produces #NAME? error" test:

```typescript
test("associated function without metadata produces #NAME? error", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe("#NAME?");
});
```

- [ ] **Step 3: Run unit tests first to verify no regressions**

Run: `npx vitest run`

Expected: ALL PASS

- [ ] **Step 4: Build and run E2E tests**

Run: `npm run build && npx playwright test`

Expected: ALL PASS

- [ ] **Step 5: Commit**

```bash
git add tests/e2e/mock-office-js.e2e.test.ts
git commit -m "test: add metadata to E2E tests for metadata-required behavior"
```

---

### Task 7: Suppress console.warn in existing custom-functions-mock tests

**Files:**
- Modify: `tests/unit/custom-functions-mock.test.ts`

- [ ] **Step 1: Suppress console.warn in tests that intentionally call associate without metadata**

Some existing tests call `associate()` without metadata — not because they're testing formula evaluation, but because they're testing the registry behavior of `associate()`. These tests will now emit `console.warn`. Add `vi.spyOn(console, "warn").mockImplementation(() => {})` in an `afterEach` cleanup pattern.

Add a `beforeEach`/`afterEach` pair at the top of the describe block to suppress warnings:

```typescript
import { describe, expect, test, vi, beforeEach, afterEach } from "vitest";
```

And inside the describe block, add:

```typescript
let warnSpy: ReturnType<typeof vi.spyOn>;
beforeEach(() => {
  warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
});
afterEach(() => {
  warnSpy.mockRestore();
});
```

Then update the three new warn-specific tests (from Task 2) to NOT use their own local `warnSpy` — they can use the shared one:

**"associate warns when no metadata is loaded for the function"**:
```typescript
test("associate warns when no metadata is loaded for the function", () => {
  const cf = new MockCustomFunctions();
  cf.associate("ADD", () => 0);
  expect(warnSpy).toHaveBeenCalledOnce();
  expect(warnSpy.mock.calls[0][0]).toContain("ADD");
  expect(warnSpy.mock.calls[0][0]).toContain("no metadata");
});
```

**"associate warns for each function without metadata (object form)"**:
```typescript
test("associate warns for each function without metadata (object form)", () => {
  const cf = new MockCustomFunctions();
  cf.associate({ ADD: () => 0, MUL: () => 0 });
  expect(warnSpy).toHaveBeenCalledTimes(2);
});
```

**"associate does not warn when metadata is loaded"**:
```typescript
test("associate does not warn when metadata is loaded", () => {
  const cf = new MockCustomFunctions();
  cf.loadMetadata({ functions: [{ id: "ADD", parameters: [{ name: "a" }] }] });
  cf.associate("ADD", () => 0);
  expect(warnSpy).not.toHaveBeenCalled();
});
```

- [ ] **Step 2: Run all custom-functions-mock tests**

Run: `npx vitest run tests/unit/custom-functions-mock.test.ts`

Expected: ALL PASS

- [ ] **Step 3: Commit**

```bash
git add tests/unit/custom-functions-mock.test.ts
git commit -m "test: suppress console.warn in custom-functions-mock tests"
```

---

### Task 8: Final verification

- [ ] **Step 1: Run all unit tests**

Run: `npx vitest run`

Expected: ALL PASS

- [ ] **Step 2: Build and run E2E tests**

Run: `npm run build && npx playwright test`

Expected: ALL PASS

- [ ] **Step 3: Verify the new behavior manually**

Quick sanity check in the test output: ensure the new test `"returns #NAME? when function has no metadata"` and `"associated function without metadata produces #NAME? error"` both pass.
