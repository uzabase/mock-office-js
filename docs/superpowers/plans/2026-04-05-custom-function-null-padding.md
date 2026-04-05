# Custom Function Null Padding Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Pad missing optional custom function arguments with `null` before appending invocation, matching real Excel's native layer behavior.

**Architecture:** `ExcelMock.create()` fetches `functions.json` via URL, stores parameter counts in `MockCustomFunctions`, and `FormulaEvaluator` uses those counts to pad args with `null` before appending invocation.

**Tech Stack:** TypeScript, Vitest, fetch API

---

## File Structure

| File | Role | Action |
|---|---|---|
| `src/custom-functions-mock.ts` | Stores metadata (param counts per function ID) | Modify |
| `tests/custom-functions-mock.test.ts` | Tests for metadata storage and reset behavior | Modify |
| `src/formula-evaluator.ts` | Pads args with `null` using metadata, appends invocation | Modify |
| `tests/formula-evaluator.test.ts` | Tests for null-padding logic | Modify |
| `src/excel-mock.ts` | `static async create(options)` factory method | Modify |
| `tests/excel-mock.test.ts` | Integration tests for `ExcelMock.create()` | Modify |
| `src/index.ts` | Export `ExcelMockOptions` type | Modify |

---

### Task 1: MockCustomFunctions — metadata storage

Store parameter counts from functions.json metadata. `loadMetadata` is an internal method (not part of the public `CustomFunctions` API).

**Files:**
- Modify: `src/custom-functions-mock.ts`
- Modify: `tests/custom-functions-mock.test.ts`

- [ ] **Step 1: Write failing test — `loadMetadata` stores parameter count**

```ts
test("loadMetadata stores parameter count from functions.json", () => {
  const cf = new MockCustomFunctions();
  cf.loadMetadata({
    functions: [
      {
        id: "ADD",
        name: "ADD",
        parameters: [
          { name: "first", type: "number" },
          { name: "second", type: "number" },
          { name: "third", type: "number", optional: true },
        ],
      },
    ],
  });
  expect(cf.getParameterCount("ADD")).toBe(3);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/custom-functions-mock.test.ts`
Expected: FAIL — `loadMetadata` and `getParameterCount` do not exist

- [ ] **Step 3: Implement `loadMetadata` and `getParameterCount`**

In `src/custom-functions-mock.ts`, add a `parameterCounts` map and two methods:

```ts
private parameterCounts = new Map<string, number>();

loadMetadata(metadata: { functions: Array<{ id: string; parameters?: Array<unknown> }> }): void {
  for (const fn of metadata.functions) {
    this.parameterCounts.set(fn.id.toUpperCase(), fn.parameters?.length ?? 0);
  }
}

getParameterCount(id: string): number | undefined {
  return this.parameterCounts.get(id.toUpperCase());
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/custom-functions-mock.test.ts`
Expected: PASS

- [ ] **Step 5: Write failing test — case-insensitive lookup**

```ts
test("getParameterCount is case-insensitive", () => {
  const cf = new MockCustomFunctions();
  cf.loadMetadata({
    functions: [
      { id: "CONTOSO.ADD", name: "CONTOSO.ADD", parameters: [{ name: "a" }] },
    ],
  });
  expect(cf.getParameterCount("contoso.add")).toBe(1);
});
```

- [ ] **Step 6: Run test to verify it passes** (should pass already since `toUpperCase()` is used)

Run: `npx vitest run tests/custom-functions-mock.test.ts`
Expected: PASS

- [ ] **Step 7: Write failing test — reset preserves metadata**

```ts
test("reset preserves metadata", () => {
  const cf = new MockCustomFunctions();
  cf.loadMetadata({
    functions: [
      { id: "ADD", name: "ADD", parameters: [{ name: "a" }, { name: "b" }] },
    ],
  });
  cf.associate("ADD", () => 0);
  cf.reset();
  expect(cf.getFunction("ADD")).toBeUndefined();
  expect(cf.getParameterCount("ADD")).toBe(2);
});
```

- [ ] **Step 8: Run test to verify it passes** (should pass since `reset()` only clears `registry`, not `parameterCounts`)

Run: `npx vitest run tests/custom-functions-mock.test.ts`
Expected: PASS

- [ ] **Step 9: Write failing test — getParameterCount returns undefined for unknown function**

```ts
test("getParameterCount returns undefined for unknown function", () => {
  const cf = new MockCustomFunctions();
  expect(cf.getParameterCount("UNKNOWN")).toBeUndefined();
});
```

- [ ] **Step 10: Run test to verify it passes**

Run: `npx vitest run tests/custom-functions-mock.test.ts`
Expected: PASS

- [ ] **Step 11: Commit**

```bash
git add src/custom-functions-mock.ts tests/custom-functions-mock.test.ts
git commit -m "feat: add metadata storage to MockCustomFunctions"
```

---

### Task 2: FormulaEvaluator — null padding

Use metadata parameter counts to pad missing args with `null`, then append invocation at the end.

**Files:**
- Modify: `src/formula-evaluator.ts`
- Modify: `tests/formula-evaluator.test.ts`

- [ ] **Step 1: Write failing test — 3-param function called with 2 args gets null padding**

```ts
test("pads missing args with null when metadata is loaded", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  cf.loadMetadata({
    functions: [
      {
        id: "FUNC3",
        name: "FUNC3",
        parameters: [{ name: "a" }, { name: "b" }, { name: "c" }],
      },
    ],
  });
  let receivedArgs: unknown[] = [];
  cf.associate("FUNC3", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC3(1, 2)");
  // args should be [1, 2, null, invocation]
  expect(receivedArgs).toHaveLength(4);
  expect(receivedArgs[0]).toBe(1);
  expect(receivedArgs[1]).toBe(2);
  expect(receivedArgs[2]).toBeNull();
  expect(receivedArgs[3]).toEqual({
    address: "Sheet1!A1",
    functionName: "FUNC3",
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: FAIL — `receivedArgs[2]` is the invocation object, not `null`

- [ ] **Step 3: Implement null padding in `evaluateAndStore`**

Replace line 27 in `src/formula-evaluator.ts`. Change from:

```ts
const result = await fn(...parsed.args, invocation);
```

To:

```ts
const paramCount = this._customFunctions.getParameterCount(parsed.functionName);
const paddedArgs = [...parsed.args];
if (paramCount !== undefined) {
  while (paddedArgs.length < paramCount) {
    paddedArgs.push(null);
  }
}
paddedArgs.push(invocation);
const result = await fn(...paddedArgs);
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: ALL PASS (new test + all existing tests)

- [ ] **Step 5: Write failing test — no padding when all args provided**

```ts
test("no padding when all args are provided", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  cf.loadMetadata({
    functions: [
      {
        id: "FUNC3",
        name: "FUNC3",
        parameters: [{ name: "a" }, { name: "b" }, { name: "c" }],
      },
    ],
  });
  let receivedArgs: unknown[] = [];
  cf.associate("FUNC3", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC3(1, 2, 3)");
  expect(receivedArgs).toHaveLength(4);
  expect(receivedArgs[0]).toBe(1);
  expect(receivedArgs[1]).toBe(2);
  expect(receivedArgs[2]).toBe(3);
  expect(receivedArgs[3]).toEqual({
    address: "Sheet1!A1",
    functionName: "FUNC3",
  });
});
```

- [ ] **Step 6: Run test to verify it passes**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: PASS

- [ ] **Step 7: Write failing test — no metadata means no padding (backward compat)**

```ts
test("no padding when no metadata is loaded", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  let receivedArgs: unknown[] = [];
  cf.associate("FUNC3", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC3(1, 2)");
  // Without metadata: args = [1, 2, invocation] (no padding)
  expect(receivedArgs).toHaveLength(3);
  expect(receivedArgs[0]).toBe(1);
  expect(receivedArgs[1]).toBe(2);
  expect(receivedArgs[2]).toEqual({
    address: "Sheet1!A1",
    functionName: "FUNC3",
  });
});
```

- [ ] **Step 8: Run test to verify it passes**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: PASS

- [ ] **Step 9: Write failing test — zero args with metadata pads all to null**

```ts
test("pads all args to null when called with zero args", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  cf.loadMetadata({
    functions: [
      {
        id: "FUNC2",
        name: "FUNC2",
        parameters: [{ name: "a" }, { name: "b" }],
      },
    ],
  });
  let receivedArgs: unknown[] = [];
  cf.associate("FUNC2", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC2()");
  expect(receivedArgs).toHaveLength(3);
  expect(receivedArgs[0]).toBeNull();
  expect(receivedArgs[1]).toBeNull();
  expect(receivedArgs[2]).toEqual({
    address: "Sheet1!A1",
    functionName: "FUNC2",
  });
});
```

- [ ] **Step 10: Run test to verify it passes**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: PASS

- [ ] **Step 11: Write failing test — extra args beyond metadata param count are not truncated**

```ts
test("does not truncate when more args provided than metadata param count", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  cf.loadMetadata({
    functions: [
      {
        id: "FUNC2",
        name: "FUNC2",
        parameters: [{ name: "a" }, { name: "b" }],
      },
    ],
  });
  let receivedArgs: unknown[] = [];
  cf.associate("FUNC2", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC2(1, 2, 3)");
  expect(receivedArgs).toHaveLength(4);
  expect(receivedArgs[0]).toBe(1);
  expect(receivedArgs[1]).toBe(2);
  expect(receivedArgs[2]).toBe(3);
});
```

- [ ] **Step 12: Run test to verify it passes**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: PASS

- [ ] **Step 13: Write failing test — function with zero parameters in metadata**

```ts
test("function with zero parameters in metadata only receives invocation", async () => {
  const { evaluator, storage, cf } = createEvaluator();
  cf.loadMetadata({
    functions: [
      { id: "NOPARAM", name: "NOPARAM", parameters: [] },
    ],
  });
  let receivedArgs: unknown[] = [];
  cf.associate("NOPARAM", (...args: unknown[]) => {
    receivedArgs = args;
    return 42;
  });
  await evaluator.evaluateAndStore("Sheet1", "A1", "=NOPARAM()");
  expect(receivedArgs).toHaveLength(1);
  expect(receivedArgs[0]).toEqual({
    address: "Sheet1!A1",
    functionName: "NOPARAM",
  });
});
```

- [ ] **Step 14: Run test to verify it passes**

Run: `npx vitest run tests/formula-evaluator.test.ts`
Expected: PASS

- [ ] **Step 15: Commit**

```bash
git add src/formula-evaluator.ts tests/formula-evaluator.test.ts
git commit -m "feat: pad missing custom function args with null using metadata"
```

---

### Task 3: ExcelMock.create() — static factory with fetch

Add `static async create(options)` that fetches functions.json and loads metadata.

**Files:**
- Modify: `src/excel-mock.ts`
- Modify: `tests/excel-mock.test.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Write failing test — `ExcelMock.create` fetches metadata and pads args**

```ts
test("ExcelMock.create fetches metadata and pads missing args with null", async () => {
  const metadata = {
    functions: [
      {
        id: "ADD3",
        name: "ADD3",
        parameters: [
          { name: "a", type: "number" },
          { name: "b", type: "number" },
          { name: "c", type: "number", optional: true },
        ],
      },
    ],
  };
  vi.stubGlobal(
    "fetch",
    vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(metadata),
    }),
  );

  const mock = await ExcelMock.create({
    functionsMetadataUrl: "/functions.json",
  });

  let receivedArgs: unknown[] = [];
  mock.customFunctions.associate("ADD3", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await mock.setCell("Sheet1", "A1", { formula: "=ADD3(1, 2)" });
  expect(receivedArgs[2]).toBeNull();
  expect(receivedArgs[3]).toEqual({
    address: "Sheet1!A1",
    functionName: "ADD3",
  });

  vi.unstubAllGlobals();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run tests/excel-mock.test.ts`
Expected: FAIL — `ExcelMock.create` does not exist

- [ ] **Step 3: Implement `ExcelMock.create` and `ExcelMockOptions`**

In `src/excel-mock.ts`:

```ts
export interface ExcelMockOptions {
  functionsMetadataUrl: string;
}

// Inside ExcelMock class:
static async create(options: ExcelMockOptions): Promise<ExcelMock> {
  const mock = new ExcelMock();
  const response = await fetch(options.functionsMetadataUrl);
  if (!response.ok) {
    throw new Error(`Failed to fetch functions metadata: ${response.status}`);
  }
  const metadata = await response.json();
  mock.customFunctions.loadMetadata(metadata);
  return mock;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run tests/excel-mock.test.ts`
Expected: ALL PASS

- [ ] **Step 5: Write failing test — `create` calls fetch with the provided URL**

```ts
test("ExcelMock.create calls fetch with the provided URL", async () => {
  const fetchMock = vi.fn().mockResolvedValue({
    ok: true,
    json: () => Promise.resolve({ functions: [] }),
  });
  vi.stubGlobal("fetch", fetchMock);

  await ExcelMock.create({ functionsMetadataUrl: "/my/functions.json" });
  expect(fetchMock).toHaveBeenCalledWith("/my/functions.json");

  vi.unstubAllGlobals();
});
```

- [ ] **Step 6: Run test to verify it passes**

Run: `npx vitest run tests/excel-mock.test.ts`
Expected: PASS

- [ ] **Step 7: Write failing test — `create` throws on fetch failure**

```ts
test("ExcelMock.create throws on fetch failure", async () => {
  vi.stubGlobal(
    "fetch",
    vi.fn().mockResolvedValue({ ok: false, status: 404 }),
  );

  await expect(
    ExcelMock.create({ functionsMetadataUrl: "/missing.json" }),
  ).rejects.toThrow("Failed to fetch functions metadata: 404");

  vi.unstubAllGlobals();
});
```

- [ ] **Step 8: Run test to verify it passes**

Run: `npx vitest run tests/excel-mock.test.ts`
Expected: PASS

- [ ] **Step 9: Write failing test — reset preserves metadata from create**

```ts
test("reset preserves metadata loaded via create", async () => {
  const metadata = {
    functions: [
      {
        id: "ADD3",
        name: "ADD3",
        parameters: [{ name: "a" }, { name: "b" }, { name: "c" }],
      },
    ],
  };
  vi.stubGlobal(
    "fetch",
    vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(metadata),
    }),
  );

  const mock = await ExcelMock.create({
    functionsMetadataUrl: "/functions.json",
  });
  mock.customFunctions.associate("ADD3", () => 0);
  mock.reset();

  expect(mock.customFunctions.getFunction("ADD3")).toBeUndefined();

  let receivedArgs: unknown[] = [];
  mock.customFunctions.associate("ADD3", (...args: unknown[]) => {
    receivedArgs = args;
    return 0;
  });
  await mock.setCell("Sheet1", "A1", { formula: "=ADD3(1, 2)" });
  expect(receivedArgs[2]).toBeNull();

  vi.unstubAllGlobals();
});
```

- [ ] **Step 10: Run test to verify it passes**

Run: `npx vitest run tests/excel-mock.test.ts`
Expected: PASS

- [ ] **Step 11: Export `ExcelMockOptions` from `src/index.ts`**

Change `src/index.ts` to:

```ts
export { ExcelMock } from "./excel-mock.js";
export type { ExcelMockOptions } from "./excel-mock.js";
export type { CellState } from "./cell-storage.js";
```

- [ ] **Step 12: Commit**

```bash
git add src/excel-mock.ts src/index.ts tests/excel-mock.test.ts
git commit -m "feat: add ExcelMock.create() for loading functions metadata via URL"
```

---

### Task 4: Run full test suite

- [ ] **Step 1: Run all tests**

Run: `npx vitest run`
Expected: ALL PASS

- [ ] **Step 2: Run type checks**

Run: `npx vitest --typecheck --run`
Expected: ALL PASS

- [ ] **Step 3: Build**

Run: `npm run build`
Expected: No errors
