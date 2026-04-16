import { describe, expect, test } from "vitest";
import { FormulaEvaluator } from "../../src/formula-evaluator.js";
import { CellStorage } from "../../src/cell-storage.js";
import { MockCustomFunctions } from "../../src/custom-functions-mock.js";

describe("FormulaEvaluator", () => {
  function createEvaluator() {
    const storage = new CellStorage();
    const cf = new MockCustomFunctions();
    const evaluator = new FormulaEvaluator(storage, cf);
    return { evaluator, storage, cf };
  }

  test("evaluates registered function and stores result", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("ADD", (a: number, b: number) => a + b);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=ADD(1, 2)");
    expect(storage.getCell("Sheet1", "A1").value).toBe(3);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("=ADD(1, 2)");
  });

  test("stores #NAME? for unregistered function", async () => {
    const { evaluator, storage } = createEvaluator();
    await evaluator.evaluateAndStore("Sheet1", "A1", "=UNKNOWN(1)");
    expect(storage.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

  test("handles async functions", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("ASYNC_ADD", async (a: number, b: number) => a + b);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=ASYNC_ADD(10, 20)");
    expect(storage.getCell("Sheet1", "A1").value).toBe(30);
  });

  test("passes Invocation with address and functionName", async () => {
    const { evaluator, cf } = createEvaluator();
    let receivedInvocation: unknown;
    cf.associate("CAPTURE", (...args: unknown[]) => {
      receivedInvocation = args[args.length - 1];
      return 0;
    });
    await evaluator.evaluateAndStore("Sheet1", "B3", "=CAPTURE()");
    expect(receivedInvocation).toEqual({ address: "Sheet1!B3", functionName: "CAPTURE" });
  });

  test("spills 2D array result to adjacent cells", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("MATRIX", () => [[1, 2], [3, 4]]);
    await evaluator.evaluateAndStore("Sheet1", "B2", "=MATRIX()");
    expect(storage.getCell("Sheet1", "B2").value).toBe(1);
    expect(storage.getCell("Sheet1", "C2").value).toBe(2);
    expect(storage.getCell("Sheet1", "B3").value).toBe(3);
    expect(storage.getCell("Sheet1", "C3").value).toBe(4);
  });

  test("stores #VALUE! when function throws", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("FAIL", () => { throw new Error("boom"); });
    await evaluator.evaluateAndStore("Sheet1", "A1", "=FAIL()");
    expect(storage.getCell("Sheet1", "A1").value).toBe("#VALUE!");
  });

  test("non-formula string stores as plain value", async () => {
    const { evaluator, storage } = createEvaluator();
    await evaluator.evaluateAndStore("Sheet1", "A1", "hello");
    expect(storage.getCell("Sheet1", "A1").value).toBe("hello");
    expect(storage.getCell("Sheet1", "A1").formula).toBe("");
  });

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
    expect(receivedArgs).toHaveLength(4);
    expect(receivedArgs[0]).toBe(1);
    expect(receivedArgs[1]).toBe(2);
    expect(receivedArgs[2]).toBeNull();
    expect(receivedArgs[3]).toEqual({
      address: "Sheet1!A1",
      functionName: "FUNC3",
    });
  });

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

  test("returns #NAME? when function has no metadata", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("FUNC3", (...args: unknown[]) => 0);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=FUNC3(1, 2)");
    expect(storage.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

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

  test("lowercase function name resolves correctly", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("ADD", (a: number, b: number) => a + b);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=add(1, 2)");
    expect(storage.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("namespaced function evaluates correctly", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("CONTOSO.ADD", (a: number, b: number) => a + b);
    await evaluator.evaluateAndStore("Sheet1", "A1", "=CONTOSO.ADD(1, 2)");
    expect(storage.getCell("Sheet1", "A1").value).toBe(3);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("=CONTOSO.ADD(1, 2)");
  });

  test("string argument is passed to function as string", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    let receivedArg: unknown;
    cf.associate("ECHO", (val: unknown) => {
      receivedArg = val;
      return val;
    });
    await evaluator.evaluateAndStore("Sheet1", "A1", '=ECHO("hello")');
    expect(receivedArg).toBe("hello");
    expect(typeof receivedArg).toBe("string");
  });

  test("boolean argument is passed to function as boolean", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    let receivedArg: unknown;
    cf.associate("ECHO", (val: unknown) => {
      receivedArg = val;
      return val;
    });
    await evaluator.evaluateAndStore("Sheet1", "A1", "=ECHO(TRUE)");
    expect(receivedArg).toBe(true);
    expect(typeof receivedArg).toBe("boolean");
  });

  test("function returning string stores string value", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("GREET", (name: string) => "hello " + name);
    await evaluator.evaluateAndStore("Sheet1", "A1", '=GREET("world")');
    expect(storage.getCell("Sheet1", "A1").value).toBe("hello world");
  });

  test("async function that throws stores #VALUE!", async () => {
    const { evaluator, storage, cf } = createEvaluator();
    cf.associate("ASYNC_FAIL", async () => { throw new Error("async boom"); });
    await evaluator.evaluateAndStore("Sheet1", "A1", "=ASYNC_FAIL()");
    expect(storage.getCell("Sheet1", "A1").value).toBe("#VALUE!");
  });
});
