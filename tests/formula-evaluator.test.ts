import { describe, expect, test } from "vitest";
import { FormulaEvaluator } from "../src/formula-evaluator.js";
import { CellStorage } from "../src/cell-storage.js";
import { MockCustomFunctions } from "../src/custom-functions-mock.js";

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
});
