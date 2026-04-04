import { describe, expect, test } from "vitest";
import { MockRequestContext } from "../src/request-context";
import { CellStorage } from "../src/cell-storage";
import { MockCustomFunctions } from "../src/custom-functions-mock";
import { MockWorksheetCollection } from "../src/worksheet-collection";
import { MockRange } from "../src/range";

describe("MockRequestContext", () => {
  function createContext() {
    const storage = new CellStorage();
    const cf = new MockCustomFunctions();
    const pendingLoads: MockRange[] = [];
    const worksheets = new MockWorksheetCollection(storage, pendingLoads);
    const context = new MockRequestContext(storage, cf, worksheets);
    return { context, storage, cf };
  }

  test("context.workbook is accessible", () => {
    const { context } = createContext();
    expect(context.workbook).toBeDefined();
    expect(context.workbook.worksheets).toBeDefined();
  });

  test("sync resolves pending loads", async () => {
    const { context, storage } = createContext();
    storage.setValue("Sheet1", "A1", 42);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.load("values");
    await context.sync();
    expect(range.values).toEqual([[42]]);
  });

  test("sync executes queued value writes", async () => {
    const { context, storage } = createContext();
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.values = [[99]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe(99);
  });

  test("sync executes queued formula writes and evaluates custom functions", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("ADD", (a: number, b: number) => a + b);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=ADD(1, 2)"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe(3);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("=ADD(1, 2)");
  });

  test("sync evaluates async custom functions", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("ASYNC_ADD", async (a: number, b: number) => a + b);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=ASYNC_ADD(10, 20)"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe(30);
  });

  test("unregistered function formula sets #NAME? value", async () => {
    const { context, storage } = createContext();
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.formulas = [["=UNKNOWN(1)"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe("#NAME?");
  });

  test("formula evaluation passes Invocation with address and functionName", async () => {
    const { context, cf } = createContext();
    let receivedInvocation: unknown;
    cf.associate("CAPTURE", (...args: unknown[]) => {
      receivedInvocation = args[args.length - 1];
      return 0;
    });
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("B3");
    range.formulas = [["=CAPTURE()"]];
    await context.sync();
    expect(receivedInvocation).toEqual({ address: "Sheet1!B3", functionName: "CAPTURE" });
  });

  test("spilling formula writes to multiple cells", async () => {
    const { context, storage, cf } = createContext();
    cf.associate("MATRIX", () => [[1, 2], [3, 4]]);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("B2");
    range.formulas = [["=MATRIX()"]];
    await context.sync();
    expect(storage.getCell("Sheet1", "B2").value).toBe(1);
    expect(storage.getCell("Sheet1", "C2").value).toBe(2);
    expect(storage.getCell("Sheet1", "B3").value).toBe(3);
    expect(storage.getCell("Sheet1", "C3").value).toBe(4);
  });

  test("write then read in same run works after sync", async () => {
    const { context, cf } = createContext();
    cf.associate("DOUBLE", (n: number) => n * 2);
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const writeRange = sheet.getRange("A1");
    writeRange.formulas = [["=DOUBLE(5)"]];
    await context.sync();
    const readRange = sheet.getRange("A1");
    readRange.load("values");
    await context.sync();
    expect(readRange.values).toEqual([[10]]);
  });

  test("clear is queued and executed on sync", async () => {
    const { context, storage } = createContext();
    storage.setValue("Sheet1", "A1", 42);
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
    range.clear();
    // Before sync, value should still be there
    expect(storage.getCell("Sheet1", "A1").value).toBe(42);
    await context.sync();
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
  });
});
