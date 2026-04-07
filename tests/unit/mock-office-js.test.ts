import { describe, expect, test } from "vitest";
import "../../src/index.js";

describe("mock-office-js global setup", () => {
  test("Excel global is defined with run method", () => {
    expect(Excel).toBeDefined();
    expect(Excel.run).toBeTypeOf("function");
  });

  test("Office global is defined with onReady method", () => {
    expect(Office).toBeDefined();
    expect(Office.onReady).toBeTypeOf("function");
  });

  test("CustomFunctions global is defined with associate method", () => {
    expect(CustomFunctions).toBeDefined();
    expect(CustomFunctions.associate).toBeTypeOf("function");
  });

  test("MockOfficeJs global is defined with excel and reset", () => {
    expect(MockOfficeJs).toBeDefined();
    expect(MockOfficeJs.excel).toBeDefined();
    expect(MockOfficeJs.reset).toBeTypeOf("function");
  });

  test("globals share state", async () => {
    MockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
    MockOfficeJs.reset();
  });
});
