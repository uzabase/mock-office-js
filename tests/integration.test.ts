import { describe, expect, test, vi, afterEach, afterAll, beforeAll } from "vitest";
import { ExcelMock } from "../src/excel-mock";

describe("E2E integration", () => {
  const mock = new ExcelMock();

  beforeAll(() => {
    vi.stubGlobal("Excel", mock.excel);
    vi.stubGlobal("CustomFunctions", mock.customFunctions);
  });

  afterEach(() => mock.reset());
  afterAll(() => vi.unstubAllGlobals());

  test("full flow: register function, write formula via Excel.run, verify value", async () => {
    CustomFunctions.associate("TRIPLE", (n: number) => n * 3);
    mock.setCell("Sheet1", "A1", 7);
    mock.setSelectedRange("Sheet1", "B1");

    await Excel.run(async (context: any) => {
      const source = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      source.load("values");
      await context.sync();
      const val = source.values[0][0];
      const selected = context.workbook.getSelectedRange();
      selected.formulas = [[`=TRIPLE(${val})`]];
      await context.sync();
    });

    expect(mock.getCell("Sheet1", "B1").value).toBe(21);
    expect(mock.getCell("Sheet1", "B1").formula).toBe("=TRIPLE(7)");

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("B1");
      range.load(["values", "formulas"]);
      await context.sync();
      expect(range.values).toEqual([[21]]);
      expect(range.formulas).toEqual([["=TRIPLE(7)"]]);
    });
  });

  test("spill flow: function returns 2D array, verify all cells", async () => {
    CustomFunctions.associate("TABLE", () => [
      ["Name", "Score"],
      ["Alice", 95],
      ["Bob", 87],
    ]);

    await mock.setCell("Sheet1", "A1", { formula: "=TABLE()" });

    expect(mock.getCell("Sheet1", "A1").value).toBe("Name");
    expect(mock.getCell("Sheet1", "B1").value).toBe("Score");
    expect(mock.getCell("Sheet1", "A2").value).toBe("Alice");
    expect(mock.getCell("Sheet1", "B2").value).toBe(95);
    expect(mock.getCell("Sheet1", "A3").value).toBe("Bob");
    expect(mock.getCell("Sheet1", "B3").value).toBe(87);
  });

  test("load/sync enforcement catches missing load", async () => {
    mock.setCell("Sheet1", "A1", 42);

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      expect(() => range.values).toThrow();
      range.load("values");
      expect(() => range.values).toThrow();
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });

  test("multiple Excel.run calls share cell state", async () => {
    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.values = [[100]];
      await context.sync();
    });

    await Excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[100]]);
    });
  });

  test("reset isolates test cases", async () => {
    mock.setCell("Sheet1", "A1", 42);
    mock.reset();
    expect(mock.getCell("Sheet1", "A1").value).toBe("");
  });
});
