import { describe, expect, test } from "vitest";
import { createMockEnvironment } from "../../src/setup.js";

describe("createMockEnvironment", () => {
  test("returns excel, office, customFunctions, and mockOfficeJs", () => {
    const env = createMockEnvironment();
    expect(env.excel).toBeDefined();
    expect(env.excel.run).toBeTypeOf("function");
    expect(env.office).toBeDefined();
    expect(env.office.onReady).toBeTypeOf("function");
    expect(env.customFunctions).toBeDefined();
    expect(env.customFunctions.associate).toBeTypeOf("function");
    expect(env.mockOfficeJs).toBeDefined();
    expect(env.mockOfficeJs.excel).toBeDefined();
    expect(env.mockOfficeJs.reset).toBeTypeOf("function");
  });
});

describe("mockOfficeJs.excel", () => {
  test("setCell and getCell with plain value", () => {
    const { mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(42);
  });

  test("setCell with formula evaluates custom function", async () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("uninitialized cell returns empty strings", () => {
    const { mockOfficeJs } = createMockEnvironment();
    const cell = mockOfficeJs.excel.getCell("Sheet1", "Z99");
    expect(cell.value).toBe("");
    expect(cell.formula).toBe("");
  });

  test("setCells writes multiple cells at once", () => {
    const { mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCells("Sheet1", "A1", [[1, 2], [3, 4]]);
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(1);
    expect(mockOfficeJs.excel.getCell("Sheet1", "B1").value).toBe(2);
    expect(mockOfficeJs.excel.getCell("Sheet1", "A2").value).toBe(3);
    expect(mockOfficeJs.excel.getCell("Sheet1", "B2").value).toBe(4);
  });

  test("getCells reads a range of cells", () => {
    const { mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCells("Sheet1", "A1", [[1, 2], [3, 4]]);
    const cells = mockOfficeJs.excel.getCells("Sheet1", "A1:B2");
    expect(cells[0][0].value).toBe(1);
    expect(cells[0][1].value).toBe(2);
    expect(cells[1][0].value).toBe(3);
    expect(cells[1][1].value).toBe(4);
  });

  test("setSelectedRange is used by Excel.run", async () => {
    const { excel, customFunctions, mockOfficeJs } = createMockEnvironment();
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

  test("addWorksheet persists across Excel.run calls", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.addWorksheet("Sheet2");
    await excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getItem("Sheet2");
      expect(sheet).toBeDefined();
    });
  });

  test("setActiveWorksheet changes active worksheet", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.addWorksheet("Sheet2");
    mockOfficeJs.excel.setActiveWorksheet("Sheet2");
    await excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      expect(sheet.name).toBe("Sheet2");
    });
  });

  test("spill collision returns #SPILL!", async () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("MATRIX", () => [[1, 2], [3, 4]]);
    mockOfficeJs.excel.setCell("Sheet1", "C2", 999);
    await mockOfficeJs.excel.setCell("Sheet1", "B2", { formula: "=MATRIX()" });
    expect(mockOfficeJs.excel.getCell("Sheet1", "B2").value).toBe("#SPILL!");
  });

  test("function name lookup is case-insensitive", async () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=add(1, 2)" });
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(3);
  });
});

describe("Excel.run shared state", () => {
  test("Excel.run reads data set via mockOfficeJs.excel.setCell", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    await excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });

  test("Excel.run formula write evaluates custom function", async () => {
    const { excel, customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("DOUBLE", (n: number) => n * 2);
    await excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.formulas = [["=DOUBLE(21)"]];
      await context.sync();
    });
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe(42);
  });
});

describe("Office.onReady", () => {
  test("calls callback with host and platform info", async () => {
    const { office } = createMockEnvironment();
    let receivedInfo: any;
    await office.onReady((info) => {
      receivedInfo = info;
    });
    expect(receivedInfo).toEqual({ host: "Excel", platform: "Web" });
  });

  test("returns a promise with host and platform info", async () => {
    const { office } = createMockEnvironment();
    const info = await office.onReady();
    expect(info).toEqual({ host: "Excel", platform: "Web" });
  });
});

describe("mockOfficeJs.reset", () => {
  test("clears all state", () => {
    const { customFunctions, mockOfficeJs } = createMockEnvironment();
    customFunctions.associate("ADD", (a: number, b: number) => a + b);
    mockOfficeJs.excel.setCell("Sheet1", "A1", 42);
    mockOfficeJs.excel.addWorksheet("Sheet2");
    mockOfficeJs.reset();
    expect(mockOfficeJs.excel.getCell("Sheet1", "A1").value).toBe("");
    expect(customFunctions.getFunction("ADD")).toBeUndefined();
  });

  test("clears selected range", async () => {
    const { excel, mockOfficeJs } = createMockEnvironment();
    mockOfficeJs.excel.setSelectedRange("Sheet1", "B1");
    mockOfficeJs.reset();
    await excel.run(async (context: any) => {
      expect(() => context.workbook.getSelectedRange()).toThrow();
    });
  });
});
