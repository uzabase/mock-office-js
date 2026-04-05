import { describe, expect, test, vi, afterEach, afterAll, beforeAll } from "vitest";
import { ExcelMock } from "../src/excel-mock.js";

describe("ExcelMock", () => {
  const mock = new ExcelMock();

  beforeAll(() => {
    vi.stubGlobal("Excel", mock.excel);
    vi.stubGlobal("CustomFunctions", mock.customFunctions);
  });

  afterEach(() => mock.reset());
  afterAll(() => vi.unstubAllGlobals());

  test("registered function returns correct value via setCell", async () => {
    mock.customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mock.setCell("Sheet1", "A1", { formula: "=ADD(1, 2)" });
    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("setCell with plain value", () => {
    mock.setCell("Sheet1", "A1", 42);
    expect(mock.getCell("Sheet1", "A1").value).toBe(42);
  });

  test("uninitialized cell returns empty strings", () => {
    const cell = mock.getCell("Sheet1", "Z99");
    expect(cell.value).toBe("");
    expect(cell.formula).toBe("");
  });

  test("Excel.run works with context", async () => {
    mock.setCell("Sheet1", "A1", 42);
    await mock.excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.load("values");
      await context.sync();
      expect(range.values).toEqual([[42]]);
    });
  });

  test("Excel.run formula write evaluates custom function", async () => {
    mock.customFunctions.associate("DOUBLE", (n: number) => n * 2);
    await mock.excel.run(async (context: any) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      range.formulas = [["=DOUBLE(21)"]];
      await context.sync();
    });
    expect(mock.getCell("Sheet1", "A1").value).toBe(42);
  });

  test("task pane E2E: read cell, write formula to selected range", async () => {
    mock.customFunctions.associate("DOUBLE", (n: number) => n * 2);
    mock.setCell("Sheet1", "A1", 5);
    mock.setSelectedRange("Sheet1", "B1");
    await mock.excel.run(async (context: any) => {
      const source = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
      source.load("values");
      await context.sync();
      const selected = context.workbook.getSelectedRange();
      selected.formulas = [["=DOUBLE(5)"]];
      await context.sync();
    });
    expect(mock.getCell("Sheet1", "B1").value).toBe(10);
    expect(mock.getCell("Sheet1", "B1").formula).toBe("=DOUBLE(5)");
  });

  test("spill collision returns #SPILL!", async () => {
    mock.customFunctions.associate("MATRIX", () => [[1, 2], [3, 4]]);
    mock.setCell("Sheet1", "C2", 999);
    await mock.setCell("Sheet1", "B2", { formula: "=MATRIX()" });
    expect(mock.getCell("Sheet1", "B2").value).toBe("#SPILL!");
  });

  test("function name lookup is case-insensitive", async () => {
    mock.customFunctions.associate("ADD", (a: number, b: number) => a + b);
    await mock.setCell("Sheet1", "A1", { formula: "=add(1, 2)" });
    expect(mock.getCell("Sheet1", "A1").value).toBe(3);
  });

  test("addWorksheet persists across Excel.run calls", async () => {
    mock.addWorksheet("Sheet2");
    await mock.excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getItem("Sheet2");
      expect(sheet).toBeDefined();
    });
  });

  test("reset clears all state", async () => {
    mock.customFunctions.associate("ADD", (a: number, b: number) => a + b);
    mock.setCell("Sheet1", "A1", 42);
    mock.addWorksheet("Sheet2");
    mock.reset();
    expect(mock.getCell("Sheet1", "A1").value).toBe("");
    expect(mock.customFunctions.getFunction("ADD")).toBeUndefined();
  });

  test("setCells writes multiple cells at once", () => {
    mock.setCells("Sheet1", "A1", [[1, 2], [3, 4]]);
    expect(mock.getCell("Sheet1", "A1").value).toBe(1);
    expect(mock.getCell("Sheet1", "B1").value).toBe(2);
    expect(mock.getCell("Sheet1", "A2").value).toBe(3);
    expect(mock.getCell("Sheet1", "B2").value).toBe(4);
  });

  test("getCells reads a range of cells", () => {
    mock.setCells("Sheet1", "A1", [[1, 2], [3, 4]]);
    const cells = mock.getCells("Sheet1", "A1:B2");
    expect(cells[0][0].value).toBe(1);
    expect(cells[0][1].value).toBe(2);
    expect(cells[1][0].value).toBe(3);
    expect(cells[1][1].value).toBe(4);
  });
});

describe("ExcelMock.create", () => {
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
});
