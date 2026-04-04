import { describe, expect, test } from "vitest";
import { MockWorkbook } from "../src/workbook";
import { MockWorksheetCollection } from "../src/worksheet-collection";
import { CellStorage } from "../src/cell-storage";
import { MockRange } from "../src/range";

describe("MockWorkbook / MockWorksheetCollection / MockWorksheet", () => {
  function createWorkbook() {
    const storage = new CellStorage();
    const pendingLoads: MockRange[] = [];
    const worksheets = new MockWorksheetCollection(storage, pendingLoads);
    const workbook = new MockWorkbook(storage, pendingLoads, worksheets);
    return { workbook, storage, pendingLoads };
  }

  test("default state has Sheet1 as active worksheet", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getActiveWorksheet();
    expect(sheet).toBeDefined();
  });

  test("getRange returns a MockRange", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1");
    expect(range).toBeInstanceOf(MockRange);
  });

  test("worksheet has name and id properties", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getActiveWorksheet();
    expect(sheet.name).toBe("Sheet1");
    expect(typeof sheet.id).toBe("string");
  });

  test("getItem returns worksheet by name", () => {
    const { workbook } = createWorkbook();
    const sheet = workbook.worksheets.getItem("Sheet1");
    expect(sheet).toBeDefined();
  });

  test("getItem throws for unknown sheet", () => {
    const { workbook } = createWorkbook();
    expect(() => workbook.worksheets.getItem("Unknown")).toThrow();
  });

  test("add creates a new worksheet", () => {
    const { workbook } = createWorkbook();
    workbook.worksheets.add("Sheet2");
    const sheet = workbook.worksheets.getItem("Sheet2");
    expect(sheet).toBeDefined();
    expect(sheet.name).toBe("Sheet2");
  });

  test("getSelectedRange returns the set selected range", () => {
    const { workbook } = createWorkbook();
    workbook.setSelectedRange("Sheet1", "B2");
    const range = workbook.getSelectedRange();
    expect(range).toBeInstanceOf(MockRange);
  });

  test("getSelectedRange throws when no selection is set", () => {
    const { workbook } = createWorkbook();
    expect(() => workbook.getSelectedRange()).toThrow();
  });

  test("setActiveWorksheet changes active worksheet", () => {
    const { workbook } = createWorkbook();
    workbook.worksheets.add("Sheet2");
    workbook.worksheets.setActiveWorksheet("Sheet2");
    const sheet = workbook.worksheets.getActiveWorksheet();
    expect(sheet.name).toBe("Sheet2");
  });

  test("reset restores default state", () => {
    const { workbook } = createWorkbook();
    workbook.worksheets.add("Sheet2");
    workbook.worksheets.setActiveWorksheet("Sheet2");
    workbook.worksheets.reset();
    expect(workbook.worksheets.getActiveWorksheet().name).toBe("Sheet1");
    expect(() => workbook.worksheets.getItem("Sheet2")).toThrow();
  });
});
