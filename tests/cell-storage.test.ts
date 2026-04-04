import { describe, expect, test } from "vitest";
import { CellStorage } from "../src/cell-storage";

describe("CellStorage", () => {
  test("returns empty string for uninitialized cell", () => {
    const storage = new CellStorage();
    const cell = storage.getCell("Sheet1", "A1");
    expect(cell.value).toBe("");
    expect(cell.formula).toBe("");
  });

  test("stores and retrieves a value", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 42);
    expect(storage.getCell("Sheet1", "A1").value).toBe(42);
    expect(storage.getCell("Sheet1", "A1").formula).toBe("");
  });

  test("stores and retrieves a formula with value", () => {
    const storage = new CellStorage();
    storage.setFormula("Sheet1", "A1", "=ADD(1,2)", 3);
    const cell = storage.getCell("Sheet1", "A1");
    expect(cell.value).toBe(3);
    expect(cell.formula).toBe("=ADD(1,2)");
  });

  test("setValue overwrites previous formula", () => {
    const storage = new CellStorage();
    storage.setFormula("Sheet1", "A1", "=ADD(1,2)", 3);
    storage.setValue("Sheet1", "A1", "hello");
    const cell = storage.getCell("Sheet1", "A1");
    expect(cell.value).toBe("hello");
    expect(cell.formula).toBe("");
  });

  test("different sheets are independent", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet2", "A1", 2);
    expect(storage.getCell("Sheet1", "A1").value).toBe(1);
    expect(storage.getCell("Sheet2", "A1").value).toBe(2);
  });

  test("clear removes cell content", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 42);
    storage.clear("Sheet1", "A1");
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
  });

  test("clearSheet removes all cells in a sheet", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet1", "B1", 2);
    storage.clearSheet("Sheet1");
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
    expect(storage.getCell("Sheet1", "B1").value).toBe("");
  });

  test("clearAll removes all cells across all sheets", () => {
    const storage = new CellStorage();
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet2", "A1", 2);
    storage.clearAll();
    expect(storage.getCell("Sheet1", "A1").value).toBe("");
    expect(storage.getCell("Sheet2", "A1").value).toBe("");
  });
});
