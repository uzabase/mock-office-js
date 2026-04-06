import { describe, expect, test } from "vitest";
import { MockRange } from "../../src/range.js";
import { CellStorage } from "../../src/cell-storage.js";

describe("MockRange", () => {
  function createRange(address: string, storage?: CellStorage) {
    const s = storage ?? new CellStorage();
    const pendingLoads: MockRange[] = [];
    const range = new MockRange("Sheet1", address, s, pendingLoads);
    const sync = async () => {
      for (const r of pendingLoads) r.resolveLoads(s);
      pendingLoads.length = 0;
    };
    return { range, storage: s, sync };
  }

  test("accessing values without load throws", () => {
    const { range } = createRange("A1");
    expect(() => range.values).toThrow();
  });

  test("accessing values after load but before sync throws", () => {
    const { range } = createRange("A1");
    range.load("values");
    expect(() => range.values).toThrow();
  });

  test("accessing values after load and sync returns value", async () => {
    const { range, storage, sync } = createRange("A1");
    storage.setValue("Sheet1", "A1", 42);
    range.load("values");
    await sync();
    expect(range.values).toEqual([[42]]);
  });

  test("load accepts comma-separated string", async () => {
    const { range, storage, sync } = createRange("A1");
    storage.setValue("Sheet1", "A1", 42);
    range.load("values, formulas");
    await sync();
    expect(range.values).toEqual([[42]]);
    expect(range.formulas).toEqual([[42]]); // no formula, returns value
  });

  test("load accepts array of strings", async () => {
    const { range, storage, sync } = createRange("A1");
    storage.setValue("Sheet1", "A1", 42);
    range.load(["values", "address"]);
    await sync();
    expect(range.values).toEqual([[42]]);
    expect(range.address).toBe("Sheet1!A1");
  });

  test("load returns MockRange for chaining", () => {
    const { range } = createRange("A1");
    const result = range.load("values");
    expect(result).toBe(range);
  });

  test("address property returns sheet-qualified address", async () => {
    const { range, sync } = createRange("B3:D5");
    range.load("address");
    await sync();
    expect(range.address).toBe("Sheet1!B3:D5");
  });

  test("rowCount and columnCount for range", async () => {
    const { range, sync } = createRange("A1:C3");
    range.load(["rowCount", "columnCount"]);
    await sync();
    expect(range.rowCount).toBe(3);
    expect(range.columnCount).toBe(3);
  });

  test("rowIndex and columnIndex for range", async () => {
    const { range, sync } = createRange("C2");
    range.load(["rowIndex", "columnIndex"]);
    await sync();
    expect(range.rowIndex).toBe(1);
    expect(range.columnIndex).toBe(2);
  });

  test("values for multi-cell range", async () => {
    const { range, storage, sync } = createRange("A1:B2");
    storage.setValue("Sheet1", "A1", 1);
    storage.setValue("Sheet1", "B1", 2);
    storage.setValue("Sheet1", "A2", 3);
    storage.setValue("Sheet1", "B2", 4);
    range.load("values");
    await sync();
    expect(range.values).toEqual([[1, 2], [3, 4]]);
  });

  test("writing values does not require load", () => {
    const { range } = createRange("A1");
    expect(() => { range.values = [[42]]; }).not.toThrow();
  });

  test("getCell returns a MockRange for a single cell", async () => {
    const { range, storage, sync } = createRange("A1:C3");
    storage.setValue("Sheet1", "B2", 99);
    const cell = range.getCell(1, 1);
    cell.load("values");
    await sync();
    expect(cell.values).toEqual([[99]]);
  });
});
