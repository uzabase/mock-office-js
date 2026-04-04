import { describe, expect, test } from "vitest";
import {
  columnLetterToIndex,
  indexToColumnLetter,
  parseAddress,
  parseCellAddress,
  resolveRangeAddresses,
} from "../src/address.js";

describe("columnLetterToIndex", () => {
  test("converts single letter columns", () => {
    expect(columnLetterToIndex("A")).toBe(0);
    expect(columnLetterToIndex("B")).toBe(1);
    expect(columnLetterToIndex("Z")).toBe(25);
  });
  test("converts multi-letter columns", () => {
    expect(columnLetterToIndex("AA")).toBe(26);
    expect(columnLetterToIndex("AB")).toBe(27);
    expect(columnLetterToIndex("AZ")).toBe(51);
  });
  test("is case-insensitive", () => {
    expect(columnLetterToIndex("a")).toBe(0);
    expect(columnLetterToIndex("aa")).toBe(26);
  });
});

describe("indexToColumnLetter", () => {
  test("converts single digit indices", () => {
    expect(indexToColumnLetter(0)).toBe("A");
    expect(indexToColumnLetter(1)).toBe("B");
    expect(indexToColumnLetter(25)).toBe("Z");
  });
  test("converts multi-letter indices", () => {
    expect(indexToColumnLetter(26)).toBe("AA");
    expect(indexToColumnLetter(27)).toBe("AB");
    expect(indexToColumnLetter(51)).toBe("AZ");
  });
});

describe("parseCellAddress", () => {
  test("parses simple cell address", () => {
    expect(parseCellAddress("A1")).toEqual({ row: 0, col: 0 });
    expect(parseCellAddress("B3")).toEqual({ row: 2, col: 1 });
    expect(parseCellAddress("AA10")).toEqual({ row: 9, col: 26 });
  });
  test("handles absolute references by stripping $", () => {
    expect(parseCellAddress("$A$1")).toEqual({ row: 0, col: 0 });
    expect(parseCellAddress("$B3")).toEqual({ row: 2, col: 1 });
  });
});

describe("parseAddress", () => {
  test("parses single cell address", () => {
    expect(parseAddress("A1")).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });
  test("parses range address", () => {
    expect(parseAddress("A1:C2")).toEqual({ startRow: 0, startCol: 0, endRow: 1, endCol: 2 });
  });
  test("parses sheet-qualified address", () => {
    expect(parseAddress("Sheet1!A1:B2")).toEqual({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
  });
  test("parses quoted sheet name with spaces", () => {
    expect(parseAddress("'My Sheet'!A1")).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
  });
});

describe("resolveRangeAddresses", () => {
  test("resolves single cell to one address", () => {
    expect(resolveRangeAddresses("A1")).toEqual(["A1"]);
  });
  test("resolves range to all cell addresses", () => {
    expect(resolveRangeAddresses("A1:B2")).toEqual(["A1", "B1", "A2", "B2"]);
  });
  test("resolves range addresses in row-major order", () => {
    expect(resolveRangeAddresses("B2:C3")).toEqual(["B2", "C2", "B3", "C3"]);
  });
});
