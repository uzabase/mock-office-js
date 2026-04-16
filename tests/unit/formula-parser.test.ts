import { describe, expect, test } from "vitest";
import { parseFormula } from "../../src/formula-parser.js";

describe("parseFormula", () => {
  test("parses function with no arguments", () => {
    expect(parseFormula("=NOW()")).toEqual({ functionName: "NOW", args: [] });
  });
  test("parses function with string argument", () => {
    expect(parseFormula('=GETPRICE("AAPL")')).toEqual({ functionName: "GETPRICE", args: ["AAPL"] });
  });
  test("parses function with number arguments", () => {
    expect(parseFormula("=ADD(1, 2)")).toEqual({ functionName: "ADD", args: [1, 2] });
  });
  test("parses function with negative number", () => {
    expect(parseFormula("=ADD(-1, 2.5)")).toEqual({ functionName: "ADD", args: [-1, 2.5] });
  });
  test("parses function with boolean arguments", () => {
    expect(parseFormula("=CHECK(TRUE, FALSE)")).toEqual({ functionName: "CHECK", args: [true, false] });
  });
  test("parses function with mixed argument types", () => {
    expect(parseFormula('=FUNC("hello", 42, TRUE)')).toEqual({ functionName: "FUNC", args: ["hello", 42, true] });
  });
  test("parses namespace-prefixed function name", () => {
    expect(parseFormula('=CONTOSO.GETPRICE("AAPL")')).toEqual({ functionName: "CONTOSO.GETPRICE", args: ["AAPL"] });
  });
  test("returns null for non-formula string", () => {
    expect(parseFormula("hello")).toBeNull();
  });
  test("parses cell references as string tokens", () => {
    expect(parseFormula("=SUM(A1:A5)")).toEqual({ functionName: "SUM", args: ["A1:A5"] });
  });
  test("parses string with escaped quotes", () => {
    expect(parseFormula('=FUNC("he said ""hi""")')).toEqual({ functionName: "FUNC", args: ['he said "hi"'] });
  });
  test("preserves quoted numeric string as string", () => {
    expect(parseFormula('=FUNC("2023")')).toEqual({ functionName: "FUNC", args: ["2023"] });
  });
  test("preserves empty quoted string", () => {
    expect(parseFormula('=FUNC("")')).toEqual({ functionName: "FUNC", args: [""] });
  });
  test("quoted string after numeric arg has no leading space", () => {
    expect(parseFormula('=FUNC(1, "hello")')).toEqual({ functionName: "FUNC", args: [1, "hello"] });
  });
  test("consecutive quoted strings have no leading space", () => {
    expect(parseFormula('=FUNC("a", "b")')).toEqual({ functionName: "FUNC", args: ["a", "b"] });
  });
  test("multiple quoted strings after comma have no leading space", () => {
    expect(parseFormula('=FUNC(1, "hello", "world")')).toEqual({ functionName: "FUNC", args: [1, "hello", "world"] });
  });
  test("parses args with no space after comma", () => {
    expect(parseFormula("=FUNC(1,2)")).toEqual({ functionName: "FUNC", args: [1, 2] });
  });
  test("quoted boolean string is preserved as string", () => {
    expect(parseFormula('=FUNC("TRUE")')).toEqual({ functionName: "FUNC", args: ["TRUE"] });
  });
  test("quoted boolean string after comma is preserved as string", () => {
    expect(parseFormula('=FUNC(1, "FALSE")')).toEqual({ functionName: "FUNC", args: [1, "FALSE"] });
  });
  test("zero is parsed as number", () => {
    expect(parseFormula("=FUNC(0)")).toEqual({ functionName: "FUNC", args: [0] });
  });
  test("single numeric argument", () => {
    expect(parseFormula("=FUNC(42)")).toEqual({ functionName: "FUNC", args: [42] });
  });
});
