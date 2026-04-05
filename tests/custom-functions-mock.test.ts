import { describe, expect, test } from "vitest";
import { MockCustomFunctions } from "../src/custom-functions-mock.js";

describe("MockCustomFunctions", () => {
  test("associate registers a single function", () => {
    const cf = new MockCustomFunctions();
    const fn = (a: number, b: number) => a + b;
    cf.associate("ADD", fn);
    expect(cf.getFunction("ADD")).toBe(fn);
  });

  test("associate registers multiple functions via object", () => {
    const cf = new MockCustomFunctions();
    const add = (a: number, b: number) => a + b;
    const mul = (a: number, b: number) => a * b;
    cf.associate({ ADD: add, MUL: mul });
    expect(cf.getFunction("ADD")).toBe(add);
    expect(cf.getFunction("MUL")).toBe(mul);
  });

  test("function lookup is case-insensitive", () => {
    const cf = new MockCustomFunctions();
    const fn = () => 1;
    cf.associate("GETPRICE", fn);
    expect(cf.getFunction("getprice")).toBe(fn);
    expect(cf.getFunction("GetPrice")).toBe(fn);
  });

  test("getFunction returns undefined for unregistered function", () => {
    const cf = new MockCustomFunctions();
    expect(cf.getFunction("UNKNOWN")).toBeUndefined();
  });

  test("namespace-prefixed function names work", () => {
    const cf = new MockCustomFunctions();
    const fn = () => 42;
    cf.associate("CONTOSO.ADD", fn);
    expect(cf.getFunction("CONTOSO.ADD")).toBe(fn);
    expect(cf.getFunction("contoso.add")).toBe(fn);
  });

  test("Error class stores code and message", () => {
    const cf = new MockCustomFunctions();
    const error = new cf.Error(cf.ErrorCode.invalidValue, "bad input");
    expect(error.code).toBe("#VALUE!");
    expect(error.message).toBe("bad input");
  });

  test("ErrorCode enum has correct values", () => {
    const cf = new MockCustomFunctions();
    expect(cf.ErrorCode.invalidValue).toBe("#VALUE!");
    expect(cf.ErrorCode.notAvailable).toBe("#N/A");
    expect(cf.ErrorCode.divisionByZero).toBe("#DIV/0!");
    expect(cf.ErrorCode.invalidNumber).toBe("#NUM!");
    expect(cf.ErrorCode.nullReference).toBe("#NULL!");
    expect(cf.ErrorCode.invalidName).toBe("#NAME?");
    expect(cf.ErrorCode.invalidReference).toBe("#REF!");
  });

  test("reset clears all registered functions", () => {
    const cf = new MockCustomFunctions();
    cf.associate("ADD", () => 0);
    cf.reset();
    expect(cf.getFunction("ADD")).toBeUndefined();
  });

  test("loadMetadata stores parameter count from functions.json", () => {
    const cf = new MockCustomFunctions();
    cf.loadMetadata({
      functions: [
        {
          id: "ADD",
          name: "ADD",
          parameters: [
            { name: "first", type: "number" },
            { name: "second", type: "number" },
            { name: "third", type: "number", optional: true },
          ],
        },
      ],
    });
    expect(cf.getParameterCount("ADD")).toBe(3);
  });

  test("getParameterCount is case-insensitive", () => {
    const cf = new MockCustomFunctions();
    cf.loadMetadata({
      functions: [
        { id: "CONTOSO.ADD", name: "CONTOSO.ADD", parameters: [{ name: "a" }] },
      ],
    });
    expect(cf.getParameterCount("contoso.add")).toBe(1);
  });

  test("reset preserves metadata", () => {
    const cf = new MockCustomFunctions();
    cf.loadMetadata({
      functions: [
        { id: "ADD", name: "ADD", parameters: [{ name: "a" }, { name: "b" }] },
      ],
    });
    cf.associate("ADD", () => 0);
    cf.reset();
    expect(cf.getFunction("ADD")).toBeUndefined();
    expect(cf.getParameterCount("ADD")).toBe(2);
  });

  test("getParameterCount returns undefined for unknown function", () => {
    const cf = new MockCustomFunctions();
    expect(cf.getParameterCount("UNKNOWN")).toBeUndefined();
  });
});
