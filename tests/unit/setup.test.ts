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
