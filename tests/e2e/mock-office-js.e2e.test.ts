import { test, expect } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const mockBundle = fs.readFileSync(
  path.join(__dirname, "../../dist/office.js"),
  "utf-8"
);

test.beforeEach(async ({ page }) => {
  await page.route("**/appsforoffice.microsoft.com/**", async (route) => {
    await route.fulfill({
      contentType: "application/javascript",
      body: mockBundle,
    });
  });

  await page.goto("/taskpane.html");
  await page.waitForFunction(() => (window as any).MockOfficeJs !== undefined);
});

test("Excel.run can read cell values set via mock", async ({ page }) => {
  const value = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    MockOfficeJs.excel.setCell("Sheet1", "A1", 42);

    let result: number[][] = [];
    await (window as any).Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("A1");
      range.load("values");
      await context.sync();
      result = range.values;
    });
    return result;
  });

  expect(value).toEqual([[42]]);
});

test("CustomFunctions.associate registers functions and formulas evaluate", async ({
  page,
}) => {
  const value = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.associate("ADD", (a: number, b: number) => a + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: "=ADD(2, 3)" });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(value).toBe(5);
});

test("worksheet operations: add and switch active worksheet", async ({
  page,
}) => {
  const sheetName = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    MockOfficeJs.excel.addWorksheet("Sheet2");
    MockOfficeJs.excel.setActiveWorksheet("Sheet2");

    let name = "";
    await (window as any).Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      name = sheet.name;
    });
    return name;
  });

  expect(sheetName).toBe("Sheet2");
});

test("accessing range properties without load/sync throws an error", async ({
  page,
}) => {
  const errorMessage = await page.evaluate(async () => {
    let error = "";
    await (window as any).Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("A1");
      try {
        const _values = range.values;
      } catch (e: any) {
        error = e.message;
      }
    });
    return error;
  });

  expect(errorMessage).toBeTruthy();
});

test("quoted numeric string argument is preserved as string", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.associate("ECHO", (val: any) => typeof val + ":" + val);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: '=ECHO("2023")' });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe("string:2023");
});

test("quoted string argument after comma has no leading space", async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    const CustomFunctions = (window as any).CustomFunctions;

    CustomFunctions.associate("JOIN", (a: any, b: any) => a + ":" + b);

    await MockOfficeJs.excel.setCell("Sheet1", "A1", { formula: '=JOIN(1, "hello")' });
    return MockOfficeJs.excel.getCell("Sheet1", "A1").value;
  });

  expect(result).toBe("1:hello");
});

test("MockOfficeJs.reset() clears cell values", async ({ page }) => {
  await page.evaluate(async () => {
    const MockOfficeJs = (window as any).MockOfficeJs;
    MockOfficeJs.excel.setCell("Sheet1", "A1", 99);
  });

  await page.evaluate(() => (window as any).MockOfficeJs.reset());

  const value = await page.evaluate(async () => {
    let result: any[][] = [];
    await (window as any).Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("A1");
      range.load("values");
      await context.sync();
      result = range.values;
    });
    return result;
  });

  expect(value).toEqual([[""]]);
});
