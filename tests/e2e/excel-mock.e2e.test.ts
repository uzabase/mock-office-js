import { test, expect } from "@playwright/test";

test.beforeEach(async ({ page }) => {
  await page.goto("/taskpane.html");
  // page.goto() waits for 'load' event by default, so mock script has executed
  await page.waitForFunction(() => (window as any).__mock__ !== undefined);
});

test("Excel.run can read cell values set via mock", async ({ page }) => {
  const value = await page.evaluate(async () => {
    const mock = (window as any).__mock__;
    mock.setCell("Sheet1", "A1", 42);

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

test("CustomFunctions.associate registers functions and formulas evaluate", async ({ page }) => {
  const value = await page.evaluate(async () => {
    const mock = (window as any).__mock__;

    // The fixture already registered ADD via associate()
    await mock.setCell("Sheet1", "A1", { formula: "=ADD(2, 3)" });

    return mock.getCell("Sheet1", "A1").value;
  });

  expect(value).toBe(5);
});

test("worksheet operations: add and switch active worksheet", async ({ page }) => {
  const sheetName = await page.evaluate(async () => {
    const mock = (window as any).__mock__;
    mock.addWorksheet("Sheet2");
    mock.setActiveWorksheet("Sheet2");

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

test("accessing range properties without load/sync throws an error", async ({ page }) => {
  const errorMessage = await page.evaluate(async () => {
    let error = "";
    await (window as any).Excel.run(async (context: any) => {
      const range = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange("A1");
      try {
        // Access values without calling load + sync
        const _values = range.values;
      } catch (e: any) {
        error = e.message;
      }
    });
    return error;
  });

  expect(errorMessage).toBeTruthy();
});

test("mock.reset() clears cell values", async ({ page }) => {
  await page.evaluate(async () => {
    const mock = (window as any).__mock__;
    mock.setCell("Sheet1", "A1", 99);
  });

  // Reset and verify value is cleared
  await page.evaluate(() => (window as any).__mock__.reset());

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
