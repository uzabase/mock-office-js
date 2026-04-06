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
