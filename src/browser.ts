import { ExcelMock } from "./excel-mock.js";

const mock = new ExcelMock();

// Replace Office.js globals with mock-office-js
(window as any).Excel = mock.excel;
(window as any).CustomFunctions = mock.customFunctions;
(window as any).Office = {
  onReady: (cb?: () => void) => {
    if (cb) cb();
  },
  actions: {
    associate: () => {},
  },
};

// Expose mock instance for test-side access
(window as any).__mock__ = mock;
