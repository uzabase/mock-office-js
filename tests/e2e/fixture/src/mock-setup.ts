import { ExcelMock } from "mock-office-js";

const mock = new ExcelMock();
(window as any).Excel = mock.excel;
(window as any).Office = { onReady: (cb: () => void) => cb() };
(window as any).CustomFunctions = mock.customFunctions;
(window as any).__mock__ = mock;
