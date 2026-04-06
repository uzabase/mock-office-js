import { ExcelMock } from "mock-office-js";

const mock = new ExcelMock();
(window as any).Excel = mock.excel;
(window as any).CustomFunctions = mock.customFunctions;
(window as any).__mock__ = mock;

// Import functions AFTER globals are set, so associate() calls work
import("../functions/functions");
