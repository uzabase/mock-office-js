import type { CellState } from "./cell-storage.js";
import { createMockEnvironment } from "./setup.js";

const env = createMockEnvironment();

globalThis.Excel = env.excel as any;
globalThis.Office = env.office as any;
globalThis.CustomFunctions = env.customFunctions as any;
globalThis.MockOfficeJs = env.mockOfficeJs;

declare global {
  var MockOfficeJs: {
    excel: {
      setCell(sheet: string, address: string, value: unknown): void | Promise<void>;
      getCell(sheet: string, address: string): CellState;
      setCells(sheet: string, startAddress: string, values: unknown[][]): void;
      getCells(sheet: string, rangeAddress: string): CellState[][];
      setSelectedRange(sheet: string, address: string): void;
      setActiveWorksheet(sheet: string): void;
      addWorksheet(name: string): void;
      loadFunctionsMetadata(url: string): Promise<void>;
    };
    reset(): void;
  };
}
