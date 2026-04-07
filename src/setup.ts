import { CellStorage, CellState } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { MockWorksheetCollection } from "./worksheet-collection.js";
import { MockRequestContext } from "./request-context.js";
import { FormulaEvaluator } from "./formula-evaluator.js";
import {
  parseAddress,
  cellAddressFromPosition,
  parseCellAddress,
} from "./address.js";
import { MockRange } from "./range.js";

export function createMockEnvironment() {
  const storage = new CellStorage();
  const customFunctions = new MockCustomFunctions();
  const dummyLoads: MockRange[] = [];
  const worksheets = new MockWorksheetCollection(storage, dummyLoads);
  const evaluator = new FormulaEvaluator(storage, customFunctions);

  let selectedSheet: string | undefined;
  let selectedAddress: string | undefined;

  const excel = {
    run: async <T>(
      callback: (context: MockRequestContext) => Promise<T>,
    ): Promise<T> => {
      const context = new MockRequestContext(
        storage,
        customFunctions,
        worksheets,
      );
      if (selectedSheet && selectedAddress) {
        context.workbook.setSelectedRange(selectedSheet, selectedAddress);
      }
      return await callback(context);
    },
  };

  const office = {
    onReady: (cb?: (info: { host: string; platform: string }) => void) => {
      const info = { host: "Excel", platform: "Web" };
      if (cb) cb(info);
      return Promise.resolve(info);
    },
    actions: {
      associate: () => {},
    },
  };

  const mockOfficeJs = {
    excel: {
      setCell(
        sheet: string,
        address: string,
        value: unknown,
      ): void | Promise<void> {
        if (typeof value === "object" && value !== null && "formula" in value) {
          return evaluator.evaluateAndStore(
            sheet,
            address,
            (value as { formula: string }).formula,
          );
        }
        storage.setValue(sheet, address, value);
      },

      getCell(sheet: string, address: string): CellState {
        return storage.getCell(sheet, address);
      },

      setCells(
        sheet: string,
        startAddress: string,
        values: unknown[][],
      ): void {
        const start = parseCellAddress(startAddress);
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const addr = cellAddressFromPosition(start.row + r, start.col + c);
            storage.setValue(sheet, addr, values[r][c]);
          }
        }
      },

      getCells(sheet: string, rangeAddress: string): CellState[][] {
        const range = parseAddress(rangeAddress);
        const rows: CellState[][] = [];
        for (let r = range.startRow; r <= range.endRow; r++) {
          const row: CellState[] = [];
          for (let c = range.startCol; c <= range.endCol; c++) {
            row.push(storage.getCell(sheet, cellAddressFromPosition(r, c)));
          }
          rows.push(row);
        }
        return rows;
      },

      setSelectedRange(sheet: string, address: string): void {
        selectedSheet = sheet;
        selectedAddress = address;
      },

      setActiveWorksheet(sheet: string): void {
        worksheets.setActiveWorksheet(sheet);
      },

      addWorksheet(name: string): void {
        worksheets.add(name);
      },

      async loadFunctionsMetadata(url: string): Promise<void> {
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(
            `Failed to fetch functions metadata: ${response.status}`,
          );
        }
        const metadata = await response.json();
        customFunctions.loadMetadata(metadata);
      },
    },

    reset(): void {
      storage.clearAll();
      customFunctions.reset();
      worksheets.reset();
      selectedSheet = undefined;
      selectedAddress = undefined;
    },
  };

  return { excel, office, customFunctions, mockOfficeJs };
}
