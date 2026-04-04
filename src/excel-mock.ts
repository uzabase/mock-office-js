import { CellStorage, CellState } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { MockWorksheetCollection } from "./worksheet-collection.js";
import { MockRequestContext } from "./request-context.js";
import { FormulaEvaluator } from "./formula-evaluator.js";
import { parseAddress, cellAddressFromPosition, parseCellAddress } from "./address.js";
import { MockRange } from "./range.js";

export class ExcelMock {
  private _storage = new CellStorage();
  private _worksheets: MockWorksheetCollection;
  private _evaluator: FormulaEvaluator;
  readonly customFunctions = new MockCustomFunctions();
  private _selectedSheet?: string;
  private _selectedAddress?: string;

  constructor() {
    const dummyLoads: MockRange[] = [];
    this._worksheets = new MockWorksheetCollection(this._storage, dummyLoads);
    this._evaluator = new FormulaEvaluator(this._storage, this.customFunctions);
  }

  readonly excel = {
    run: async <T>(callback: (context: MockRequestContext) => Promise<T>): Promise<T> => {
      const context = new MockRequestContext(this._storage, this.customFunctions, this._worksheets);
      if (this._selectedSheet && this._selectedAddress) {
        context.workbook.setSelectedRange(this._selectedSheet, this._selectedAddress);
      }
      return await callback(context);
    },
  };

  setCell(sheet: string, address: string, value: unknown): void | Promise<void> {
    if (typeof value === "object" && value !== null && "formula" in value) {
      return this._evaluator.evaluateAndStore(sheet, address, (value as { formula: string }).formula);
    }
    this._storage.setValue(sheet, address, value);
  }

  setCells(sheet: string, startAddress: string, values: unknown[][]): void {
    const start = parseCellAddress(startAddress);
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const addr = cellAddressFromPosition(start.row + r, start.col + c);
        this._storage.setValue(sheet, addr, values[r][c]);
      }
    }
  }

  getCell(sheet: string, address: string): CellState {
    return this._storage.getCell(sheet, address);
  }

  getCells(sheet: string, rangeAddress: string): CellState[][] {
    const range = parseAddress(rangeAddress);
    const rows: CellState[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: CellState[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(this._storage.getCell(sheet, cellAddressFromPosition(r, c)));
      }
      rows.push(row);
    }
    return rows;
  }

  setSelectedRange(sheet: string, address: string): void {
    this._selectedSheet = sheet;
    this._selectedAddress = address;
  }

  setActiveWorksheet(sheet: string): void {
    this._worksheets.setActiveWorksheet(sheet);
  }

  addWorksheet(name: string): void {
    this._worksheets.add(name);
  }

  reset(): void {
    this._storage.clearAll();
    this.customFunctions.reset();
    this._worksheets.reset();
    this._selectedSheet = undefined;
    this._selectedAddress = undefined;
  }
}
