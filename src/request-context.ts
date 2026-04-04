import { CellStorage } from "./cell-storage";
import { MockCustomFunctions } from "./custom-functions-mock";
import { MockWorksheetCollection } from "./worksheet-collection";
import { MockWorkbook } from "./workbook";
import { MockRange } from "./range";
import { FormulaEvaluator } from "./formula-evaluator";
import { cellAddressFromPosition } from "./address";

export class MockRequestContext {
  readonly workbook: MockWorkbook;
  private _storage: CellStorage;
  private _evaluator: FormulaEvaluator;
  private _pendingLoads: MockRange[] = [];

  constructor(
    storage: CellStorage,
    customFunctions: MockCustomFunctions,
    _worksheets: MockWorksheetCollection,
  ) {
    this._storage = storage;
    this._evaluator = new FormulaEvaluator(storage, customFunctions);
    const contextWorksheets = _worksheets.cloneWithPendingLoads(this._pendingLoads);
    this.workbook = new MockWorkbook(
      storage,
      this._pendingLoads,
      contextWorksheets,
    );
  }

  async sync(): Promise<void> {
    // Process writes first, then resolve loads.
    // Use a snapshot so mutations during iteration are safe.
    const snapshot = [...this._pendingLoads];
    const processed = new Set<MockRange>();

    for (const range of snapshot) {
      if (processed.has(range)) continue;
      processed.add(range);
      await this.processRangeWrites(range);
    }

    // Resolve pending loads
    for (const range of snapshot) {
      range.resolveLoads(this._storage);
    }

    this._pendingLoads.length = 0;
  }

  private async processRangeWrites(range: MockRange): Promise<void> {
    const writes = range.getWriteQueue();
    for (const write of writes) {
      if (write.property === "values") {
        this.writeValues(range, write.data);
      } else if (write.property === "formulas") {
        await this.writeFormulas(range, write.data);
      } else if (write.property === "numberFormat") {
        // numberFormat writes are no-ops for now
      }
    }
    range.clearWriteQueue();

    // Handle pending clear
    range.executeClear(this._storage);
  }

  private writeValues(range: MockRange, data: unknown[][]): void {
    const sheetName = range.getSheetName();
    const startRow = range.getStartRow();
    const startCol = range.getStartCol();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const addr = cellAddressFromPosition(startRow + r, startCol + c);
        this._storage.setValue(sheetName, addr, data[r][c]);
      }
    }
  }

  private async writeFormulas(
    range: MockRange,
    data: unknown[][],
  ): Promise<void> {
    const sheetName = range.getSheetName();
    const startRow = range.getStartRow();
    const startCol = range.getStartCol();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const addr = cellAddressFromPosition(startRow + r, startCol + c);
        await this._evaluator.evaluateAndStore(
          sheetName,
          addr,
          String(data[r][c]),
        );
      }
    }
  }
}
