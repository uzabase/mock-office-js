export interface CellState {
  value: unknown;
  formula: string;
  spilledFrom?: string;
}

const EMPTY_CELL: CellState = { value: "", formula: "" };

export class CellStorage {
  private sheets = new Map<string, Map<string, CellState>>();

  getCell(sheet: string, address: string): CellState {
    return this.sheets.get(sheet)?.get(address) ?? { ...EMPTY_CELL };
  }

  setValue(sheet: string, address: string, value: unknown): void {
    this.ensureSheet(sheet).set(address, { value, formula: "" });
  }

  setFormula(sheet: string, address: string, formula: string, value: unknown): void {
    this.ensureSheet(sheet).set(address, { value, formula });
  }

  clear(sheet: string, address: string): void {
    this.sheets.get(sheet)?.delete(address);
  }

  clearSheet(sheet: string): void {
    this.sheets.delete(sheet);
  }

  clearAll(): void {
    this.sheets.clear();
  }

  private ensureSheet(sheet: string): Map<string, CellState> {
    let sheetMap = this.sheets.get(sheet);
    if (!sheetMap) {
      sheetMap = new Map();
      this.sheets.set(sheet, sheetMap);
    }
    return sheetMap;
  }
}
