import { cellAddressFromPosition, parseCellAddress } from "./address.js";

export interface CellState {
  value: unknown;
  formula: string;
  spilledFrom?: string;
}

interface SpillInfo {
  formula: string;
  results: unknown[][];
  spilledAddresses: string[];
}

const EMPTY_CELL: CellState = { value: "", formula: "" };

export class CellStorage {
  private sheets = new Map<string, Map<string, CellState>>();
  private spillOrigins = new Map<string, Map<string, SpillInfo>>();

  getCell(sheet: string, address: string): CellState {
    return this.sheets.get(sheet)?.get(address) ?? { ...EMPTY_CELL };
  }

  setValue(sheet: string, address: string, value: unknown): void {
    const existing = this.sheets.get(sheet)?.get(address);

    // If overwriting a spilled (non-origin) cell, invalidate the origin
    if (existing?.spilledFrom) {
      this.invalidateSpillOrigin(sheet, existing.spilledFrom);
    }

    this.ensureSheet(sheet).set(address, { value, formula: "" });
  }

  setFormula(sheet: string, address: string, formula: string, value: unknown): void {
    this.ensureSheet(sheet).set(address, { value, formula });
  }

  setFormulaWithSpill(
    sheet: string,
    address: string,
    formula: string,
    results: unknown[][],
  ): void {
    const origin = parseCellAddress(address);
    const sheetMap = this.ensureSheet(sheet);

    // Check for collisions: any non-empty, non-spilled-from-this-origin cell in the target range
    for (let r = 0; r < results.length; r++) {
      for (let c = 0; c < results[r].length; c++) {
        if (r === 0 && c === 0) continue; // skip origin
        const targetAddr = cellAddressFromPosition(origin.row + r, origin.col + c);
        const existing = sheetMap.get(targetAddr);
        if (existing && existing.spilledFrom !== address) {
          // Collision: set #SPILL! on origin, do not spill
          sheetMap.set(address, { value: "#SPILL!", formula });
          return;
        }
      }
    }

    // No collision: write all cells
    const spilledAddresses: string[] = [];

    for (let r = 0; r < results.length; r++) {
      for (let c = 0; c < results[r].length; c++) {
        const targetAddr = cellAddressFromPosition(origin.row + r, origin.col + c);
        if (r === 0 && c === 0) {
          sheetMap.set(targetAddr, { value: results[r][c], formula });
        } else {
          sheetMap.set(targetAddr, { value: results[r][c], formula: "", spilledFrom: address });
          spilledAddresses.push(targetAddr);
        }
      }
    }

    // Track spill info for cleanup
    const spillMap = this.ensureSpillMap(sheet);
    spillMap.set(address, { formula, results, spilledAddresses });
  }

  clear(sheet: string, address: string): void {
    // If clearing a spill origin, clear all spilled cells too
    const spillMap = this.spillOrigins.get(sheet);
    const spillInfo = spillMap?.get(address);
    if (spillInfo) {
      const sheetMap = this.sheets.get(sheet);
      for (const spilledAddr of spillInfo.spilledAddresses) {
        sheetMap?.delete(spilledAddr);
      }
      spillMap!.delete(address);
    }

    this.sheets.get(sheet)?.delete(address);
  }

  clearSheet(sheet: string): void {
    this.sheets.delete(sheet);
    this.spillOrigins.delete(sheet);
  }

  clearAll(): void {
    this.sheets.clear();
    this.spillOrigins.clear();
  }

  private invalidateSpillOrigin(sheet: string, originAddress: string): void {
    const spillMap = this.spillOrigins.get(sheet);
    const spillInfo = spillMap?.get(originAddress);
    if (!spillInfo) return;

    const sheetMap = this.sheets.get(sheet);
    if (!sheetMap) return;

    // Remove all spilled cells
    for (const spilledAddr of spillInfo.spilledAddresses) {
      sheetMap.delete(spilledAddr);
    }

    // Set origin to #SPILL!
    sheetMap.set(originAddress, { value: "#SPILL!", formula: spillInfo.formula });

    // Remove spill tracking
    spillMap!.delete(originAddress);
  }

  private ensureSheet(sheet: string): Map<string, CellState> {
    let sheetMap = this.sheets.get(sheet);
    if (!sheetMap) {
      sheetMap = new Map();
      this.sheets.set(sheet, sheetMap);
    }
    return sheetMap;
  }

  private ensureSpillMap(sheet: string): Map<string, SpillInfo> {
    let spillMap = this.spillOrigins.get(sheet);
    if (!spillMap) {
      spillMap = new Map();
      this.spillOrigins.set(sheet, spillMap);
    }
    return spillMap;
  }
}
