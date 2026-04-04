import {
  parseAddress,
  cellAddressFromPosition,
  indexToColumnLetter,
  type RangePosition,
} from "./address.js";
import type { CellStorage } from "./cell-storage.js";

type PropertyName =
  | "values"
  | "formulas"
  | "text"
  | "numberFormat"
  | "address"
  | "rowCount"
  | "columnCount"
  | "rowIndex"
  | "columnIndex"
  | "hasSpill";

interface WriteEntry {
  property: "values" | "formulas" | "numberFormat";
  data: unknown[][];
}

export class MockRange {
  private readonly sheetName: string;
  private readonly addressStr: string;
  private readonly pos: RangePosition;
  private readonly pendingLoads: MockRange[];

  private requestedProps = new Set<PropertyName>();
  private loadedProps = new Set<PropertyName>();
  private cache = new Map<string, unknown>();
  private writeQueue: WriteEntry[] = [];
  private pendingClear = false;

  constructor(
    sheetName: string,
    address: string,
    _storage: CellStorage,
    pendingLoads: MockRange[],
  ) {
    this.sheetName = sheetName;
    this.addressStr = address;
    this.pos = parseAddress(address);
    this.pendingLoads = pendingLoads;
  }

  // --- load / sync ---

  load(properties: string | string[]): MockRange {
    const props = Array.isArray(properties)
      ? properties
      : properties.split(",").map((s) => s.trim());

    for (const p of props) {
      this.requestedProps.add(p as PropertyName);
    }
    this.pendingLoads.push(this);
    return this;
  }

  resolveLoads(storage: CellStorage): void {
    for (const prop of this.requestedProps) {
      switch (prop) {
        case "values":
          this.cache.set("values", this.readValues(storage));
          break;
        case "formulas":
          this.cache.set("formulas", this.readFormulas(storage));
          break;
        case "text":
          this.cache.set("text", this.readText(storage));
          break;
        case "numberFormat":
          this.cache.set("numberFormat", this.readNumberFormat());
          break;
        case "address":
          this.cache.set("address", `${this.sheetName}!${this.addressStr}`);
          break;
        case "rowCount":
          this.cache.set("rowCount", this.pos.endRow - this.pos.startRow + 1);
          break;
        case "columnCount":
          this.cache.set(
            "columnCount",
            this.pos.endCol - this.pos.startCol + 1,
          );
          break;
        case "rowIndex":
          this.cache.set("rowIndex", this.pos.startRow);
          break;
        case "columnIndex":
          this.cache.set("columnIndex", this.pos.startCol);
          break;
        case "hasSpill":
          this.cache.set("hasSpill", this.readHasSpill(storage));
          break;
      }
      this.loadedProps.add(prop);
    }
    this.requestedProps.clear();
  }

  // --- Property getters ---

  get values(): any[][] {
    return this.getLoaded("values") as any[][];
  }

  set values(data: any[][]) {
    this.writeQueue.push({ property: "values", data });
    this.pendingLoads.push(this);
  }

  get formulas(): any[][] {
    return this.getLoaded("formulas") as any[][];
  }

  set formulas(data: any[][]) {
    this.writeQueue.push({ property: "formulas", data });
    this.pendingLoads.push(this);
  }

  get text(): string[][] {
    return this.getLoaded("text") as string[][];
  }

  get numberFormat(): any[][] {
    return this.getLoaded("numberFormat") as any[][];
  }

  set numberFormat(data: any[][]) {
    this.writeQueue.push({ property: "numberFormat", data });
    this.pendingLoads.push(this);
  }

  get address(): string {
    return this.getLoaded("address") as string;
  }

  get rowCount(): number {
    return this.getLoaded("rowCount") as number;
  }

  get columnCount(): number {
    return this.getLoaded("columnCount") as number;
  }

  get rowIndex(): number {
    return this.getLoaded("rowIndex") as number;
  }

  get columnIndex(): number {
    return this.getLoaded("columnIndex") as number;
  }

  get hasSpill(): boolean {
    return this.getLoaded("hasSpill") as boolean;
  }

  // --- Methods ---

  getCell(row: number, col: number): MockRange {
    const targetRow = this.pos.startRow + row;
    const targetCol = this.pos.startCol + col;
    const addr = cellAddressFromPosition(targetRow, targetCol);
    return new MockRange(this.sheetName, addr, undefined as never, this.pendingLoads);
  }

  clear(_applyTo?: unknown): void {
    this.pendingClear = true;
    this.pendingLoads.push(this);
  }

  executeClear(storage: CellStorage): void {
    if (!this.pendingClear) return;
    for (let r = this.pos.startRow; r <= this.pos.endRow; r++) {
      for (let c = this.pos.startCol; c <= this.pos.endCol; c++) {
        storage.clear(this.sheetName, cellAddressFromPosition(r, c));
      }
    }
    this.pendingClear = false;
  }

  // --- Public helpers for RequestContext ---

  getWriteQueue(): WriteEntry[] {
    return this.writeQueue;
  }

  clearWriteQueue(): void {
    this.writeQueue = [];
  }

  getSheetName(): string {
    return this.sheetName;
  }

  getStartRow(): number {
    return this.pos.startRow;
  }

  getStartCol(): number {
    return this.pos.startCol;
  }

  // --- Private helpers ---

  private getLoaded(prop: string): unknown {
    if (!this.loadedProps.has(prop as PropertyName)) {
      throw new Error(
        `Property '${prop}' has not been loaded. Call load("${prop}") and context.sync() first.`,
      );
    }
    return this.cache.get(prop);
  }

  private readValues(storage: CellStorage): unknown[][] {
    const rows: unknown[][] = [];
    for (let r = this.pos.startRow; r <= this.pos.endRow; r++) {
      const row: unknown[] = [];
      for (let c = this.pos.startCol; c <= this.pos.endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        const cell = storage.getCell(this.sheetName, addr);
        row.push(cell.value);
      }
      rows.push(row);
    }
    return rows;
  }

  private readFormulas(storage: CellStorage): unknown[][] {
    const rows: unknown[][] = [];
    for (let r = this.pos.startRow; r <= this.pos.endRow; r++) {
      const row: unknown[] = [];
      for (let c = this.pos.startCol; c <= this.pos.endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        const cell = storage.getCell(this.sheetName, addr);
        row.push(cell.formula || cell.value);
      }
      rows.push(row);
    }
    return rows;
  }

  private readText(storage: CellStorage): string[][] {
    const rows: string[][] = [];
    for (let r = this.pos.startRow; r <= this.pos.endRow; r++) {
      const row: string[] = [];
      for (let c = this.pos.startCol; c <= this.pos.endCol; c++) {
        const addr = cellAddressFromPosition(r, c);
        const cell = storage.getCell(this.sheetName, addr);
        row.push(String(cell.value));
      }
      rows.push(row);
    }
    return rows;
  }

  private readNumberFormat(): unknown[][] {
    const rows: unknown[][] = [];
    for (let r = this.pos.startRow; r <= this.pos.endRow; r++) {
      const row: unknown[] = [];
      for (let c = this.pos.startCol; c <= this.pos.endCol; c++) {
        row.push("General");
      }
      rows.push(row);
    }
    return rows;
  }

  private readHasSpill(storage: CellStorage): boolean {
    const addr = cellAddressFromPosition(this.pos.startRow, this.pos.startCol);
    const cell = storage.getCell(this.sheetName, addr);
    return cell.value === "#SPILL!";
  }
}
