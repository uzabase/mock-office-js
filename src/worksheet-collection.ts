import { CellStorage } from "./cell-storage";
import { MockWorksheet } from "./worksheet";
import { MockRange } from "./range";

export class MockWorksheetCollection {
  private _worksheets = new Map<string, MockWorksheet>();
  private _activeWorksheetName: string;
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];
  private _nextId = 1;

  constructor(storage: CellStorage, pendingLoads: MockRange[]) {
    this._storage = storage;
    this._pendingLoads = pendingLoads;
    this.add("Sheet1");
    this._activeWorksheetName = "Sheet1";
  }

  getActiveWorksheet(): MockWorksheet {
    return this._worksheets.get(this._activeWorksheetName)!;
  }

  setActiveWorksheet(name: string): void {
    if (!this._worksheets.has(name)) throw new Error(`Worksheet '${name}' not found.`);
    this._activeWorksheetName = name;
  }

  getItem(name: string): MockWorksheet {
    const sheet = this._worksheets.get(name);
    if (!sheet) throw new Error(`Worksheet '${name}' not found.`);
    return sheet;
  }

  add(name: string): MockWorksheet {
    const id = `{${String(this._nextId++).padStart(8, "0")}}`;
    const sheet = new MockWorksheet(name, id, this._storage, this._pendingLoads);
    this._worksheets.set(name, sheet);
    return sheet;
  }

  reset(): void {
    this._worksheets.clear();
    this._nextId = 1;
    this.add("Sheet1");
    this._activeWorksheetName = "Sheet1";
  }
}
