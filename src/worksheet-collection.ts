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

  cloneWithPendingLoads(pendingLoads: MockRange[]): MockWorksheetCollection {
    const clone = Object.create(MockWorksheetCollection.prototype) as MockWorksheetCollection;
    (clone as any)._storage = this._storage;
    (clone as any)._pendingLoads = pendingLoads;
    (clone as any)._activeWorksheetName = this._activeWorksheetName;
    (clone as any)._nextId = this._nextId;
    // Create new worksheet instances with the new pendingLoads
    const newWorksheets = new Map<string, MockWorksheet>();
    for (const [name, ws] of this._worksheets) {
      newWorksheets.set(name, new MockWorksheet(ws.name, ws.id, this._storage, pendingLoads));
    }
    (clone as any)._worksheets = newWorksheets;
    return clone;
  }

  reset(): void {
    this._worksheets.clear();
    this._nextId = 1;
    this.add("Sheet1");
    this._activeWorksheetName = "Sheet1";
  }
}
