import { CellStorage } from "./cell-storage";
import { MockRange } from "./range";

export class MockWorksheet {
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];
  readonly name: string;
  readonly id: string;

  constructor(name: string, id: string, storage: CellStorage, pendingLoads: MockRange[]) {
    this.name = name;
    this.id = id;
    this._storage = storage;
    this._pendingLoads = pendingLoads;
  }

  getRange(address: string): MockRange {
    return new MockRange(this.name, address, this._storage, this._pendingLoads);
  }
}
