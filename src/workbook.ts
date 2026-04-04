import { CellStorage } from "./cell-storage";
import { MockWorksheetCollection } from "./worksheet-collection";
import { MockRange } from "./range";

export class MockWorkbook {
  readonly worksheets: MockWorksheetCollection;
  private _storage: CellStorage;
  private _pendingLoads: MockRange[];
  private _selectedSheet?: string;
  private _selectedAddress?: string;

  constructor(storage: CellStorage, pendingLoads: MockRange[], worksheets: MockWorksheetCollection) {
    this._storage = storage;
    this._pendingLoads = pendingLoads;
    this.worksheets = worksheets;
  }

  getSelectedRange(): MockRange {
    if (!this._selectedSheet || !this._selectedAddress) {
      throw new Error("No range is currently selected.");
    }
    return new MockRange(this._selectedSheet, this._selectedAddress, this._storage, this._pendingLoads);
  }

  setSelectedRange(sheet: string, address: string): void {
    this._selectedSheet = sheet;
    this._selectedAddress = address;
  }

  resetSelection(): void {
    this._selectedSheet = undefined;
    this._selectedAddress = undefined;
  }
}
