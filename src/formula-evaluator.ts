import { CellStorage } from "./cell-storage.js";
import { MockCustomFunctions } from "./custom-functions-mock.js";
import { parseFormula } from "./formula-parser.js";

export class FormulaEvaluator {
  constructor(
    private _storage: CellStorage,
    private _customFunctions: MockCustomFunctions,
  ) {}

  async evaluateAndStore(sheet: string, address: string, formulaStr: string): Promise<void> {
    const parsed = parseFormula(formulaStr);
    if (!parsed) {
      this._storage.setValue(sheet, address, formulaStr);
      return;
    }
    const fn = this._customFunctions.getFunction(parsed.functionName);
    if (!fn) {
      this._storage.setFormula(sheet, address, formulaStr, "#NAME?");
      return;
    }
    const invocation = {
      address: `${sheet}!${address}`,
      functionName: parsed.functionName.toUpperCase(),
    };
    try {
      const result = await fn(...parsed.args, invocation);
      if (Array.isArray(result) && Array.isArray(result[0])) {
        this._storage.setFormulaWithSpill(sheet, address, formulaStr, result);
      } else {
        this._storage.setFormula(sheet, address, formulaStr, result);
      }
    } catch {
      this._storage.setFormula(sheet, address, formulaStr, "#VALUE!");
    }
  }
}
