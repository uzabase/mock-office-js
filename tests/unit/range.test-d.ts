import { expectTypeOf } from "vitest";
import { MockRange } from "../../src/range.js";

declare const mockRange: MockRange;

// Properties that can be directly compared for type conformance.
// Excluded from Pick:
//   - getCell: returns MockRange instead of Excel.Range (by design, mock does not implement all Excel.Range members)
type ImplementedRange = Pick<
  Excel.Range,
  | "values"
  | "formulas"
  | "address"
  | "rowCount"
  | "columnCount"
  | "columnIndex"
  | "rowIndex"
  | "text"
  | "numberFormat"
  | "hasSpill"
  | "clear"
>;

expectTypeOf(mockRange).toMatchTypeOf<ImplementedRange>();

// Verify getCell signature is compatible (parameter types match, return type is MockRange instead of Excel.Range)
expectTypeOf(mockRange.getCell).parameter(0).toBeNumber();
expectTypeOf(mockRange.getCell).parameter(1).toBeNumber();
expectTypeOf(mockRange.getCell).returns.toEqualTypeOf<MockRange>();
