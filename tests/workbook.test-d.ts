import { expectTypeOf } from "vitest";
import { MockWorkbook } from "../src/workbook.js";
import { MockWorksheetCollection } from "../src/worksheet-collection.js";
import { MockWorksheet } from "../src/worksheet.js";
import { MockRange } from "../src/range.js";

declare const mockWorkbook: MockWorkbook;
declare const mockCollection: MockWorksheetCollection;
declare const mockWorksheet: MockWorksheet;

// Workbook: worksheets and getSelectedRange return mock types instead of Excel types,
// so we verify the property/method existence and parameter compatibility separately.
expectTypeOf(mockWorkbook).toHaveProperty("worksheets");
expectTypeOf(mockWorkbook).toHaveProperty("getSelectedRange");
expectTypeOf(mockWorkbook.getSelectedRange).returns.toEqualTypeOf<MockRange>();

// WorksheetCollection: methods return MockWorksheet instead of Excel.Worksheet,
// so we verify method signatures individually.
expectTypeOf(mockCollection.getActiveWorksheet).returns.toEqualTypeOf<MockWorksheet>();
expectTypeOf(mockCollection.getItem).parameter(0).toBeString();
expectTypeOf(mockCollection.getItem).returns.toEqualTypeOf<MockWorksheet>();
expectTypeOf(mockCollection.add).parameter(0).toBeString();
expectTypeOf(mockCollection.add).returns.toEqualTypeOf<MockWorksheet>();

// Worksheet: getRange returns MockRange instead of Excel.Range.
// Verify property types match the real API.
type ImplementedWorksheet = Pick<Excel.Worksheet, "name" | "id">;
expectTypeOf(mockWorksheet).toMatchTypeOf<ImplementedWorksheet>();

expectTypeOf(mockWorksheet.getRange).parameter(0).toEqualTypeOf<string>();
expectTypeOf(mockWorksheet.getRange).returns.toEqualTypeOf<MockRange>();
