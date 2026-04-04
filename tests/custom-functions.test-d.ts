import { expectTypeOf } from "vitest";
import { MockCustomFunctions } from "../src/custom-functions-mock";

declare const cf: MockCustomFunctions;

expectTypeOf(cf.associate).toBeCallableWith("ADD", (a: number) => a);
expectTypeOf(cf.associate).toBeCallableWith({ ADD: (a: number) => a });

const error = new cf.Error(cf.ErrorCode.invalidValue, "msg");
expectTypeOf(error.code).toBeString();
expectTypeOf(error.message).toEqualTypeOf<string | undefined>();
