export class MockCustomFunctions {
  private registry = new Map<string, Function>();

  associate(
    idOrMappings: string | Record<string, Function>,
    fn?: Function,
  ): void {
    if (typeof idOrMappings === "string") {
      this.registry.set(idOrMappings.toUpperCase(), fn!);
    } else {
      for (const [id, func] of Object.entries(idOrMappings)) {
        this.registry.set(id.toUpperCase(), func);
      }
    }
  }

  getFunction(id: string): Function | undefined {
    return this.registry.get(id.toUpperCase());
  }

  reset(): void {
    this.registry.clear();
  }

  Error = MockCustomFunctionsError;

  ErrorCode = {
    invalidValue: "#VALUE!" as const,
    notAvailable: "#N/A" as const,
    divisionByZero: "#DIV/0!" as const,
    invalidNumber: "#NUM!" as const,
    nullReference: "#NULL!" as const,
    invalidName: "#NAME?" as const,
    invalidReference: "#REF!" as const,
  };
}

class MockCustomFunctionsError {
  code: string;
  message?: string;
  constructor(code: string, message?: string) {
    this.code = code;
    this.message = message;
  }
}
