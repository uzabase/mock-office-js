export class MockCustomFunctions {
  private registry = new Map<string, Function>();
  private parameterCounts = new Map<string, number>();

  associate(
    idOrMappings: string | Record<string, Function>,
    fn?: Function,
  ): void {
    if (typeof idOrMappings === "string") {
      this.registry.set(idOrMappings.toUpperCase(), fn!);
      if (!this.parameterCounts.has(idOrMappings.toUpperCase())) {
        console.warn(
          `[mock-office-js] CustomFunctions.associate("${idOrMappings}"): no metadata loaded for this function. ` +
          `Call loadFunctionsMetadata() or loadMetadata() first. Without metadata, the function will return #NAME?.`
        );
      }
    } else {
      for (const [id, func] of Object.entries(idOrMappings)) {
        this.registry.set(id.toUpperCase(), func);
        if (!this.parameterCounts.has(id.toUpperCase())) {
          console.warn(
            `[mock-office-js] CustomFunctions.associate("${id}"): no metadata loaded for this function. ` +
            `Call loadFunctionsMetadata() or loadMetadata() first. Without metadata, the function will return #NAME?.`
          );
        }
      }
    }
  }

  getFunction(id: string): Function | undefined {
    return this.registry.get(id.toUpperCase());
  }

  loadMetadata(metadata: { functions: Array<{ id: string; name?: string; parameters?: Array<unknown> }> }): void {
    for (const fn of metadata.functions) {
      this.parameterCounts.set(fn.id.toUpperCase(), fn.parameters?.length ?? 0);
    }
  }

  getParameterCount(id: string): number | undefined {
    return this.parameterCounts.get(id.toUpperCase());
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
