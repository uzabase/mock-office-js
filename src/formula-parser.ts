export interface ParsedFormula {
  functionName: string;
  args: unknown[];
}

export function parseFormula(formula: string): ParsedFormula | null {
  if (!formula.startsWith("=")) return null;
  const content = formula.substring(1);
  const parenOpen = content.indexOf("(");
  if (parenOpen === -1) return null;
  const functionName = content.substring(0, parenOpen).trim();
  if (!functionName) return null;
  const parenClose = content.lastIndexOf(")");
  if (parenClose === -1) return null;
  const argsString = content.substring(parenOpen + 1, parenClose).trim();
  const args = argsString === "" ? [] : parseArgs(argsString);
  return { functionName, args };
}

function parseArgs(argsString: string): unknown[] {
  const args: unknown[] = [];
  let current = "";
  let inString = false;
  let wasQuoted = false;
  let i = 0;
  while (i < argsString.length) {
    const char = argsString[i];
    if (inString) {
      if (char === '"') {
        if (i + 1 < argsString.length && argsString[i + 1] === '"') {
          current += '"';
          i += 2;
          continue;
        }
        inString = false;
        i++;
        continue;
      }
      current += char;
      i++;
      continue;
    }
    if (char === '"') { inString = true; wasQuoted = true; current = current.trimStart(); i++; continue; }
    if (char === ",") { args.push(wasQuoted ? current : parseArgValue(current.trim())); current = ""; wasQuoted = false; i++; continue; }
    current += char;
    i++;
  }
  if (current.trim() !== "" || wasQuoted) args.push(wasQuoted ? current : parseArgValue(current.trim()));
  return args;
}

function parseArgValue(value: string): unknown {
  if (value.startsWith('"') && value.endsWith('"')) return value.slice(1, -1);
  const upper = value.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  const num = Number(value);
  if (!isNaN(num) && value !== "") return num;
  return value; // Unresolved token (e.g., cell reference)
}
