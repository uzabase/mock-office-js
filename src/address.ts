export function columnLetterToIndex(letter: string): number {
  let index = 0;
  const upper = letter.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

export function indexToColumnLetter(index: number): string {
  let letter = "";
  let n = index + 1;
  while (n > 0) {
    n--;
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26);
  }
  return letter;
}

export interface CellPosition {
  row: number;
  col: number;
}

export interface RangePosition {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function parseCellAddress(address: string): CellPosition {
  const cleaned = address.replace(/\$/g, "");
  const match = cleaned.match(/^([A-Za-z]+)(\d+)$/);
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return { col: columnLetterToIndex(match[1]), row: parseInt(match[2], 10) - 1 };
}

export function parseAddress(address: string): RangePosition {
  let cellPart = address;
  const sheetSepIndex = address.lastIndexOf("!");
  if (sheetSepIndex !== -1) cellPart = address.substring(sheetSepIndex + 1);
  const parts = cellPart.split(":");
  const start = parseCellAddress(parts[0]);
  if (parts.length === 1)
    return { startRow: start.row, startCol: start.col, endRow: start.row, endCol: start.col };
  const end = parseCellAddress(parts[1]);
  return { startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col };
}

export function cellAddressFromPosition(row: number, col: number): string {
  return `${indexToColumnLetter(col)}${row + 1}`;
}

export function resolveRangeAddresses(address: string): string[] {
  const range = parseAddress(address);
  const addresses: string[] = [];
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      addresses.push(cellAddressFromPosition(row, col));
    }
  }
  return addresses;
}
