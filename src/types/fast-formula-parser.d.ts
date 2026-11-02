declare module "fast-formula-parser" {
  interface CellReference {
    sheet?: string;
    row: number;
    col: number;
  }

  interface RangeReference {
    sheet?: string;
    from: { row: number; col: number };
    to: { row: number; col: number };
  }

  type FormulaReference = CellReference | RangeReference;

  interface ParserOptions {
    onCell?: (reference: CellReference) => unknown;
    onRange?: (reference: RangeReference) => unknown[][];
    onVariable?: (name: string, sheetName?: string) => FormulaReference | undefined;
  }

  class FormulaParser {
    constructor(options?: ParserOptions);
    parse(expression: string, position: { row: number; col: number; sheet: string }): { error: string | null; result: unknown };
  }

  export default FormulaParser;
}