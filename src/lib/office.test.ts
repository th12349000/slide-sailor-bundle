import { describe, expect, it } from "vitest";
import {
  calculateWorkbook,
  cellAddress,
  makeWorkbook,
  matchesConditionalFormat,
  parseCellAddress,
  setSheetCell,
} from "@/lib/office";

describe("workbook calculations", () => {
  it("recalculates formulas from referenced cells", () => {
    const workbook = makeWorkbook();
    let sheet = workbook.sheets[0]!;
    sheet = setSheetCell(sheet, 0, 0, { input: "2" });
    sheet = setSheetCell(sheet, 1, 0, { input: "3" });
    sheet = setSheetCell(sheet, 2, 0, { input: "=SUM(A1:A2)" });
    workbook.sheets[0] = sheet;

    expect(calculateWorkbook(workbook).get(sheet.id)?.[2]?.[0]).toBe(5);
  });

  it("starts with a valid empty workbook", () => {
    const workbook = makeWorkbook();
    const values = calculateWorkbook(workbook).get(workbook.sheets[0]!.id)!;
    expect(values).toHaveLength(40);
    expect(values[0]).toHaveLength(14);
    expect(values[0]![0]).toBeNull();
  });

  it("converts spreadsheet coordinates", () => {
    expect(cellAddress(8, 27)).toBe("AB9");
    expect(parseCellAddress("AB9")).toEqual({ row: 8, column: 27 });
  });

  it("matches numeric and text conditional rules", () => {
    expect(
      matchesConditionalFormat(12, {
        range: "A1",
        operator: "greaterThan",
        value: "10",
        color: "#000",
        background: "#fff",
      }),
    ).toBe(true);
    expect(
      matchesConditionalFormat("Needs review", {
        range: "A1",
        operator: "containsText",
        value: "review",
        color: "#000",
        background: "#fff",
      }),
    ).toBe(true);
  });
});
