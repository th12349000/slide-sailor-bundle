import ExcelJS from "exceljs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { buildDocumentBlob, buildWorkbookBlob } from "@/lib/office-export";
import { makeDocument, makeWorkbook, setSheetCell } from "@/lib/office";

function readBlob(blob: Blob): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
}

describe("Office exports", () => {
  it("writes formulas and conditional formatting to XLSX", async () => {
    const workbook = makeWorkbook("Budget");
    let sheet = workbook.sheets[0]!;
    sheet = setSheetCell(sheet, 0, 0, { input: "4" });
    sheet = setSheetCell(sheet, 1, 0, { input: "6" });
    sheet = setSheetCell(sheet, 2, 0, { input: "=SUM(A1:A2)" });
    workbook.sheets[0] = {
      ...sheet,
      conditionalFormats: [
        {
          range: "A1:A3",
          operator: "greaterThan",
          value: "5",
          color: "#1b1b1f",
          background: "#c6f432",
        },
      ],
    };

    const blob = await buildWorkbookBlob(workbook);
    const bytes = await readBlob(blob);
    const parsed = new ExcelJS.Workbook();
    await parsed.xlsx.load(bytes);

    expect(parsed.getWorksheet("Sheet 1")?.getCell("A3").value).toMatchObject({
      formula: "SUM(A1:A2)",
    });
    const zip = await JSZip.loadAsync(bytes);
    const worksheet = await zip.file("xl/worksheets/sheet1.xml")!.async("text");
    expect(worksheet).toContain("conditionalFormatting");
    expect(blob.type).toBe("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  });

  it("writes rich text to a valid DOCX package", async () => {
    const document = makeDocument("Quarterly report");
    document.content = {
      type: "doc",
      content: [
        {
          type: "heading",
          attrs: { level: 1 },
          content: [{ type: "text", text: "Quarterly report" }],
        },
        { type: "paragraph", content: [{ type: "text", text: "Revenue increased." }] },
      ],
    };

    const blob = await buildDocumentBlob(document);
    const zip = await JSZip.loadAsync(await readBlob(blob));
    const xml = await zip.file("word/document.xml")!.async("text");

    expect(xml).toContain("Quarterly report");
    expect(xml).toContain("Revenue increased.");
    expect(blob.type).toContain(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
  });
});
