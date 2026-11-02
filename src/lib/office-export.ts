import ExcelJS from "exceljs";
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import type { JSONContent } from "@tiptap/core";
import {
  calculateWorkbook,
  parseCellRange,
  type OfficeDocument,
  type Workbook,
  type WorkbookCell,
  type WorkbookSheet,
} from "@/lib/office";

const safeName = (title: string) => title.replace(/[^\w\- ]+/g, "").trim() || "document";

function download(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function cellValue(cell: WorkbookCell | null | undefined): string | number | boolean | null {
  const input = cell?.input ?? "";
  if (!input) return null;
  if (input.startsWith("=")) return null;
  if (/^-?(?:\d+\.?\d*|\.\d+)$/.test(input)) return Number(input);
  if (input === "TRUE") return true;
  if (input === "FALSE") return false;
  return input;
}

function argb(color: string | undefined, fallback = "FFFFFFFF") {
  const hex = (color ?? "").replace("#", "");
  return /^[\da-f]{6}$/i.test(hex) ? `FF${hex.toUpperCase()}` : fallback;
}

function applyCellStyle(target: ExcelJS.Cell, source: WorkbookCell | null | undefined) {
  if (!source) return;
  if (source.bold || source.italic || source.color) {
    target.font = {
      ...(source.bold !== undefined ? { bold: source.bold } : {}),
      ...(source.italic !== undefined ? { italic: source.italic } : {}),
      ...(source.color ? { color: { argb: argb(source.color, "FF1B1B1F") } } : {}),
    };
  }
  if (source.background) {
    target.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(source.background) } };
  }
  if (source.numberFormat)
    target.numFmt = source.numberFormat === "general" ? "General" : source.numberFormat;
}

function chartPng(
  sheet: WorkbookSheet,
  chart: WorkbookSheet["charts"][number],
  values: (string | number | boolean | null)[][],
): string | null {
  if (typeof document === "undefined") return null;
  const range = parseCellRange(chart.range);
  if (!range) return null;
  const sourceRows = values
    .slice(range.start.row, range.end.row + 1)
    .map((row) => row.slice(range.start.column, range.end.column + 1));
  if (sourceRows.length < 2) return null;

  const canvas = document.createElement("canvas");
  canvas.width = 900;
  canvas.height = 500;
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#202124";
  context.font = "bold 24px sans-serif";
  context.fillText(chart.title || sheet.name, 50, 42);

  const labels = sourceRows.slice(1).map((row, index) => String(row[0] ?? index + 1));
  const series = sourceRows[0]!.slice(1).map((heading, index) => ({
    name: String(heading || `Series ${index + 1}`),
    values: sourceRows.slice(1).map((row) => Number(row[index + 1] ?? 0)),
  }));
  const palette = ["#166d62", "#e17b48", "#3579a8", "#d2a72c", "#925f9d"];

  if (chart.kind === "pie") {
    const data = series[0]?.values ?? [];
    const total = data.reduce((sum, value) => sum + Math.max(0, value), 0);
    if (!total) return null;
    let angle = -Math.PI / 2;
    const centerX = 440;
    const centerY = 280;
    const radius = 175;
    data.forEach((value, index) => {
      const next = angle + (Math.max(0, value) / total) * Math.PI * 2;
      context.beginPath();
      context.moveTo(centerX, centerY);
      context.arc(centerX, centerY, radius, angle, next);
      context.closePath();
      context.fillStyle = palette[index % palette.length]!;
      context.fill();
      context.fillStyle = "#31343b";
      context.font = "16px sans-serif";
      context.fillText(labels[index] ?? "", centerX + radius + 24, 145 + index * 30);
      angle = next;
    });
  } else {
    const plot = { left: 80, top: 90, width: 760, height: 330 };
    const allValues = series.flatMap((item) => item.values).filter(Number.isFinite);
    const max = Math.max(1, ...allValues);
    context.strokeStyle = "#d9dde2";
    context.lineWidth = 1;
    for (let tick = 0; tick <= 4; tick++) {
      const y = plot.top + (plot.height * tick) / 4;
      context.beginPath();
      context.moveTo(plot.left, y);
      context.lineTo(plot.left + plot.width, y);
      context.stroke();
    }
    const groupWidth = plot.width / Math.max(1, labels.length);
    series.forEach((item, seriesIndex) => {
      context.strokeStyle = palette[seriesIndex % palette.length]!;
      context.fillStyle = palette[seriesIndex % palette.length]!;
      context.lineWidth = 4;
      context.beginPath();
      item.values.forEach((value, index) => {
        const x = plot.left + groupWidth * (index + 0.5);
        const y = plot.top + plot.height - (value / max) * plot.height;
        if (chart.kind === "line") {
          if (index === 0) context.moveTo(x, y);
          else context.lineTo(x, y);
          context.fillRect(x - 5, y - 5, 10, 10);
        } else {
          const barWidth = (groupWidth * 0.7) / series.length;
          context.fillRect(
            x - groupWidth * 0.35 + seriesIndex * barWidth,
            y,
            barWidth - 4,
            plot.top + plot.height - y,
          );
        }
      });
      if (chart.kind === "line") context.stroke();
    });
    context.fillStyle = "#4d5158";
    context.font = "14px sans-serif";
    labels.forEach((label, index) =>
      context.fillText(
        label.slice(0, 12),
        plot.left + groupWidth * (index + 0.2),
        plot.top + plot.height + 28,
      ),
    );
    series.forEach((item, index) => {
      context.fillStyle = palette[index % palette.length]!;
      context.fillRect(plot.left + index * 175, 455, 16, 16);
      context.fillStyle = "#4d5158";
      context.fillText(item.name.slice(0, 20), plot.left + index * 175 + 24, 469);
    });
  }
  return canvas.toDataURL("image/png").split(",")[1] ?? null;
}

export async function buildWorkbookBlob(workbook: Workbook) {
  const file = new ExcelJS.Workbook();
  file.creator = "SlideForge";
  file.created = new Date();
  const calculated = calculateWorkbook(workbook);

  for (const source of workbook.sheets) {
    const target = file.addWorksheet(source.name.slice(0, 31) || "Sheet");
    target.properties.defaultRowHeight = 22;
    target.columns = Array.from(
      { length: Math.max(12, source.columnCount, ...source.cells.map((row) => row.length)) },
      () => ({ width: 16 }),
    );
    source.cells.forEach((row, rowIndex) => {
      row.forEach((cell, columnIndex) => {
        const targetCell = target.getCell(rowIndex + 1, columnIndex + 1);
        targetCell.value = cell?.input.startsWith("=")
          ? { formula: cell.input.slice(1) }
          : cellValue(cell);
        applyCellStyle(targetCell, cell);
      });
    });

    for (const rule of source.conditionalFormats) {
      const style = {
        fill: {
          type: "pattern" as const,
          pattern: "solid" as const,
          fgColor: { argb: argb(rule.background) },
        },
        font: { color: { argb: argb(rule.color, "FF1B1B1F") } },
      };
      const excelRule =
        rule.operator === "containsText"
          ? {
              type: "containsText" as const,
              operator: "containsText" as const,
              text: rule.value,
              priority: 1,
              style,
            }
          : {
              type: "cellIs" as const,
              operator: rule.operator,
              formulae: [rule.value],
              priority: 1,
              style,
            };
      target.addConditionalFormatting({
        ref: rule.range,
        rules: [excelRule],
      });
    }

    const values = calculated.get(source.id) ?? [];
    source.charts.forEach((chart, index) => {
      const base64 = chartPng(source, chart, values);
      if (!base64) return;
      const imageId = file.addImage({ base64, extension: "png" });
      const range = parseCellRange(chart.range);
      target.addImage(imageId, {
        tl: { col: Math.max(4, (range?.end.column ?? 0) + 2), row: index * 18 },
        ext: { width: 450, height: 250 },
      });
    });
  }

  const buffer = await file.xlsx.writeBuffer();
  return new Blob([buffer as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export async function exportWorkbook(workbook: Workbook) {
  download(await buildWorkbookBlob(workbook), `${safeName(workbook.title)}.xlsx`);
}

function textRuns(nodes: JSONContent[] = []): TextRun[] {
  const runs: TextRun[] = [];
  for (const node of nodes) {
    if (node.type === "text" && node.text) {
      const marks = new Set((node.marks ?? []).map((mark) => mark.type));
      runs.push(
        new TextRun({
          text: node.text,
          bold: marks.has("bold"),
          italics: marks.has("italic"),
          ...(marks.has("underline") ? { underline: {} } : {}),
          strike: marks.has("strike"),
          ...(marks.has("code") ? { font: "Consolas" } : {}),
        }),
      );
    } else if (node.type === "hardBreak") {
      runs.push(new TextRun({ text: "\n" }));
    } else if (node.content) {
      runs.push(...textRuns(node.content));
    }
  }
  return runs;
}

function documentBlocks(node: JSONContent): (Paragraph | Table)[] {
  const blocks: (Paragraph | Table)[] = [];
  for (const child of node.content ?? []) {
    if (child.type === "table") {
      const rows = (child.content ?? []).map(
        (row) =>
          new TableRow({
            children: (row.content ?? []).map(
              (cell) =>
                new TableCell({
                  children: [new Paragraph({ children: textRuns(cell.content) })],
                  width: {
                    size: 100 / Math.max(1, row.content?.length ?? 1),
                    type: WidthType.PERCENTAGE,
                  },
                }),
            ),
          }),
      );
      blocks.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
      continue;
    }

    if (child.type === "bulletList" || child.type === "orderedList") {
      (child.content ?? []).forEach((item, index) => {
        const text = textRuns(item.content?.flatMap((part) => part.content ?? []) ?? []);
        if (child.type === "bulletList")
          blocks.push(new Paragraph({ children: text, bullet: { level: 0 } }));
        else
          blocks.push(
            new Paragraph({
              children: [new TextRun(`${index + 1}. `), ...text],
              indent: { left: 360 },
            }),
          );
      });
      continue;
    }

    if (child.type === "horizontalRule") {
      blocks.push(
        new Paragraph({
          border: { bottom: { color: "A8ADB4", size: 6, style: "single", space: 4 } },
        }),
      );
      continue;
    }

    const heading = child.type === "heading" ? child.attrs?.["level"] : undefined;
    const headingLevel =
      heading === 1
        ? HeadingLevel.HEADING_1
        : heading === 2
          ? HeadingLevel.HEADING_2
          : heading === 3
            ? HeadingLevel.HEADING_3
            : undefined;
    const text =
      child.type === "blockquote"
        ? textRuns(child.content?.flatMap((part) => part.content ?? []) ?? [])
        : textRuns(child.content);
    blocks.push(
      new Paragraph({
        children: text,
        ...(headingLevel ? { heading: headingLevel } : {}),
        ...(child.type === "blockquote" ? { indent: { left: 600 }, style: "Quote" } : {}),
        ...(child.type === "codeBlock" ? { style: "code" } : {}),
      }),
    );
  }
  return blocks.length ? blocks : [new Paragraph("")];
}

export async function buildDocumentBlob(document: OfficeDocument) {
  const file = new Document({
    title: document.title,
    sections: [{ children: documentBlocks(document.content) }],
  });
  return Packer.toBlob(file);
}

export async function exportDocument(document: OfficeDocument) {
  download(await buildDocumentBlob(document), `${safeName(document.title)}.docx`);
}
