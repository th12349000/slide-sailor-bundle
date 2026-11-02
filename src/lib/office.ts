import FormulaParser from "fast-formula-parser";
import type { JSONContent } from "@tiptap/core";
import { uid } from "@/lib/deck";

export interface WorkbookCell {
  input: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  background?: string;
  numberFormat?: string;
}

export interface ConditionalFormatRule {
  range: string;
  operator: "greaterThan" | "lessThan" | "equal" | "containsText";
  value: string;
  color: string;
  background: string;
}

export interface WorkbookChart {
  id: string;
  title: string;
  range: string;
  kind: "bar" | "line" | "pie";
}

export interface WorkbookSheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: (WorkbookCell | null)[][];
  conditionalFormats: ConditionalFormatRule[];
  charts: WorkbookChart[];
}

export interface Workbook {
  id: string;
  title: string;
  updatedAt: number;
  sheets: WorkbookSheet[];
}

export interface OfficeDocument {
  id: string;
  title: string;
  updatedAt: number;
  content: JSONContent;
}

const WORKBOOKS_KEY = "slideforge.workbooks.v1";
const DOCUMENTS_KEY = "slideforge.documents.v1";

export function makeSheet(name = "Sheet 1"): WorkbookSheet {
  return {
    id: uid(),
    name,
    rowCount: 40,
    columnCount: 14,
    cells: [],
    conditionalFormats: [],
    charts: [],
  };
}

export function makeWorkbook(title = "Untitled workbook"): Workbook {
  return { id: uid(), title, updatedAt: Date.now(), sheets: [makeSheet()] };
}

export function makeDocument(title = "Untitled document"): OfficeDocument {
  return {
    id: uid(),
    title,
    updatedAt: Date.now(),
    content: { type: "doc", content: [{ type: "paragraph" }] },
  };
}

function loadCollection<T>(key: string): T[] {
  if (typeof window === "undefined") return [];
  try {
    const data: unknown = JSON.parse(localStorage.getItem(key) || "[]");
    return Array.isArray(data) ? (data as T[]) : [];
  } catch {
    return [];
  }
}

function saveCollection<T>(key: string, items: T[]) {
  localStorage.setItem(key, JSON.stringify(items));
}

export const loadWorkbooks = () => loadCollection<Workbook>(WORKBOOKS_KEY);
export const loadDocuments = () => loadCollection<OfficeDocument>(DOCUMENTS_KEY);
export const getWorkbook = (id: string) => loadWorkbooks().find((item) => item.id === id);
export const getDocument = (id: string) => loadDocuments().find((item) => item.id === id);

export function upsertWorkbook(workbook: Workbook) {
  const items = loadWorkbooks();
  const index = items.findIndex((item) => item.id === workbook.id);
  const next = { ...workbook, updatedAt: Date.now() };
  if (index < 0) items.unshift(next);
  else items[index] = next;
  saveCollection(WORKBOOKS_KEY, items);
}

export function upsertDocument(document: OfficeDocument) {
  const items = loadDocuments();
  const index = items.findIndex((item) => item.id === document.id);
  const next = { ...document, updatedAt: Date.now() };
  if (index < 0) items.unshift(next);
  else items[index] = next;
  saveCollection(DOCUMENTS_KEY, items);
}

export function deleteWorkbook(id: string) {
  saveCollection(
    WORKBOOKS_KEY,
    loadWorkbooks().filter((item) => item.id !== id),
  );
}

export function deleteDocument(id: string) {
  saveCollection(
    DOCUMENTS_KEY,
    loadDocuments().filter((item) => item.id !== id),
  );
}

export function setSheetCell(
  sheet: WorkbookSheet,
  row: number,
  column: number,
  patch: Partial<WorkbookCell>,
): WorkbookSheet {
  const cells = sheet.cells.map((line) => [...line]);
  while (cells.length <= row) cells.push([]);
  while (cells[row]!.length <= column) cells[row]!.push(null);
  cells[row]![column] = { input: "", ...cells[row]![column], ...patch };
  return {
    ...sheet,
    rowCount: Math.max(sheet.rowCount, row + 1),
    columnCount: Math.max(sheet.columnCount, column + 1),
    cells,
  };
}

export function columnLabel(column: number): string {
  let value = column + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

export function cellAddress(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}

export function parseCellAddress(address: string): { row: number; column: number } | null {
  const match = address.toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) return null;
  let column = 0;
  for (const char of match[1]!) column = column * 26 + char.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

export function parseCellRange(
  range: string,
): { start: { row: number; column: number }; end: { row: number; column: number } } | null {
  const [startText, endText = startText] = range.replaceAll("$", "").split(":");
  const start = parseCellAddress(startText ?? "");
  const end = parseCellAddress(endText ?? "");
  return start && end ? { start, end } : null;
}

export function calculateWorkbook(
  workbook: Workbook,
): Map<string, (string | number | boolean | null)[][]> {
  const sheetsByName = new Map(workbook.sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]));
  const values = new Map<string, (string | number | boolean | null)[][]>(
    workbook.sheets.map((sheet) => [
      sheet.id,
      Array.from({ length: sheet.rowCount }, () => Array.from({ length: sheet.columnCount }, () => null)),
    ]),
  );
  const calculating = new Set<string>();

  const readCell = (sheetName: string, row: number, column: number): string | number | boolean | null => {
    const sheet = sheetsByName.get(sheetName.toLowerCase());
    if (!sheet || row < 1 || column < 1) return null;
    const key = `${sheet.id}:${row}:${column}`;
    const cached = values.get(sheet.id)?.[row - 1]?.[column - 1];
    if (cached !== null && cached !== undefined) return cached;
    if (calculating.has(key)) return "#CYCLE!";
    const input = sheet.cells[row - 1]?.[column - 1]?.input ?? "";
    if (!input) return null;
    if (!input.startsWith("=")) {
      const rawValue = input === "TRUE" ? true : input === "FALSE" ? false : /^-?(?:\d+\.?\d*|\.\d+)$/.test(input) ? Number(input) : input;
      values.get(sheet.id)![row - 1]![column - 1] = rawValue;
      return rawValue;
    }

    calculating.add(key);
    const parser = new FormulaParser({
      onCell: (reference) => readCell(reference.sheet ?? sheet.name, reference.row, reference.col),
      onRange: (reference) => Array.from(
        { length: reference.to.row - reference.from.row + 1 },
        (_, rowOffset) => Array.from(
          { length: reference.to.col - reference.from.col + 1 },
          (_, columnOffset) => readCell(reference.sheet ?? sheet.name, reference.from.row + rowOffset, reference.from.col + columnOffset),
        ),
      ),
      onVariable: (name, referencedSheet) => {
        const cleaned = name.replaceAll("$", "");
        const referenceSheet = referencedSheet ?? sheet.name;
        const range = parseCellRange(cleaned);
        if (range) {
          return {
            sheet: referenceSheet,
            from: { row: range.start.row + 1, col: range.start.column + 1 },
            to: { row: range.end.row + 1, col: range.end.column + 1 },
          };
        }
        const address = parseCellAddress(cleaned);
        return address ? { sheet: referenceSheet, row: address.row + 1, col: address.column + 1 } : undefined;
      },
    });
    const result = parser.parse(input.slice(1), { row, col: column, sheet: sheet.name });
    calculating.delete(key);
    const value = typeof result === "object" && result !== null && "result" in result
      ? "error" in result && result.error ? result.error : result.result
      : result;
    let normalized: string | number | boolean | null;
    if (value === null || value === undefined) normalized = null;
    else if (Array.isArray(value)) normalized = String(value[0]?.[0] ?? "");
    else if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") normalized = value;
    else normalized = String(value);
    values.get(sheet.id)![row - 1]![column - 1] = normalized;
    return normalized;
  };

  for (const sheet of workbook.sheets) {
    sheet.cells.forEach((row, rowIndex) => row.forEach((_, columnIndex) => readCell(sheet.name, rowIndex + 1, columnIndex + 1)));
  }
  return values;
}

export function matchesConditionalFormat(value: unknown, rule: ConditionalFormatRule): boolean {
  if (rule.operator === "containsText")
    return String(value ?? "")
      .toLowerCase()
      .includes(rule.value.toLowerCase());
  const left = Number(value);
  const right = Number(rule.value);
  if (!Number.isFinite(left) || !Number.isFinite(right)) return false;
  if (rule.operator === "greaterThan") return left > right;
  if (rule.operator === "lessThan") return left < right;
  return left === right;
}
