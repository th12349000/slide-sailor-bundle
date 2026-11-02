import { useEffect, useMemo, useState, type FormEvent, type KeyboardEvent } from "react";
import { Link } from "@tanstack/react-router";
import {
  ArrowLeft,
  BarChart3,
  Bold,
  Download,
  Italic,
  Plus,
  Sparkles,
  Trash2,
  Wand2,
} from "lucide-react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import {
  calculateWorkbook,
  cellAddress,
  columnLabel,
  makeSheet,
  matchesConditionalFormat,
  parseCellRange,
  setSheetCell,
  type ConditionalFormatRule,
  type Workbook,
  type WorkbookChart,
  type WorkbookSheet,
} from "@/lib/office";
import { generateOfficeLocally, localAiErrorMessage, localAiStatusLabel, type LocalAiStatus } from "@/lib/local-ai";
import { upsertWorkbook } from "@/lib/office";

const GRID_ROWS = 40;
const GRID_COLUMNS = 14;
const CHART_COLORS = ["#166d62", "#e17b48", "#3579a8", "#d2a72c", "#925f9d"];

type Selection = { row: number; column: number };

function normalizedColor(value: string | undefined, fallback: string) {
  return /^#[\da-f]{6}$/i.test(value ?? "") ? value! : fallback;
}

function shownValue(value: unknown, format: string | undefined) {
  if (typeof value !== "number") return value == null ? "" : String(value);
  if (format === "currency")
    return new Intl.NumberFormat(undefined, { style: "currency", currency: "USD" }).format(value);
  if (format === "percent")
    return `${new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }).format(value * 100)}%`;
  if (format === "number")
    return new Intl.NumberFormat(undefined, { maximumFractionDigits: 3 }).format(value);
  return String(value);
}

function rangeContains(range: string, row: number, column: number) {
  const parsed = parseCellRange(range);
  return (
    parsed !== null &&
    row >= parsed.start.row &&
    row <= parsed.end.row &&
    column >= parsed.start.column &&
    column <= parsed.end.column
  );
}

function chartTable(chart: WorkbookChart, values: (string | number | boolean | null)[][]) {
  const parsed = parseCellRange(chart.range);
  if (!parsed)
    return {
      labels: [] as string[],
      series: [] as string[],
      data: [] as Record<string, string | number>[],
    };
  const rows = values
    .slice(parsed.start.row, parsed.end.row + 1)
    .map((row) => row.slice(parsed.start.column, parsed.end.column + 1));
  if (rows.length < 2)
    return { labels: [], series: [], data: [] as Record<string, string | number>[] };
  const series = rows[0]!
    .slice(1)
    .map((heading, index) => String(heading || `Series ${index + 1}`));
  const labels = rows.slice(1).map((row, index) => String(row[0] ?? index + 1));
  const data = rows.slice(1).map((row, rowIndex) => {
    const point: Record<string, string | number> = { name: labels[rowIndex]! };
    series.forEach((name, seriesIndex) => {
      point[name] = Number(row[seriesIndex + 1] ?? 0);
    });
    return point;
  });
  return { labels, series, data };
}

function ChartPreview({
  chart,
  values,
}: {
  chart: WorkbookChart;
  values: (string | number | boolean | null)[][];
}) {
  const table = chartTable(chart, values);
  if (!table.data.length || !table.series.length) return null;
  return (
    <section className="min-w-0 border-t border-border px-4 py-4">
      <div className="mb-3 flex items-baseline justify-between gap-3">
        <h3 className="truncate text-sm font-semibold">{chart.title}</h3>
        <span className="text-xs text-muted-foreground">{chart.range}</span>
      </div>
      <div className="h-56 w-full">
        <ResponsiveContainer width="100%" height="100%">
          {chart.kind === "pie" ? (
            <PieChart>
              <Pie
                data={table.data.map((item) => ({
                  name: item["name"],
                  value: Number(item[table.series[0]!] ?? 0),
                }))}
                dataKey="value"
                nameKey="name"
                outerRadius={78}
                label
              >
                {table.data.map((_, index) => (
                  <Cell key={index} fill={CHART_COLORS[index % CHART_COLORS.length]!} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          ) : chart.kind === "line" ? (
            <LineChart data={table.data} margin={{ top: 8, right: 12, bottom: 4, left: -18 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
              <XAxis
                dataKey="name"
                tick={{ fill: "var(--color-muted-foreground)", fontSize: 11 }}
              />
              <YAxis tick={{ fill: "var(--color-muted-foreground)", fontSize: 11 }} />
              <Tooltip />
              {table.series.map((series, index) => (
                <Line
                  key={series}
                  dataKey={series}
                  stroke={CHART_COLORS[index % CHART_COLORS.length]!}
                  strokeWidth={2}
                  dot={false}
                />
              ))}
            </LineChart>
          ) : (
            <BarChart data={table.data} margin={{ top: 8, right: 12, bottom: 4, left: -18 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--color-border)" />
              <XAxis
                dataKey="name"
                tick={{ fill: "var(--color-muted-foreground)", fontSize: 11 }}
              />
              <YAxis tick={{ fill: "var(--color-muted-foreground)", fontSize: 11 }} />
              <Tooltip />
              {table.series.map((series, index) => (
                <Bar
                  key={series}
                  dataKey={series}
                  fill={CHART_COLORS[index % CHART_COLORS.length]!}
                />
              ))}
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </section>
  );
}

export function SpreadsheetEditor({ initial }: { initial: Workbook }) {
  const [workbook, setWorkbook] = useState(initial);
  const [activeSheetId, setActiveSheetId] = useState(initial.sheets[0]?.id ?? "");
  const [selection, setSelection] = useState<Selection>({ row: 0, column: 0 });
  const [aiPrompt, setAiPrompt] = useState("");
  const [aiBusy, setAiBusy] = useState(false);
  const [aiStatus, setAiStatus] = useState<LocalAiStatus | null>(null);
  const [showRuleForm, setShowRuleForm] = useState(false);
  const [ruleRange, setRuleRange] = useState("A1:A10");
  const [ruleOperator, setRuleOperator] =
    useState<ConditionalFormatRule["operator"]>("greaterThan");
  const [ruleValue, setRuleValue] = useState("0");
  const [ruleBackground, setRuleBackground] = useState("#f5c27a");
  const [ruleColor, setRuleColor] = useState("#402b12");
  const [showChartForm, setShowChartForm] = useState(false);
  const [chartTitle, setChartTitle] = useState("New chart");
  const [chartRange, setChartRange] = useState("A1:B5");
  const [chartKind, setChartKind] = useState<WorkbookChart["kind"]>("bar");
  const sheet = workbook.sheets.find((item) => item.id === activeSheetId) ?? workbook.sheets[0]!;
  const calculated = useMemo(() => calculateWorkbook(workbook), [workbook]);
  const values = calculated.get(sheet.id) ?? [];
  const selectedCell = sheet.cells[selection.row]?.[selection.column] ?? null;
  const cellInput = selectedCell?.input ?? "";

  useEffect(() => {
    const timer = window.setTimeout(() => upsertWorkbook(workbook), 450);
    return () => window.clearTimeout(timer);
  }, [workbook]);

  const updateSheet = (update: (current: WorkbookSheet) => WorkbookSheet) => {
    setWorkbook((current) => ({
      ...current,
      sheets: current.sheets.map((item) => (item.id === sheet.id ? update(item) : item)),
    }));
  };

  const updateInput = (input: string) => {
    updateSheet((current) => setSheetCell(current, selection.row, selection.column, { input }));
  };

  const moveSelection = (next: Selection) => {
    setSelection({ row: Math.max(0, next.row), column: Math.max(0, next.column) });
  };

  const addSheet = () => {
    const next = makeSheet(`Sheet ${workbook.sheets.length + 1}`);
    setWorkbook((current) => ({ ...current, sheets: [...current.sheets, next] }));
    setActiveSheetId(next.id);
    setSelection({ row: 0, column: 0 });
  };

  const applyAi = async (event: FormEvent) => {
    event.preventDefault();
    if (!aiPrompt.trim() || aiBusy) return;
    setAiBusy(true);
    try {
      const result = await generateOfficeLocally({
        kind: "spreadsheet",
        prompt: aiPrompt.trim(),
        dataJson: JSON.stringify(workbook),
        onStatus: setAiStatus,
      });
      if (!result.workbook) {
        toast.info(result.reply || "No workbook changes were returned.");
        return;
      }
      const sheets = result.workbook.sheets.map((item, index) => ({
        ...makeSheet(item.name || `Sheet ${index + 1}`),
        rowCount: Math.max(GRID_ROWS, item.cells.length),
        columnCount: Math.max(GRID_COLUMNS, ...item.cells.map((row) => row.length)),
        cells: item.cells.map((row) => row.map((input) => (input ? { input } : null))),
        conditionalFormats: item.conditionalFormats,
        charts: item.charts.map((chart) => ({
          ...chart,
          id: Math.random().toString(36).slice(2, 10),
        })),
      }));
      setWorkbook((current) => ({
        ...current,
        title: result.workbook!.title || current.title,
        sheets,
      }));
      setActiveSheetId(sheets[0]!.id);
      setSelection({ row: 0, column: 0 });
      setAiPrompt("");
      toast.success(result.reply || "Workbook updated.");
    } catch (error) {
      toast.error(localAiErrorMessage(error));
    } finally {
      setAiBusy(false);
      setAiStatus(null);
    }
  };

  const addConditionalRule = (event: FormEvent) => {
    event.preventDefault();
    const parsedRange = parseCellRange(ruleRange);
    if (!parsedRange) {
      toast.error("Enter a valid range such as B2:B20.");
      return;
    }
    updateSheet((current) => ({
      ...current,
      conditionalFormats: [
        ...current.conditionalFormats,
        {
          range: ruleRange.toUpperCase(),
          operator: ruleOperator,
          value: ruleValue,
          background: ruleBackground,
          color: ruleColor,
        },
      ],
    }));
    setShowRuleForm(false);
    toast.success("Conditional format added.");
  };

  const addChart = (event: FormEvent) => {
    event.preventDefault();
    if (!parseCellRange(chartRange)) {
      toast.error("Enter a valid chart range such as A1:B5.");
      return;
    }
    updateSheet((current) => ({
      ...current,
      charts: [
        ...current.charts,
        {
          id: Math.random().toString(36).slice(2, 10),
          title: chartTitle || "Chart",
          range: chartRange.toUpperCase(),
          kind: chartKind,
        },
      ],
    }));
    setShowChartForm(false);
  };

  const runExport = async () => {
    upsertWorkbook(workbook);
    try {
      const { exportWorkbook } = await import("@/lib/office-export");
      await exportWorkbook(workbook);
      toast.success("Excel workbook exported.");
    } catch (error) {
      toast.error(`Export failed: ${(error as Error).message}`);
    }
  };

  const activeCellRange = cellAddress(selection.row, selection.column);
  const visibleRows = Math.max(GRID_ROWS, sheet.rowCount);
  const visibleColumns = Math.max(GRID_COLUMNS, sheet.columnCount);

  return (
    <div className="flex h-screen min-h-[620px] flex-col bg-background">
      <header className="flex min-h-14 shrink-0 items-center gap-2 border-b border-border px-3">
        <Button variant="ghost" size="icon" asChild>
          <Link to="/" aria-label="Back to workspace">
            <ArrowLeft />
          </Link>
        </Button>
        <input
          aria-label="Workbook title"
          className="h-9 min-w-0 flex-1 border-0 bg-transparent font-display font-bold outline-none md:max-w-80"
          value={workbook.title}
          onChange={(event) =>
            setWorkbook((current) => ({ ...current, title: event.target.value }))
          }
        />
        <Button variant="outline" onClick={runExport}>
          <Download /> <span className="hidden sm:inline">Export .xlsx</span>
        </Button>
      </header>

      <div className="flex shrink-0 flex-wrap items-center gap-2 border-b border-border px-3 py-2">
        <Button
          size="sm"
          variant="outline"
          onClick={() =>
            updateSheet((current) => ({ ...current, rowCount: current.rowCount + 20 }))
          }
        >
          Add rows
        </Button>
        <Button
          size="sm"
          variant="outline"
          onClick={() =>
            updateSheet((current) => ({ ...current, columnCount: current.columnCount + 2 }))
          }
        >
          Add columns
        </Button>
        <span className="mx-1 h-6 w-px bg-border" />
        <span className="w-14 text-center font-mono text-xs text-muted-foreground">
          {activeCellRange}
        </span>
        <input
          aria-label="Formula bar"
          value={cellInput}
          onChange={(event) => updateInput(event.target.value)}
          className="h-8 min-w-0 flex-1 rounded border border-input bg-card px-2 font-mono text-sm outline-none focus:ring-1 focus:ring-ring"
          placeholder="Type a value or formula, for example =SUM(B2:B8)"
        />
        <Button
          size="icon"
          variant={selectedCell?.bold ? "secondary" : "ghost"}
          aria-label="Bold selected cell"
          onClick={() =>
            updateSheet((current) =>
              setSheetCell(current, selection.row, selection.column, { bold: !selectedCell?.bold }),
            )
          }
        >
          <Bold />
        </Button>
        <Button
          size="icon"
          variant={selectedCell?.italic ? "secondary" : "ghost"}
          aria-label="Italic selected cell"
          onClick={() =>
            updateSheet((current) =>
              setSheetCell(current, selection.row, selection.column, {
                italic: !selectedCell?.italic,
              }),
            )
          }
        >
          <Italic />
        </Button>
        <label
          className="grid h-9 w-9 cursor-pointer place-items-center rounded-md border border-border"
          title="Cell text color"
        >
          <span
            className="text-sm font-bold"
            style={{ color: selectedCell?.color ?? "currentColor" }}
          >
            A
          </span>
          <input
            aria-label="Cell text color"
            className="absolute h-px w-px opacity-0"
            type="color"
            value={normalizedColor(selectedCell?.color, "#202124")}
            onChange={(event) =>
              updateSheet((current) =>
                setSheetCell(current, selection.row, selection.column, {
                  color: event.target.value,
                }),
              )
            }
          />
        </label>
        <label
          className="grid h-9 w-9 cursor-pointer place-items-center rounded-md border border-border"
          title="Cell fill color"
        >
          <span
            className="h-4 w-4 border border-border"
            style={{ backgroundColor: selectedCell?.background ?? "transparent" }}
          />
          <input
            aria-label="Cell fill color"
            className="absolute h-px w-px opacity-0"
            type="color"
            value={normalizedColor(selectedCell?.background, "#ffffff")}
            onChange={(event) =>
              updateSheet((current) =>
                setSheetCell(current, selection.row, selection.column, {
                  background: event.target.value,
                }),
              )
            }
          />
        </label>
        <select
          aria-label="Number format"
          className="h-9 rounded border border-input bg-card px-2 text-sm"
          value={selectedCell?.numberFormat ?? "general"}
          onChange={(event) =>
            updateSheet((current) =>
              setSheetCell(current, selection.row, selection.column, {
                numberFormat: event.target.value,
              }),
            )
          }
        >
          <option value="general">General</option>
          <option value="number">Number</option>
          <option value="currency">Currency</option>
          <option value="percent">Percent</option>
        </select>
        <Button
          size="sm"
          variant={showRuleForm ? "secondary" : "ghost"}
          onClick={() => setShowRuleForm((open) => !open)}
        >
          Conditional format
        </Button>
        <Button
          size="sm"
          variant={showChartForm ? "secondary" : "ghost"}
          onClick={() => setShowChartForm((open) => !open)}
        >
          <BarChart3 /> Chart
        </Button>
      </div>

      {(showRuleForm || showChartForm) && (
        <div className="flex shrink-0 flex-wrap items-end gap-3 border-b border-border bg-card px-4 py-3">
          {showRuleForm && (
            <form className="flex flex-wrap items-end gap-2" onSubmit={addConditionalRule}>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Range
                <input
                  className="h-9 w-24 rounded border border-input bg-background px-2 text-sm text-foreground"
                  value={ruleRange}
                  onChange={(event) => setRuleRange(event.target.value)}
                />
              </label>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Condition
                <select
                  className="h-9 rounded border border-input bg-background px-2 text-sm text-foreground"
                  value={ruleOperator}
                  onChange={(event) =>
                    setRuleOperator(event.target.value as ConditionalFormatRule["operator"])
                  }
                >
                  <option value="greaterThan">Greater than</option>
                  <option value="lessThan">Less than</option>
                  <option value="equal">Equal to</option>
                  <option value="containsText">Contains text</option>
                </select>
              </label>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Value
                <input
                  className="h-9 w-28 rounded border border-input bg-background px-2 text-sm text-foreground"
                  value={ruleValue}
                  onChange={(event) => setRuleValue(event.target.value)}
                />
              </label>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Fill
                <input
                  aria-label="Conditional fill color"
                  className="h-9 w-12"
                  type="color"
                  value={ruleBackground}
                  onChange={(event) => setRuleBackground(event.target.value)}
                />
              </label>
              <Button size="sm" type="submit">
                Add rule
              </Button>
            </form>
          )}
          {showChartForm && (
            <form className="flex flex-wrap items-end gap-2" onSubmit={addChart}>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Title
                <input
                  className="h-9 w-36 rounded border border-input bg-background px-2 text-sm text-foreground"
                  value={chartTitle}
                  onChange={(event) => setChartTitle(event.target.value)}
                />
              </label>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Data range
                <input
                  className="h-9 w-24 rounded border border-input bg-background px-2 text-sm text-foreground"
                  value={chartRange}
                  onChange={(event) => setChartRange(event.target.value)}
                />
              </label>
              <label className="grid gap-1 text-xs text-muted-foreground">
                Type
                <select
                  className="h-9 rounded border border-input bg-background px-2 text-sm text-foreground"
                  value={chartKind}
                  onChange={(event) => setChartKind(event.target.value as WorkbookChart["kind"])}
                >
                  <option value="bar">Bar</option>
                  <option value="line">Line</option>
                  <option value="pie">Pie</option>
                </select>
              </label>
              <Button size="sm" type="submit">
                Add chart
              </Button>
            </form>
          )}
        </div>
      )}

      <div className="grid min-h-0 flex-1 grid-cols-1 xl:grid-cols-[minmax(0,1fr)_320px]">
        <main className="flex min-h-0 min-w-0 flex-col">
          <div className="min-h-0 flex-1 overflow-auto">
            <table className="office-grid border-separate border-spacing-0">
              <thead>
                <tr>
                  <th className="office-grid-corner sticky left-0 top-0 z-30"> </th>
                  {Array.from({ length: visibleColumns }, (_, column) => (
                    <th className="office-grid-column sticky top-0 z-20" key={column}>
                      {columnLabel(column)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: visibleRows }, (_, row) => (
                  <tr key={row}>
                    <th className="office-grid-row sticky left-0 z-10">{row + 1}</th>
                    {Array.from({ length: visibleColumns }, (_, column) => {
                      const cell = sheet.cells[row]?.[column] ?? null;
                      const selected = selection.row === row && selection.column === column;
                      const value = values[row]?.[column] ?? cell?.input ?? "";
                      const matchedRule = sheet.conditionalFormats.find(
                        (rule) =>
                          rangeContains(rule.range, row, column) &&
                          matchesConditionalFormat(value, rule),
                      );
                      const style = {
                        ...(cell?.background ? { backgroundColor: cell.background } : {}),
                        ...(cell?.color ? { color: cell.color } : {}),
                        ...(cell?.bold ? { fontWeight: 700 } : {}),
                        ...(cell?.italic ? { fontStyle: "italic" } : {}),
                        ...(matchedRule
                          ? { backgroundColor: matchedRule.background, color: matchedRule.color }
                          : {}),
                      };
                      return (
                        <td
                          className={`office-grid-cell ${selected ? "is-selected" : ""}`}
                          key={column}
                          onClick={() => moveSelection({ row, column })}
                          style={style}
                        >
                          {selected ? (
                            <input
                              aria-label={`Cell ${cellAddress(row, column)}`}
                              autoFocus
                              value={cell?.input ?? ""}
                              onChange={(event) => updateInput(event.target.value)}
                              onKeyDown={(event: KeyboardEvent<HTMLInputElement>) => {
                                if (event.key === "Enter") {
                                  event.preventDefault();
                                  moveSelection({ row: row + 1, column });
                                } else if (event.key === "Tab") {
                                  event.preventDefault();
                                  moveSelection({
                                    row,
                                    column: column + (event.shiftKey ? -1 : 1),
                                  });
                                } else if (event.key === "ArrowDown" && event.altKey) {
                                  moveSelection({ row: row + 1, column });
                                }
                              }}
                            />
                          ) : (
                            shownValue(value, cell?.numberFormat)
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {sheet.charts.map((chart) => (
            <ChartPreview key={chart.id} chart={chart} values={values} />
          ))}
          <footer className="flex shrink-0 items-center gap-1 overflow-x-auto border-t border-border px-3 py-2">
            {workbook.sheets.map((item) => (
              <button
                key={item.id}
                className={`shrink-0 border-b-2 px-3 py-1.5 text-sm ${item.id === sheet.id ? "border-primary font-semibold text-foreground" : "border-transparent text-muted-foreground hover:text-foreground"}`}
                onClick={() => {
                  setActiveSheetId(item.id);
                  setSelection({ row: 0, column: 0 });
                }}
              >
                {item.name}
              </button>
            ))}
            <Button size="icon" variant="ghost" aria-label="Add worksheet" onClick={addSheet}>
              <Plus />
            </Button>
            {sheet.charts.length > 0 && (
              <Button
                size="icon"
                variant="ghost"
                aria-label="Remove latest chart"
                onClick={() =>
                  updateSheet((current) => ({ ...current, charts: current.charts.slice(0, -1) }))
                }
              >
                <Trash2 />
              </Button>
            )}
          </footer>
        </main>

        <aside className="flex min-h-0 flex-col border-l border-border bg-card/60">
          <div className="flex items-center gap-2 border-b border-border px-4 py-3 text-sm font-semibold">
            <Sparkles className="h-4 w-4 text-primary" /> Spreadsheet assistant
          </div>
          <p className="px-4 pt-3 text-xs leading-relaxed text-muted-foreground">
            Runs in your browser. First use downloads the model once; then ask for formulas, charts,
            or a full workbook.
          </p>
          <form onSubmit={applyAi} className="flex flex-1 flex-col gap-3 p-4">
            <Textarea
              value={aiPrompt}
              onChange={(event) => setAiPrompt(event.target.value)}
              placeholder="Build a monthly budget with totals, variance formulas, and a category chart..."
              className="min-h-32 resize-y"
            />
            <Button type="submit" disabled={aiBusy || !aiPrompt.trim()}>
              {aiBusy ? <span className="animate-spin">◌</span> : <Wand2 />}{" "}
              {aiBusy ? (aiStatus ? localAiStatusLabel(aiStatus) : "Preparing local AI…") : "Generate or update"}
            </Button>
          </form>
          {sheet.conditionalFormats.length > 0 && (
            <div className="border-t border-border px-4 py-3">
              <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-muted-foreground">
                Rules
              </div>
              {sheet.conditionalFormats.map((rule, index) => (
                <div
                  key={`${rule.range}-${index}`}
                  className="flex items-center justify-between gap-2 py-1 text-xs"
                >
                  <span className="truncate">
                    {rule.range} {rule.operator} {rule.value}
                  </span>
                  <button
                    aria-label="Remove conditional rule"
                    onClick={() =>
                      updateSheet((current) => ({
                        ...current,
                        conditionalFormats: current.conditionalFormats.filter(
                          (_, ruleIndex) => ruleIndex !== index,
                        ),
                      }))
                    }
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </aside>
      </div>
    </div>
  );
}
