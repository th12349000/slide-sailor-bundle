import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { SpreadsheetEditor } from "@/components/office/SpreadsheetEditor";
import { getWorkbook, type Workbook } from "@/lib/office";

export const Route = createFileRoute("/spreadsheet/$workbookId")({
  head: () => ({ meta: [{ title: "Spreadsheet — SlideForge" }] }),
  component: SpreadsheetPage,
});

function SpreadsheetPage() {
  const { workbookId } = Route.useParams();
  const [workbook, setWorkbook] = useState<Workbook | null | undefined>(undefined);
  useEffect(() => setWorkbook(getWorkbook(workbookId) ?? null), [workbookId]);
  if (workbook === undefined) return <div className="min-h-screen bg-background" />;
  if (!workbook)
    return (
      <div className="grid min-h-screen place-items-center text-center">
        <div>
          <p className="mb-4 text-muted-foreground">Spreadsheet not found.</p>
          <Link to="/" className="text-primary underline">
            Back to workspace
          </Link>
        </div>
      </div>
    );
  return <SpreadsheetEditor key={workbook.id} initial={workbook} />;
}
