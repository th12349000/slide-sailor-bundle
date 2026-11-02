import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { DocumentEditor } from "@/components/office/DocumentEditor";
import { getDocument, type OfficeDocument } from "@/lib/office";

export const Route = createFileRoute("/document/$documentId")({
  head: () => ({ meta: [{ title: "Word document — SlideForge" }] }),
  component: DocumentPage,
});

function DocumentPage() {
  const { documentId } = Route.useParams();
  const [document, setDocument] = useState<OfficeDocument | null | undefined>(undefined);
  useEffect(() => setDocument(getDocument(documentId) ?? null), [documentId]);
  if (document === undefined) return <div className="min-h-screen bg-background" />;
  if (!document)
    return (
      <div className="grid min-h-screen place-items-center text-center">
        <div>
          <p className="mb-4 text-muted-foreground">Document not found.</p>
          <Link to="/" className="text-primary underline">
            Back to workspace
          </Link>
        </div>
      </div>
    );
  return <DocumentEditor key={document.id} initial={document} />;
}
