import { createFileRoute, Link } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { Editor } from "@/components/slides/Editor";
import { getDeck, type Deck } from "@/lib/deck";

export const Route = createFileRoute("/edit/$deckId")({
  head: () => ({
    meta: [
      { title: "Editing — SlideForge" },
      { name: "description", content: "Design slides with text, shapes, images, links and AI help." },
      { property: "og:title", content: "Editing — SlideForge" },
      { property: "og:description", content: "Design slides with text, shapes, images, links and AI help." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: EditPage,
});

function EditPage() {
  const { deckId } = Route.useParams();
  const [deck, setDeck] = useState<Deck | null | undefined>(undefined);
  useEffect(() => setDeck(getDeck(deckId) ?? null), [deckId]);
  if (deck === undefined) return <div className="min-h-screen bg-background" />;
  if (deck === null)
    return (
      <div className="grid min-h-screen place-items-center text-center">
        <div><p className="mb-4 text-muted-foreground">Presentation not found.</p><Link to="/" className="text-primary underline">Back</Link></div>
      </div>
    );
  return <Editor key={deck.id} initial={deck} />;
}
