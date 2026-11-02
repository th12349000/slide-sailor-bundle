import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { Copy, Loader2, Play, Plus, Sparkles, Trash2, Wand2 } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { ScaledSlide, SlideStatic } from "@/components/slides/SlideRender";
import { deleteDeck, loadDecks, makeDeck, makeElement, makeSlide, uid, upsertDeck, type Deck, type Slide } from "@/lib/deck";
import { askSlideAi } from "@/lib/ai.functions";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "SlideForge — Your presentations" },
      { name: "description", content: "Create, edit and present slideshows with text, shapes, images, links and an AI co-designer." },
      { property: "og:title", content: "SlideForge — AI presentation maker" },
      { property: "og:description", content: "Create, edit and present slideshows with an AI co-designer." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

function Index() {
  const [decks, setDecks] = useState<Deck[]>([]);
  const [aiPrompt, setAiPrompt] = useState("");
  const [building, setBuilding] = useState(false);
  const navigate = useNavigate();
  useEffect(() => setDecks(loadDecks()), []);

  const create = () => {
    const d = makeDeck();
    upsertDeck(d);
    navigate({ to: "/edit/$deckId", params: { deckId: d.id } });
  };

  const buildWithAi = async () => {
    const topic = aiPrompt.trim();
    if (!topic || building) return;
    setBuilding(true);
    try {
      const res = await askSlideAi({
        data: {
          prompt: `Build a complete, polished presentation from scratch about: ${topic}. Use mode "replace_all" with 5-8 slides: a strong title slide, an agenda, content slides with one idea each, and a closing slide. Vary the layouts and use a bold, consistent color palette.`,
          deckJson: JSON.stringify({ title: topic, slides: [] }),
          currentIndex: 0,
        },
      });
      if (!res.ok) { toast.error(res.error); return; }
      const r = res.result;
      if (!r.slides.length) { toast.error("The AI didn't produce any slides — try rephrasing."); return; }
      const slides: Slide[] = r.slides.map((s) =>
        makeSlide({
          background: s.background,
          notes: s.notes,
          elements: s.elements.map((e) =>
            makeElement(e.type, {
              x: e.x, y: e.y, w: e.w, h: e.h, text: e.text, fontSize: e.fontSize, fontFamily: e.fontFamily, color: e.color,
              bold: e.bold, italic: e.italic, align: e.align, fill: e.fill, borderColor: e.borderColor, borderWidth: e.borderWidth,
              radius: e.radius, shape: e.shape, src: e.imageUrl,
              link: e.linkUrl ? { kind: "url", url: e.linkUrl } : e.linkSlideIndex !== null ? { kind: "slide", slideIndex: e.linkSlideIndex } : null,
            }),
          ),
        }),
      );
      const d = makeDeck(topic.slice(0, 60));
      d.slides = slides;
      upsertDeck(d);
      navigate({ to: "/edit/$deckId", params: { deckId: d.id } });
    } catch (e) {
      toast.error((e as Error).message || "Something went wrong.");
    } finally {
      setBuilding(false);
    }
  };

  return (
    <div className="min-h-screen grid-dots">
      <header className="mx-auto flex max-w-7xl items-center justify-between px-8 py-6">
        <div className="flex items-center gap-3">
          <div className="grid h-9 w-9 place-items-center rounded-lg bg-primary text-primary-foreground font-display font-extrabold">S</div>
          <span className="font-display text-xl font-bold">SlideForge</span>
        </div>
        <Button onClick={create}><Plus /> New presentation</Button>
      </header>

      <section className="mx-auto max-w-7xl px-8 pb-10 pt-12">
        <h1 className="max-w-3xl text-6xl font-extrabold leading-[0.95] tracking-tight">
          Decks that <span className="text-primary">link, move</span> and build themselves.
        </h1>
        <p className="mt-5 max-w-xl text-lg text-muted-foreground">
          Text, shapes, images and hyperlinks between slides and presentations — with an AI co-designer on call. Export to real PowerPoint files anytime.
        </p>
        <div className="mt-8 max-w-2xl rounded-2xl border border-border bg-card p-4">
          <div className="mb-2 flex items-center gap-2 text-sm font-semibold"><Wand2 className="h-4 w-4 text-primary" /> Build a whole deck with AI</div>
          <Textarea
            value={aiPrompt}
            onChange={(e) => setAiPrompt(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); buildWithAi(); } }}
            placeholder="e.g. A pitch for a solar-powered bike startup, aimed at investors…"
            rows={2}
            className="resize-none"
          />
          <div className="mt-3 flex gap-3">
            <Button onClick={buildWithAi} disabled={building || !aiPrompt.trim()}>
              {building ? <Loader2 className="animate-spin" /> : <Sparkles />} {building ? "Building your deck…" : "Build it"}
            </Button>
            <Button variant="outline" onClick={create}><Plus /> Start blank</Button>
          </div>
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-8 pb-24">
        <h2 className="mb-5 text-sm font-semibold uppercase tracking-widest text-muted-foreground">Your presentations</h2>
        {decks.length === 0 ? (
          <button onClick={create} className="grid aspect-video w-72 place-items-center rounded-xl border-2 border-dashed border-border text-muted-foreground hover:border-primary hover:text-primary">
            <span className="flex items-center gap-2"><Plus /> Create your first</span>
          </button>
        ) : (
          <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
            {decks.map((d) => (
              <div key={d.id} className="group rounded-xl border border-border bg-card p-3 transition hover:border-primary">
                <Link to="/edit/$deckId" params={{ deckId: d.id }}>
                  <ScaledSlide className="aspect-video rounded-lg">
                    {d.slides[0] && <SlideStatic slide={d.slides[0]} />}
                  </ScaledSlide>
                </Link>
                <div className="mt-3 flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <div className="truncate font-semibold">{d.title}</div>
                    <div className="text-xs text-muted-foreground">{d.slides.length} slides · {new Date(d.updatedAt).toLocaleDateString()}</div>
                  </div>
                  <div className="flex shrink-0 gap-1 opacity-70 group-hover:opacity-100">
                    <Button size="icon" variant="ghost" asChild><Link to="/view/$deckId" params={{ deckId: d.id }} search={{ slide: 1 }}><Play /></Link></Button>
                    <Button size="icon" variant="ghost" onClick={() => { upsertDeck({ ...structuredClone(d), id: uid(), title: d.title + " (copy)" }); setDecks(loadDecks()); }}><Copy /></Button>
                    <Button size="icon" variant="ghost" onClick={() => { if (confirm(`Delete "${d.title}"?`)) { deleteDeck(d.id); setDecks(loadDecks()); } }}><Trash2 /></Button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}
