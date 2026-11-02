import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useEffect, useState } from "react";
import { Copy, FileSpreadsheet, FileText, Loader2, Play, Plus, Sparkles, Trash2, Wand2 } from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { ScaledSlide, SlideStatic } from "@/components/slides/SlideRender";
import { deleteDeck, loadDecks, makeDeck, makeElement, makeSlide, uid, upsertDeck, type Deck, type Slide } from "@/lib/deck";
import { makeSlidesFromAi } from "@/lib/ai-slide-mapper";
import { generateSlidesLocally, localAiErrorMessage, localAiStatusLabel, type LocalAiStatus } from "@/lib/local-ai";
import { deleteDocument, deleteWorkbook, loadDocuments, loadWorkbooks, makeDocument, makeWorkbook, upsertDocument, upsertWorkbook, type OfficeDocument, type Workbook } from "@/lib/office";

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
  const [workbooks, setWorkbooks] = useState<Workbook[]>([]);
  const [documents, setDocuments] = useState<OfficeDocument[]>([]);
  const [aiPrompt, setAiPrompt] = useState("");
  const [building, setBuilding] = useState(false);
  const [aiStatus, setAiStatus] = useState<LocalAiStatus | null>(null);
  const navigate = useNavigate();
  useEffect(() => {
    setDecks(loadDecks());
    setWorkbooks(loadWorkbooks());
    setDocuments(loadDocuments());
  }, []);

  const create = () => {
    const d = makeDeck();
    upsertDeck(d);
    navigate({ to: "/edit/$deckId", params: { deckId: d.id } });
  };

  const createWorkbook = () => {
    const workbook = makeWorkbook();
    upsertWorkbook(workbook);
    navigate({ to: "/spreadsheet/$workbookId", params: { workbookId: workbook.id } });
  };

  const createDocument = () => {
    const document = makeDocument();
    upsertDocument(document);
    navigate({ to: "/document/$documentId", params: { documentId: document.id } });
  };

  const refreshLibrary = () => {
    setDecks(loadDecks());
    setWorkbooks(loadWorkbooks());
    setDocuments(loadDocuments());
  };

  const buildWithAi = async () => {
    const topic = aiPrompt.trim();
    if (!topic || building) return;
    setBuilding(true);
    try {
      const result = await generateSlidesLocally({
        prompt: `Build a complete, polished presentation from scratch about: ${topic}. Use mode "replace_all" with 5-8 slides: a strong title slide, an agenda, content slides with one idea each, and a closing slide. Vary the layouts and use a bold, consistent color palette. Include relevant licensed photos when available.`,
        deckJson: JSON.stringify({ title: topic, slides: [] }),
        currentIndex: 0,
        onStatus: setAiStatus,
      });
      if (!result.slides.length) { toast.error("The local AI didn't produce any slides — try a shorter prompt."); return; }
      const slides: Slide[] = makeSlidesFromAi(result.slides);
      const d = makeDeck(topic.slice(0, 60));
      d.slides = slides;
      upsertDeck(d);
      navigate({ to: "/edit/$deckId", params: { deckId: d.id } });
    } catch (e) {
      toast.error(localAiErrorMessage(e));
    } finally {
      setBuilding(false);
      setAiStatus(null);
    }
  };

  return (
    <div className="min-h-screen grid-dots">
      <header className="mx-auto flex max-w-7xl items-center justify-between px-8 py-6">
        <div className="flex items-center gap-3">
          <div className="grid h-9 w-9 place-items-center rounded-lg bg-primary text-primary-foreground font-display font-extrabold">S</div>
          <span className="font-display text-xl font-bold">SlideForge</span>
        </div>
        <div className="flex flex-wrap justify-end gap-2">
          <Button onClick={create}><Plus /> <span className="hidden sm:inline">Presentation</span></Button>
          <Button variant="outline" onClick={createWorkbook}><FileSpreadsheet /> <span className="hidden sm:inline">Spreadsheet</span></Button>
          <Button variant="outline" onClick={createDocument}><FileText /> <span className="hidden sm:inline">Word document</span></Button>
        </div>
      </header>

      <section className="mx-auto max-w-7xl px-8 pb-10 pt-12">
        <h1 className="max-w-3xl text-6xl font-extrabold leading-[0.95] tracking-tight">
          Presentations, <span className="text-primary">spreadsheets</span> and documents.
        </h1>
        <p className="mt-5 max-w-xl text-lg text-muted-foreground">
          Create, calculate, write and present in one workspace. Each format has its own AI assistant and Microsoft Office export.
        </p>
        <div className="mt-8 max-w-2xl rounded-2xl border border-border bg-card p-4">
          <div className="mb-2 flex items-center gap-2 text-sm font-semibold"><Wand2 className="h-4 w-4 text-primary" /> Build a whole deck with AI</div>
          <p className="mb-3 text-xs text-muted-foreground">AI runs in your browser. The model downloads once, then stays cached on this device.</p>
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
              {building ? <Loader2 className="animate-spin" /> : <Sparkles />} {building ? (aiStatus ? localAiStatusLabel(aiStatus) : "Preparing local AI…") : "Build it"}
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

      <section className="mx-auto max-w-7xl px-8 pb-12">
        <div className="mb-5 flex items-center justify-between gap-4">
          <h2 className="text-sm font-semibold uppercase tracking-widest text-muted-foreground">Your spreadsheets</h2>
          <Button size="sm" variant="ghost" onClick={createWorkbook}><Plus /> New spreadsheet</Button>
        </div>
        {workbooks.length === 0 ? (
          <p className="text-sm text-muted-foreground">Your workbooks will appear here.</p>
        ) : (
          <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
            {workbooks.map((workbook) => (
              <div key={workbook.id} className="flex items-center justify-between gap-3 border-b border-border py-3">
                <Link to="/spreadsheet/$workbookId" params={{ workbookId: workbook.id }} className="flex min-w-0 items-center gap-3 hover:text-primary">
                  <FileSpreadsheet className="h-5 w-5 shrink-0 text-emerald-500" />
                  <span className="min-w-0"><span className="block truncate font-semibold">{workbook.title}</span><span className="block text-xs text-muted-foreground">{workbook.sheets.length} sheets · {new Date(workbook.updatedAt).toLocaleDateString()}</span></span>
                </Link>
                <Button size="icon" variant="ghost" aria-label={`Delete ${workbook.title}`} onClick={() => { if (confirm(`Delete "${workbook.title}"?`)) { deleteWorkbook(workbook.id); refreshLibrary(); } }}><Trash2 /></Button>
              </div>
            ))}
          </div>
        )}
      </section>

      <section className="mx-auto max-w-7xl px-8 pb-24">
        <div className="mb-5 flex items-center justify-between gap-4">
          <h2 className="text-sm font-semibold uppercase tracking-widest text-muted-foreground">Your Word documents</h2>
          <Button size="sm" variant="ghost" onClick={createDocument}><Plus /> New document</Button>
        </div>
        {documents.length === 0 ? (
          <p className="text-sm text-muted-foreground">Your documents will appear here.</p>
        ) : (
          <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
            {documents.map((document) => (
              <div key={document.id} className="flex items-center justify-between gap-3 border-b border-border py-3">
                <Link to="/document/$documentId" params={{ documentId: document.id }} className="flex min-w-0 items-center gap-3 hover:text-primary">
                  <FileText className="h-5 w-5 shrink-0 text-sky-400" />
                  <span className="min-w-0"><span className="block truncate font-semibold">{document.title}</span><span className="block text-xs text-muted-foreground">Word document · {new Date(document.updatedAt).toLocaleDateString()}</span></span>
                </Link>
                <Button size="icon" variant="ghost" aria-label={`Delete ${document.title}`} onClick={() => { if (confirm(`Delete "${document.title}"?`)) { deleteDocument(document.id); refreshLibrary(); } }}><Trash2 /></Button>
              </div>
            ))}
          </div>
        )}
      </section>
    </div>
  );
}
