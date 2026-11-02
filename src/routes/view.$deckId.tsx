import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useCallback, useEffect, useRef, useState } from "react";
import { ChevronLeft, ChevronRight, Maximize, Pencil, X } from "lucide-react";
import { z } from "zod";
import { ScaledSlide, SlideStatic } from "@/components/slides/SlideRender";
import { loadDecks, resolveLink, type Deck, type SlideElement } from "@/lib/deck";

export const Route = createFileRoute("/view/$deckId")({
  validateSearch: z.object({ slide: z.number().int().min(1).catch(1) }),
  head: () => ({
    meta: [
      { title: "Presenting — SlideForge" },
      { name: "description", content: "Present your SlideForge deck full screen." },
      { property: "og:title", content: "Presenting — SlideForge" },
      { property: "og:description", content: "Present your SlideForge deck full screen." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Viewer,
});

function Viewer() {
  const { deckId } = Route.useParams();
  const { slide } = Route.useSearch();
  const navigate = useNavigate();
  const [decks, setDecks] = useState<Deck[] | null>(null);
  const [chrome, setChrome] = useState(true);
  const timer = useRef<number>(0);
  const touchX = useRef<number | null>(null);

  useEffect(() => setDecks(loadDecks()), [deckId]);
  const deck = decks?.find((d) => d.id === deckId);
  const current = deck?.slides[Math.min(slide - 1, Math.max(0, (deck?.slides.length ?? 1) - 1))];
  const total = deck?.slides.length ?? 0;
  const index = Math.min(slide - 1, Math.max(0, total - 1));

  const dir = useRef<1 | -1>(1);
  const go = useCallback(
    (i: number, id = deckId) => {
      dir.current = i >= index ? 1 : -1;
      navigate({ to: "/view/$deckId", params: { deckId: id }, search: { slide: i + 1 }, replace: id === deckId });
    },
    [navigate, deckId, index],
  );
  const next = () => index < total - 1 && go(index + 1);
  const prev = () => index > 0 && go(index - 1);

  const transitionClass = (t: string | undefined) => {
    switch (t ?? "fade") {
      case "fade": return "sf-transition sf-fade";
      case "slide": return `sf-transition ${dir.current === 1 ? "sf-slide-next" : "sf-slide-prev"}`;
      case "zoom": return "sf-transition sf-zoom";
      case "wipe": return `sf-transition ${dir.current === 1 ? "sf-wipe-next" : "sf-wipe-prev"}`;
      default: return "";
    }
  };

  useEffect(() => {
    if (deck) document.title = `${index + 1}/${total} — ${deck.title}`;
  }, [deck, index, total]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (["ArrowRight", "ArrowDown", " ", "PageDown"].includes(e.key)) { e.preventDefault(); next(); }
      if (["ArrowLeft", "ArrowUp", "PageUp"].includes(e.key)) { e.preventDefault(); prev(); }
      if (e.key === "Home") go(0);
      if (e.key === "End") go(total - 1);
      if (e.key === "Escape" && !document.fullscreenElement) navigate({ to: "/edit/$deckId", params: { deckId } });
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  const poke = () => {
    setChrome(true);
    window.clearTimeout(timer.current);
    timer.current = window.setTimeout(() => setChrome(false), 2500);
  };
  useEffect(() => { poke(); return () => window.clearTimeout(timer.current); }, []);

  const onLink = (el: SlideElement) => {
    if (!el.link || !decks) return;
    const r = resolveLink(el.link, deckId, decks);
    if (!r) return;
    if (r.type === "external") window.open(r.url, "_blank", "noopener,noreferrer");
    else go(r.slideIndex, r.deckId);
  };

  if (decks === null) return <div className="min-h-screen bg-canvas" />;
  if (!deck) return (
    <div className="grid min-h-screen place-items-center bg-canvas text-center">
      <div><p className="mb-4 text-muted-foreground">This presentation doesn't exist on this device.</p><Link className="text-primary underline" to="/">Back to presentations</Link></div>
    </div>
  );

  return (
    <div
      className="fixed inset-0 bg-canvas"
      style={{ cursor: chrome ? "default" : "none" }}
      onMouseMove={poke}
      onTouchStart={(e) => (touchX.current = e.touches[0]!.clientX)}
      onTouchEnd={(e) => {
        if (touchX.current === null) return;
        const dx = e.changedTouches[0]!.clientX - touchX.current;
        if (dx < -50) next(); else if (dx > 50) prev();
        touchX.current = null;
      }}
    >
      <ScaledSlide className="h-full w-full">
        {current && (
          <div key={`${deck.id}-${index}`} className={`h-full w-full ${transitionClass(current.transition)}`}>
            <SlideStatic slide={current} onLink={onLink} />
          </div>
        )}
      </ScaledSlide>
      <div className={`absolute bottom-5 left-1/2 flex -translate-x-1/2 items-center gap-1 rounded-full border border-border bg-popover/90 px-2 py-1 backdrop-blur transition-opacity ${chrome ? "opacity-100" : "pointer-events-none opacity-0"}`}>
        <IconBtn onClick={prev} label="Previous"><ChevronLeft /></IconBtn>
        <span className="px-2 text-sm tabular-nums">{index + 1} / {total}</span>
        <IconBtn onClick={next} label="Next"><ChevronRight /></IconBtn>
        <IconBtn label="Fullscreen" onClick={() => document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen()}><Maximize className="h-4 w-4" /></IconBtn>
        <Link to="/edit/$deckId" params={{ deckId }} className="grid h-9 w-9 place-items-center rounded-full hover:bg-accent" aria-label="Edit"><Pencil className="h-4 w-4" /></Link>
        <Link to="/" className="grid h-9 w-9 place-items-center rounded-full hover:bg-accent" aria-label="Close"><X className="h-4 w-4" /></Link>
      </div>
    </div>
  );
}

function IconBtn({ children, onClick, label }: { children: React.ReactNode; onClick: () => void; label: string }) {
  return <button aria-label={label} onClick={onClick} className="grid h-9 w-9 place-items-center rounded-full hover:bg-accent">{children}</button>;
}
