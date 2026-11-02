import { Link, useNavigate } from "@tanstack/react-router";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  ArrowDown, ArrowLeft, ArrowUp, BringToFront, Circle, Copy, Diamond, Download, Image as ImageIcon, Link2, Loader2, Minus,
  MoveRight, Play, Plus, Redo2, SendToBack, Send, Sparkles, Square, Star, Trash2, Triangle, Type, Undo2, Bold, Italic,
  Underline, AlignLeft, AlignCenter, AlignRight,
} from "lucide-react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { ElementBody, elementBoxStyle, ScaledSlide, SlideStatic, useSlideScale } from "./SlideRender";
import {
  FONT_STACK, SLIDE_H, SLIDE_W, clearAiChat, loadAiChat, loadDecks, makeElement, makeSlide, saveAiChat, uid, upsertDeck,
  type AiChatMsg, type Deck, type ShapeKind, type Slide, type SlideElement, type SlideLink, type TransitionKind,
} from "@/lib/deck";
import type { AiResult } from "@/lib/ai.server";
import { exportDeck, type ExportFormat } from "@/lib/pptx";

const SHAPES: { kind: ShapeKind; icon: typeof Square; label: string }[] = [
  { kind: "rect", icon: Square, label: "Rectangle" },
  { kind: "ellipse", icon: Circle, label: "Ellipse" },
  { kind: "triangle", icon: Triangle, label: "Triangle" },
  { kind: "diamond", icon: Diamond, label: "Diamond" },
  { kind: "star", icon: Star, label: "Star" },
  { kind: "arrow", icon: MoveRight, label: "Arrow" },
  { kind: "line", icon: Minus, label: "Line" },
];

type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

export function Editor({ initial }: { initial: Deck }) {
  const navigate = useNavigate();
  const [deck, setDeckRaw] = useState<Deck>(initial);
  const [cur, setCur] = useState(0);
  const [sel, setSel] = useState<string | null>(null);
  const [editing, setEditing] = useState<string | null>(null);
  const past = useRef<Deck[]>([]);
  const future = useRef<Deck[]>([]);
  const clipboard = useRef<SlideElement | null>(null);

  const setDeck = useCallback((fn: (d: Deck) => Deck, record = true) => {
    setDeckRaw((d) => {
      const n = fn(d);
      if (record && n !== d) {
        past.current.push(d);
        if (past.current.length > 100) past.current.shift();
        future.current = [];
      }
      return n;
    });
  }, []);

  const undo = () => { const p = past.current.pop(); if (p) { future.current.push(deck); setDeckRaw(p); } };
  const redo = () => { const f = future.current.pop(); if (f) { past.current.push(deck); setDeckRaw(f); } };

  // autosave
  useEffect(() => {
    const t = setTimeout(() => {
      try { upsertDeck(deck); } catch { toast.error("Couldn't save — try smaller images."); }
    }, 400);
    return () => clearTimeout(t);
  }, [deck]);

  const slide = deck.slides[Math.min(cur, deck.slides.length - 1)]!;
  const selected = slide.elements.find((e) => e.id === sel) ?? null;

  const updateSlide = (patch: Partial<Slide> | ((s: Slide) => Slide), record = true) =>
    setDeck((d) => ({
      ...d,
      slides: d.slides.map((s, i) => (i === cur ? (typeof patch === "function" ? patch(s) : { ...s, ...patch }) : s)),
    }), record);

  const updateEl = (id: string, patch: Partial<SlideElement>, record = true) =>
    updateSlide((s) => ({ ...s, elements: s.elements.map((e) => (e.id === id ? { ...e, ...patch } : e)) }), record);

  const addEl = (el: SlideElement) => { updateSlide((s) => ({ ...s, elements: [...s.elements, el] })); setSel(el.id); };
  const removeEl = (id: string) => { updateSlide((s) => ({ ...s, elements: s.elements.filter((e) => e.id !== id) })); setSel(null); };
  const duplicateEl = (el: SlideElement) => addEl({ ...structuredClone(el), id: uid(), x: el.x + 40, y: el.y + 40 });
  const reorder = (id: string, dir: "front" | "back" | "up" | "down") =>
    updateSlide((s) => {
      const els = [...s.elements];
      const i = els.findIndex((e) => e.id === id);
      const el = els.splice(i, 1)[0]!;
      const j = dir === "front" ? els.length : dir === "back" ? 0 : dir === "up" ? Math.min(els.length, i + 1) : Math.max(0, i - 1);
      els.splice(j, 0, el);
      return { ...s, elements: els };
    });

  // slides ops
  const addSlide = (after = cur) => {
    setDeck((d) => { const s = [...d.slides]; s.splice(after + 1, 0, makeSlide({ background: slide.background })); return { ...d, slides: s }; });
    setCur(after + 1); setSel(null);
  };
  const dupSlide = (i: number) => {
    setDeck((d) => { const s = [...d.slides]; const c = structuredClone(s[i]!); c.id = uid(); c.elements.forEach((e) => (e.id = uid())); s.splice(i + 1, 0, c); return { ...d, slides: s }; });
    setCur(i + 1);
  };
  const delSlide = (i: number) => {
    if (deck.slides.length === 1) return;
    setDeck((d) => ({ ...d, slides: d.slides.filter((_, j) => j !== i) }));
    setCur((c) => Math.max(0, c >= i ? c - 1 : c)); setSel(null);
  };
  const moveSlide = (i: number, dir: -1 | 1) => {
    const j = i + dir;
    if (j < 0 || j >= deck.slides.length) return;
    setDeck((d) => { const s = [...d.slides]; [s[i], s[j]] = [s[j]!, s[i]!]; return { ...d, slides: s }; });
    setCur(j);
  };

  // keyboard
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const t = e.target as HTMLElement;
      if (t.isContentEditable || ["INPUT", "TEXTAREA", "SELECT"].includes(t.tagName)) return;
      const mod = e.metaKey || e.ctrlKey;
      if (mod && e.key.toLowerCase() === "z") { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
      if (mod && e.key.toLowerCase() === "y") { e.preventDefault(); redo(); return; }
      if (mod && e.key.toLowerCase() === "c" && selected) { clipboard.current = structuredClone(selected); return; }
      if (mod && e.key.toLowerCase() === "v" && clipboard.current) { e.preventDefault(); duplicateEl(clipboard.current); return; }
      if (mod && e.key.toLowerCase() === "d" && selected) { e.preventDefault(); duplicateEl(selected); return; }
      if (e.key === "F5") { e.preventDefault(); present(); return; }
      if (selected) {
        if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); removeEl(selected.id); }
        const step = e.shiftKey ? 50 : 5;
        const mv: Record<string, [number, number]> = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] };
        if (mv[e.key]) { e.preventDefault(); updateEl(selected.id, { x: selected.x + mv[e.key]![0], y: selected.y + mv[e.key]![1] }); }
        if (e.key === "Escape") setSel(null);
      } else {
        if (e.key === "ArrowDown" || e.key === "PageDown") setCur((c) => Math.min(deck.slides.length - 1, c + 1));
        if (e.key === "ArrowUp" || e.key === "PageUp") setCur((c) => Math.max(0, c - 1));
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  const present = () => { upsertDeck(deck); navigate({ to: "/view/$deckId", params: { deckId: deck.id }, search: { slide: cur + 1 } }); };

  const [exporting, setExporting] = useState(false);
  const doExport = async (format: ExportFormat) => {
    setExporting(true);
    try {
      upsertDeck(deck);
      const name = await exportDeck(deck, format);
      toast.success(`Downloaded ${name}`);
    } catch (e) {
      toast.error(`Export failed: ${(e as Error).message}`);
    } finally {
      setExporting(false);
    }
  };

  const fileRef = useRef<HTMLInputElement>(null);
  const onImageFile = (f: File | undefined) => {
    if (!f) return;
    const img = new window.Image();
    const reader = new FileReader();
    reader.onload = () => {
      img.onload = () => {
        // downscale to keep storage small
        const max = 1600;
        const r = Math.min(1, max / Math.max(img.width, img.height));
        const c = document.createElement("canvas");
        c.width = img.width * r; c.height = img.height * r;
        c.getContext("2d")!.drawImage(img, 0, 0, c.width, c.height);
        const src = c.toDataURL("image/jpeg", 0.85);
        const w = Math.min(900, c.width), h = (w / c.width) * c.height;
        addEl(makeElement("image", { src, w, h, x: (SLIDE_W - w) / 2, y: (SLIDE_H - h) / 2 }));
      };
      img.src = reader.result as string;
    };
    reader.readAsDataURL(f);
  };

  return (
    <div className="flex h-screen flex-col bg-background">
      {/* Toolbar */}
      <header className="flex h-14 shrink-0 items-center gap-2 border-b border-border px-3">
        <Button variant="ghost" size="icon" asChild><Link to="/" aria-label="Back"><ArrowLeft /></Link></Button>
        <Input value={deck.title} onChange={(e) => setDeck((d) => ({ ...d, title: e.target.value }), false)} className="h-9 w-64 border-transparent bg-transparent font-display font-bold hover:border-border" />
        <div className="mx-2 h-6 w-px bg-border" />
        <Button variant="ghost" size="sm" onClick={() => addEl(makeElement("text"))}><Type /> Text</Button>
        <DropdownMenu>
          <DropdownMenuTrigger asChild><Button variant="ghost" size="sm"><Square /> Shape</Button></DropdownMenuTrigger>
          <DropdownMenuContent>
            {SHAPES.map((s) => (
              <DropdownMenuItem key={s.kind} onClick={() => addEl(makeElement("shape", { shape: s.kind, ...(s.kind === "line" ? { h: 40, fill: "#1b1b1f" } : s.kind === "rect" || s.kind === "ellipse" ? {} : { w: 400, h: 400 }) }))}>
                <s.icon /> {s.label}
              </DropdownMenuItem>
            ))}
          </DropdownMenuContent>
        </DropdownMenu>
        <DropdownMenu>
          <DropdownMenuTrigger asChild><Button variant="ghost" size="sm"><ImageIcon /> Image</Button></DropdownMenuTrigger>
          <DropdownMenuContent>
            <DropdownMenuItem onClick={() => fileRef.current?.click()}>Upload from device</DropdownMenuItem>
            <DropdownMenuItem onClick={() => { const u = prompt("Image URL"); if (u) addEl(makeElement("image", { src: u, w: 800, h: 500, x: 560, y: 290 })); }}>From URL…</DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <input ref={fileRef} type="file" accept="image/*" hidden onChange={(e) => { onImageFile(e.target.files?.[0]); e.target.value = ""; }} />
        <div className="mx-2 h-6 w-px bg-border" />
        <Button variant="ghost" size="icon" onClick={undo} aria-label="Undo"><Undo2 /></Button>
        <Button variant="ghost" size="icon" onClick={redo} aria-label="Redo"><Redo2 /></Button>
        <div className="flex-1" />
        <DropdownMenu>
          <DropdownMenuTrigger asChild><Button variant="outline" disabled={exporting}>{exporting ? <Loader2 className="animate-spin" /> : <Download />} Export</Button></DropdownMenuTrigger>
          <DropdownMenuContent>
            <DropdownMenuItem onClick={() => doExport("pptx")}>PowerPoint (.pptx)</DropdownMenuItem>
            <DropdownMenuItem onClick={() => doExport("ppsx")}>PowerPoint Slideshow (.ppsx)</DropdownMenuItem>
            <DropdownMenuItem onClick={() => doExport("both")}>Both formats</DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
        <Button onClick={present}><Play /> Present</Button>
      </header>

      <div className="flex min-h-0 flex-1">
        {/* Slide list */}
        <aside className="flex w-52 shrink-0 flex-col border-r border-border bg-sidebar">
          <div className="flex-1 space-y-3 overflow-y-auto p-3">
            {deck.slides.map((s, i) => (
              <div key={s.id} className="group relative flex gap-2">
                <span className="w-4 pt-1 text-xs text-muted-foreground tabular-nums">{i + 1}</span>
                <button onClick={() => { setCur(i); setSel(null); }} className={`flex-1 overflow-hidden rounded-md border-2 ${i === cur ? "border-primary" : "border-transparent hover:border-border"}`}>
                  <ScaledSlide className="aspect-video"><SlideStatic slide={s} /></ScaledSlide>
                </button>
                <div className="absolute right-1 top-1 hidden flex-col gap-0.5 group-hover:flex">
                  {[
                    { i: ArrowUp, f: () => moveSlide(i, -1), l: "Move up" },
                    { i: ArrowDown, f: () => moveSlide(i, 1), l: "Move down" },
                    { i: Copy, f: () => dupSlide(i), l: "Duplicate" },
                    { i: Trash2, f: () => delSlide(i), l: "Delete" },
                  ].map((b) => (
                    <button key={b.l} aria-label={b.l} title={b.l} onClick={b.f} className="grid h-5 w-5 place-items-center rounded bg-popover/90 hover:bg-accent"><b.i className="h-3 w-3" /></button>
                  ))}
                </div>
              </div>
            ))}
          </div>
          <div className="border-t border-border p-3"><Button variant="secondary" className="w-full" onClick={() => addSlide()}><Plus /> New slide</Button></div>
        </aside>

        {/* Canvas */}
        <main className="flex min-w-0 flex-1 flex-col bg-canvas grid-dots">
          <div className="min-h-0 flex-1 p-8" onPointerDown={() => { setSel(null); setEditing(null); }}>
            <ScaledSlide className="h-full w-full">
              <div className="slide-shadow" style={{ width: SLIDE_W, height: SLIDE_H, background: slide.background, position: "relative", overflow: "hidden" }}>
                {slide.elements.map((el) => (
                  <EditableElement
                    key={el.id}
                    el={el}
                    selected={el.id === sel}
                    editing={el.id === editing}
                    onSelect={() => { setSel(el.id); if (editing !== el.id) setEditing(null); }}
                    onEdit={() => el.type !== "image" && setEditing(el.id)}
                    onChange={(p, rec) => updateEl(el.id, p, rec)}
                    onCommitStart={() => { past.current.push(deck); future.current = []; }}
                  />
                ))}
              </div>
            </ScaledSlide>
          </div>
          <div className="h-24 shrink-0 border-t border-border bg-card px-4 py-2">
            <Textarea placeholder="Speaker notes…" value={slide.notes} onChange={(e) => updateSlide({ notes: e.target.value }, false)} className="h-full resize-none border-0 bg-transparent focus-visible:ring-0" />
          </div>
        </main>

        {/* Right panel */}
        <aside className="w-80 shrink-0 border-l border-border bg-sidebar">
          <Tabs defaultValue="design" className="flex h-full flex-col">
            <TabsList className="m-3 grid grid-cols-2">
              <TabsTrigger value="design">Design</TabsTrigger>
              <TabsTrigger value="ai"><Sparkles className="h-4 w-4" /> AI</TabsTrigger>
            </TabsList>
            <TabsContent value="design" className="min-h-0 flex-1 overflow-y-auto px-4 pb-6">
              {selected ? (
                <ElementPanel
                  el={selected}
                  deck={deck}
                  onChange={(p) => updateEl(selected.id, p)}
                  onDelete={() => removeEl(selected.id)}
                  onDuplicate={() => duplicateEl(selected)}
                  onOrder={(d) => reorder(selected.id, d)}
                />
              ) : (
                <SlidePanel slide={slide} onChange={(p) => updateSlide(p)} />
              )}
            </TabsContent>
            <TabsContent value="ai" className="min-h-0 flex-1">
              <AiPanel
                deck={deck}
                cur={cur}
                apply={(fn, newCur) => { setDeck(fn); if (newCur !== undefined) setCur(newCur); setSel(null); }}
              />
            </TabsContent>
          </Tabs>
        </aside>
      </div>
    </div>
  );
}

function EditableElement({
  el, selected, editing, onSelect, onEdit, onChange, onCommitStart,
}: {
  el: SlideElement; selected: boolean; editing: boolean;
  onSelect: () => void; onEdit: () => void;
  onChange: (p: Partial<SlideElement>, record: boolean) => void;
  onCommitStart: () => void;
}) {
  const scale = useSlideScale();
  const textRef = useRef<HTMLDivElement>(null);

  const startDrag = (e: React.PointerEvent, handle?: Handle | "rotate") => {
    e.stopPropagation();
    if (editing && !handle) return;
    onSelect();
    const sx = e.clientX, sy = e.clientY;
    const start = { ...el };
    let moved = false;
    const move = (ev: PointerEvent) => {
      const dx = (ev.clientX - sx) / scale, dy = (ev.clientY - sy) / scale;
      if (!moved) { if (Math.abs(dx) + Math.abs(dy) < 2) return; moved = true; onCommitStart(); }
      if (!handle) {
        let x = start.x + dx, y = start.y + dy;
        // snap to slide center / edges
        const snap = (v: number, size: number, total: number) => {
          for (const t of [0, (total - size) / 2, total - size]) if (Math.abs(v - t) < 12) return t;
          return v;
        };
        x = snap(x, start.w, SLIDE_W); y = snap(y, start.h, SLIDE_H);
        onChange({ x: Math.round(x), y: Math.round(y) }, false);
      } else if (handle === "rotate") {
        const cx = start.x + start.w / 2, cy = start.y + start.h / 2;
        const rect = (e.target as HTMLElement).closest("[data-slide-root]")?.getBoundingClientRect();
        const ox = rect ? rect.left + cx * scale : sx, oy = rect ? rect.top + cy * scale : sy;
        let deg = (Math.atan2(ev.clientY - oy, ev.clientX - ox) * 180) / Math.PI + 90;
        if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
        onChange({ rotation: Math.round(deg) }, false);
      } else {
        let { x, y, w, h } = start;
        if (handle.includes("e")) w = Math.max(20, start.w + dx);
        if (handle.includes("s")) h = Math.max(20, start.h + dy);
        if (handle.includes("w")) { w = Math.max(20, start.w - dx); x = start.x + start.w - w; }
        if (handle.includes("n")) { h = Math.max(20, start.h - dy); y = start.y + start.h - h; }
        if (ev.shiftKey && handle.length === 2) { const r = start.w / start.h; h = w / r; if (handle.includes("n")) y = start.y + start.h - h; }
        onChange({ x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(h) }, false);
      }
    };
    const up = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up); };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  useEffect(() => {
    if (editing && textRef.current) {
      textRef.current.focus();
      const r = document.createRange(); r.selectNodeContents(textRef.current);
      const s = window.getSelection(); s?.removeAllRanges(); s?.addRange(r);
    }
  }, [editing]);

  const editor = editing ? (
    <div
      ref={textRef}
      contentEditable
      suppressContentEditableWarning
      onPointerDown={(e) => e.stopPropagation()}
      onBlur={(e) => onChange({ text: e.currentTarget.innerText }, true)}
      style={{
        outline: "none", fontSize: el.fontSize, fontFamily: FONT_STACK[el.fontFamily], color: el.color,
        fontWeight: el.bold ? 700 : 400, fontStyle: el.italic ? "italic" : "normal", textDecoration: el.underline ? "underline" : "none",
        textAlign: el.align, lineHeight: 1.15, whiteSpace: "pre-wrap", wordBreak: "break-word", minHeight: "1em", cursor: "text",
      }}
    >{el.text}</div>
  ) : undefined;

  const hs = 14 / scale;
  return (
    <div
      data-slide-root
      style={{ ...elementBoxStyle(el), cursor: editing ? "text" : "move", outline: selected ? `${2 / scale}px solid var(--selection)` : undefined, outlineOffset: 0 }}
      onPointerDown={(e) => startDrag(e)}
      onDoubleClick={(e) => { e.stopPropagation(); onEdit(); }}
    >
      <ElementBody el={el} editingText={editor} />
      {el.link && <div style={{ position: "absolute", top: -hs * 2.2, left: 0, fontSize: 12 / scale, padding: `${2 / scale}px ${6 / scale}px`, borderRadius: 4 / scale, background: "var(--selection)", color: "var(--primary-foreground)" }}>🔗 link</div>}
      {selected && !editing && (
        <>
          {HANDLES.map((h) => {
            const pos: React.CSSProperties = {
              left: h.includes("w") ? -hs / 2 : h.includes("e") ? undefined : `calc(50% - ${hs / 2}px)`,
              right: h.includes("e") ? -hs / 2 : undefined,
              top: h.includes("n") ? -hs / 2 : h.includes("s") ? undefined : `calc(50% - ${hs / 2}px)`,
              bottom: h.includes("s") ? -hs / 2 : undefined,
              cursor: `${h}-resize`,
            };
            return <div key={h} onPointerDown={(e) => startDrag(e, h)} style={{ position: "absolute", width: hs, height: hs, background: "var(--background)", border: `${2 / scale}px solid var(--selection)`, borderRadius: 2 / scale, ...pos }} />;
          })}
          <div onPointerDown={(e) => startDrag(e, "rotate")} title="Rotate (Shift to snap)" style={{ position: "absolute", left: `calc(50% - ${hs / 2}px)`, top: -hs * 3, width: hs, height: hs, borderRadius: "50%", background: "var(--selection)", cursor: "grab" }} />
        </>
      )}
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="block space-y-1.5"><span className="text-xs font-medium uppercase tracking-wider text-muted-foreground">{label}</span>{children}</label>;
}

function ColorInput({ value, onChange, allowNone }: { value: string; onChange: (v: string) => void; allowNone?: boolean }) {
  const isNone = value === "transparent";
  return (
    <div className="flex items-center gap-2">
      <input type="color" value={isNone ? "#ffffff" : value} onChange={(e) => onChange(e.target.value)} className="h-9 w-10 cursor-pointer rounded border border-input bg-transparent" />
      <Input value={value} onChange={(e) => onChange(e.target.value)} className="h-9 font-mono text-xs" />
      {allowNone && <Button size="sm" variant={isNone ? "default" : "outline"} onClick={() => onChange("transparent")}>None</Button>}
    </div>
  );
}

function Num({ value, onChange, min, max, step = 1 }: { value: number; onChange: (v: number) => void; min?: number; max?: number; step?: number }) {
  return <Input type="number" value={Math.round(value * 100) / 100} min={min} max={max} step={step} onChange={(e) => onChange(Number(e.target.value))} className="h-9" />;
}

const PRESET_BG = ["#f6f3ec", "#ffffff", "#111114", "#1d2b53", "#c6f432", "#ff6b4a", "#ffd23f", "#0f766e", "#e8e0ff"];

const TRANSITIONS: { kind: TransitionKind; label: string }[] = [
  { kind: "none", label: "None" },
  { kind: "fade", label: "Fade" },
  { kind: "slide", label: "Slide" },
  { kind: "zoom", label: "Zoom" },
  { kind: "wipe", label: "Wipe" },
];

function SlidePanel({ slide, onChange }: { slide: Slide; onChange: (p: Partial<Slide>) => void }) {
  return (
    <div className="space-y-5">
      <h3 className="text-lg font-bold">Slide</h3>
      <Field label="Background"><ColorInput value={slide.background} onChange={(background) => onChange({ background })} /></Field>
      <div className="flex flex-wrap gap-2">
        {PRESET_BG.map((c) => <button key={c} aria-label={c} onClick={() => onChange({ background: c })} className="h-8 w-8 rounded-full border border-border" style={{ background: c }} />)}
      </div>
      <Field label="Transition">
        <div className="flex flex-wrap gap-1.5">
          {TRANSITIONS.map((t) => (
            <Button key={t.kind} size="sm" variant={(slide.transition ?? "fade") === t.kind ? "default" : "outline"} onClick={() => onChange({ transition: t.kind })}>
              {t.label}
            </Button>
          ))}
        </div>
        <p className="text-xs text-muted-foreground">Plays when presenting, and is saved into exported PowerPoint files.</p>
      </Field>
      <p className="text-sm text-muted-foreground">Select something on the slide to edit it. Double-click text to type. Shift while resizing keeps proportions.</p>
      <div className="rounded-lg border border-border p-3 text-xs text-muted-foreground space-y-1">
        <div><b>Ctrl+Z / Ctrl+Y</b> undo / redo</div>
        <div><b>Ctrl+C / V / D</b> copy, paste, duplicate</div>
        <div><b>Arrows</b> nudge (Shift = 50px)</div>
        <div><b>Delete</b> remove · <b>F5</b> present</div>
      </div>
    </div>
  );
}

function ElementPanel({
  el, deck, onChange, onDelete, onDuplicate, onOrder,
}: {
  el: SlideElement; deck: Deck; onChange: (p: Partial<SlideElement>) => void;
  onDelete: () => void; onDuplicate: () => void; onOrder: (d: "front" | "back" | "up" | "down") => void;
}) {
  const hasText = el.type !== "image";
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <h3 className="text-lg font-bold capitalize">{el.type === "shape" ? el.shape : el.type}</h3>
        <div className="flex gap-1">
          <Button size="icon" variant="ghost" title="Bring to front" onClick={() => onOrder("front")}><BringToFront /></Button>
          <Button size="icon" variant="ghost" title="Send to back" onClick={() => onOrder("back")}><SendToBack /></Button>
          <Button size="icon" variant="ghost" title="Duplicate" onClick={onDuplicate}><Copy /></Button>
          <Button size="icon" variant="ghost" title="Delete" onClick={onDelete}><Trash2 /></Button>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <Field label="X"><Num value={el.x} onChange={(x) => onChange({ x })} /></Field>
        <Field label="Y"><Num value={el.y} onChange={(y) => onChange({ y })} /></Field>
        <Field label="Width"><Num value={el.w} onChange={(w) => onChange({ w })} /></Field>
        <Field label="Height"><Num value={el.h} onChange={(h) => onChange({ h })} /></Field>
        <Field label="Rotation°"><Num value={el.rotation} onChange={(rotation) => onChange({ rotation })} /></Field>
        <Field label="Opacity"><Num value={el.opacity} min={0} max={1} step={0.05} onChange={(opacity) => onChange({ opacity })} /></Field>
      </div>
      <div className="flex gap-2">
        <Button size="sm" variant="outline" className="flex-1" onClick={() => onChange({ x: (SLIDE_W - el.w) / 2 })}>Center H</Button>
        <Button size="sm" variant="outline" className="flex-1" onClick={() => onChange({ y: (SLIDE_H - el.h) / 2 })}>Center V</Button>
      </div>

      {hasText && (
        <div className="space-y-3">
          <Field label="Text"><Textarea value={el.text} onChange={(e) => onChange({ text: e.target.value })} rows={3} /></Field>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Size"><Num value={el.fontSize} min={8} onChange={(fontSize) => onChange({ fontSize })} /></Field>
            <Field label="Font">
              <select value={el.fontFamily} onChange={(e) => onChange({ fontFamily: e.target.value as SlideElement["fontFamily"] })} className="h-9 w-full rounded-md border border-input bg-transparent px-2 text-sm">
                <option value="display">Display</option><option value="sans">Sans</option><option value="serif">Serif</option><option value="mono">Mono</option>
              </select>
            </Field>
          </div>
          <div className="flex gap-1">
            <Toggle on={el.bold} onClick={() => onChange({ bold: !el.bold })}><Bold /></Toggle>
            <Toggle on={el.italic} onClick={() => onChange({ italic: !el.italic })}><Italic /></Toggle>
            <Toggle on={el.underline} onClick={() => onChange({ underline: !el.underline })}><Underline /></Toggle>
            <div className="w-2" />
            <Toggle on={el.align === "left"} onClick={() => onChange({ align: "left" })}><AlignLeft /></Toggle>
            <Toggle on={el.align === "center"} onClick={() => onChange({ align: "center" })}><AlignCenter /></Toggle>
            <Toggle on={el.align === "right"} onClick={() => onChange({ align: "right" })}><AlignRight /></Toggle>
          </div>
          <Field label="Text color"><ColorInput value={el.color} onChange={(color) => onChange({ color })} /></Field>
        </div>
      )}

      {el.type === "image" && (
        <Field label="Image URL"><Input value={el.src.startsWith("data:") ? "(uploaded)" : el.src} onChange={(e) => onChange({ src: e.target.value })} /></Field>
      )}
      {el.type !== "image" && <Field label="Fill"><ColorInput value={el.fill} onChange={(fill) => onChange({ fill })} allowNone /></Field>}
      {el.type === "shape" && (
        <Field label="Shape">
          <div className="flex flex-wrap gap-1">
            {SHAPES.map((s) => <Toggle key={s.kind} on={el.shape === s.kind} onClick={() => onChange({ shape: s.kind })}><s.icon /></Toggle>)}
          </div>
        </Field>
      )}
      <div className="grid grid-cols-2 gap-3">
        <Field label="Border"><Num value={el.borderWidth} min={0} onChange={(borderWidth) => onChange({ borderWidth })} /></Field>
        <Field label="Corner radius"><Num value={el.radius} min={0} onChange={(radius) => onChange({ radius })} /></Field>
      </div>
      {el.borderWidth > 0 && <Field label="Border color"><ColorInput value={el.borderColor} onChange={(borderColor) => onChange({ borderColor })} /></Field>}

      <LinkEditor link={el.link} deck={deck} onChange={(link) => onChange({ link })} />
    </div>
  );
}

function Toggle({ on, onClick, children }: { on: boolean; onClick: () => void; children: React.ReactNode }) {
  return <Button size="icon" variant={on ? "default" : "outline"} className="h-8 w-8" onClick={onClick}>{children}</Button>;
}

function LinkEditor({ link, deck, onChange }: { link: SlideLink | null; deck: Deck; onChange: (l: SlideLink | null) => void }) {
  const [decks, setDecks] = useState<Deck[]>([]);
  useEffect(() => setDecks(loadDecks().filter((d) => d.id !== deck.id)), [deck.id]);
  const kind = link?.kind ?? "none";
  const target = link?.kind === "deck" ? decks.find((d) => d.id === link.deckId) : undefined;
  return (
    <div className="space-y-3 rounded-lg border border-border p-3">
      <div className="flex items-center gap-2 text-sm font-semibold"><Link2 className="h-4 w-4 text-primary" /> Hyperlink</div>
      <select
        value={kind}
        onChange={(e) => {
          const k = e.target.value;
          onChange(k === "none" ? null : k === "url" ? { kind: "url", url: "" } : k === "slide" ? { kind: "slide", slideIndex: 0 } : { kind: "deck", deckId: decks[0]?.id ?? "", slideIndex: 0 });
        }}
        className="h-9 w-full rounded-md border border-input bg-transparent px-2 text-sm"
      >
        <option value="none">No link</option>
        <option value="url">Website URL</option>
        <option value="slide">Slide in this presentation</option>
        <option value="deck" disabled={decks.length === 0}>Slide in another presentation{decks.length === 0 ? " (none yet)" : ""}</option>
      </select>
      {link?.kind === "url" && <Input placeholder="https://example.com" value={link.url} onChange={(e) => onChange({ kind: "url", url: e.target.value })} />}
      {link?.kind === "slide" && (
        <select value={link.slideIndex} onChange={(e) => onChange({ kind: "slide", slideIndex: Number(e.target.value) })} className="h-9 w-full rounded-md border border-input bg-transparent px-2 text-sm">
          {deck.slides.map((s, i) => <option key={s.id} value={i}>Slide {i + 1}{slideLabel(s)}</option>)}
        </select>
      )}
      {link?.kind === "deck" && (
        <>
          <select value={link.deckId} onChange={(e) => onChange({ kind: "deck", deckId: e.target.value, slideIndex: 0 })} className="h-9 w-full rounded-md border border-input bg-transparent px-2 text-sm">
            {decks.map((d) => <option key={d.id} value={d.id}>{d.title}</option>)}
          </select>
          {target && (
            <select value={link.slideIndex} onChange={(e) => onChange({ ...link, slideIndex: Number(e.target.value) })} className="h-9 w-full rounded-md border border-input bg-transparent px-2 text-sm">
              {target.slides.map((s, i) => <option key={s.id} value={i}>Slide {i + 1}{slideLabel(s)}</option>)}
            </select>
          )}
        </>
      )}
      {link && <p className="text-xs text-muted-foreground">Links are clickable when presenting.</p>}
    </div>
  );
}

function slideLabel(s: Slide) {
  const t = s.elements.find((e) => e.text)?.text.split("\n")[0];
  return t ? ` — ${t.slice(0, 30)}` : "";
}

const GREETING: AiChatMsg = { role: "ai", text: "Hi! I'm your deck builder. Describe the presentation you want — e.g. “Build a 6-slide pitch about solar bikes” — or ask me to redesign, add or polish slides. I remember our chat, so you can build on earlier requests." };

function AiPanel({ deck, cur, apply }: { deck: Deck; cur: number; apply: (fn: (d: Deck) => Deck, cur?: number) => void }) {
  const [msgs, setMsgs] = useState<AiChatMsg[]>(() => {
    const saved = loadAiChat(deck.id);
    return saved.length ? saved : [GREETING];
  });
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<{ slides: number; elements: number; title: string; thinking: string; started: number }>({ slides: 0, elements: 0, title: "", thinking: "", started: 0 });
  const [now, setNow] = useState(0);
  useEffect(() => {
    if (!busy) return;
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [busy]);
  const endRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  useEffect(() => endRef.current?.scrollIntoView({ behavior: "smooth" }), [msgs]);
  useEffect(() => { saveAiChat(deck.id, msgs); }, [deck.id, msgs]);
  useEffect(() => { if (!busy) inputRef.current?.focus(); }, [busy]);

  const send = async (text = input) => {
    const prompt = text.trim();
    if (!prompt || busy) return;
    setInput("");
    setMsgs((m) => [...m, { role: "user", text: prompt }]);
    setBusy(true);
    setProgress({ slides: 0, elements: 0, title: "", thinking: "", started: Date.now() });
    setNow(Date.now());
    try {
      const stripped = {
        title: deck.title,
        slides: deck.slides.map((s) => ({ ...s, elements: s.elements.map((e) => ({ ...e, src: e.src.startsWith("data:") ? "(uploaded image)" : e.src })) })),
      };
      // include recent chat so follow-ups like "make it darker" have context
      const history = msgs.filter((m) => m !== GREETING).slice(-12).map((m) => `${m.role === "user" ? "User" : "Assistant"}: ${m.text}`).join("\n");
      const fullPrompt = history ? `Conversation so far:\n${history}\n\nNew request: ${prompt}` : prompt;
      const resp = await fetch("/api/slide-ai", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt: fullPrompt, deckJson: JSON.stringify(stripped), currentIndex: cur }),
      });
      if (!resp.ok || !resp.body) throw new Error(`The AI request failed (${resp.status}).`);
      const reader = resp.body.pipeThrough(new TextDecoderStream()).getReader();
      let buf = "";
      let r: AiResult | null = null;
      let failure: string | null = null;
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        buf += value;
        let nl: number;
        while ((nl = buf.indexOf("\n")) >= 0) {
          const line = buf.slice(0, nl).trim();
          buf = buf.slice(nl + 1);
          if (!line) continue;
          const ev = JSON.parse(line);
          if (ev.type === "reasoning") setProgress((p) => ({ ...p, thinking: (p.thinking + ev.text).slice(-600) }));
          else if (ev.type === "progress") setProgress((p) => ({ ...p, slides: ev.slides, elements: ev.elements, title: ev.currentTitle || p.title }));
          else if (ev.type === "done") r = ev.result;
          else if (ev.type === "error") failure = ev.error;
        }
      }
      if (failure || !r) { setMsgs((m) => [...m, { role: "ai", text: `⚠️ ${failure ?? "The AI stopped before finishing."}` }]); return; }
      const newSlides: Slide[] = r.slides.map((s) =>
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
      if (newSlides.length && r.mode !== "none") {
        if (r.mode === "replace_all") apply((d) => ({ ...d, slides: newSlides }), 0);
        else if (r.mode === "replace_current") apply((d) => { const s = [...d.slides]; s.splice(cur, 1, ...newSlides); return { ...d, slides: s }; }, cur);
        else apply((d) => { const s = [...d.slides]; s.splice(cur + 1, 0, ...newSlides); return { ...d, slides: s }; }, cur + 1);
      }
      setMsgs((m) => [...m, { role: "ai", text: r.reply + (newSlides.length && r.mode !== "none" ? "  (Ctrl+Z to undo)" : "") }]);
    } catch (e) {
      setMsgs((m) => [...m, { role: "ai", text: `⚠️ ${(e as Error).message || "Something went wrong."}` }]);
    } finally {
      setBusy(false);
    }
  };

  const ideas = ["Build a full deck about…", "Polish this slide's design", "Add a title + agenda slide", "Add a closing slide with a link back to slide 1"];
  return (
    <div className="flex h-full flex-col">
      <div className="flex items-center justify-between border-b border-border px-4 py-2">
        <span className="text-xs text-muted-foreground">{msgs.length > 1 ? `${msgs.length - 1} messages · saved` : "Chat is saved per presentation"}</span>
        {msgs.length > 1 && (
          <button onClick={() => { clearAiChat(deck.id); setMsgs([GREETING]); }} className="text-xs text-muted-foreground underline hover:text-foreground">Clear chat</button>
        )}
      </div>
      <div className="flex-1 space-y-3 overflow-y-auto px-4 py-3">
        {msgs.map((m, i) => (
          <div key={i} className={`rounded-xl px-3 py-2 text-sm whitespace-pre-wrap ${m.role === "user" ? "ml-6 bg-primary text-primary-foreground" : "mr-6 bg-secondary"}`}>{m.text}</div>
        ))}
        {busy && <AiProgressCard progress={progress} seconds={Math.max(0, Math.round((now - progress.started) / 1000))} />}
        {msgs.length === 1 && !busy && (
          <div className="flex flex-wrap gap-2 pt-2">
            {ideas.map((i) => <button key={i} onClick={() => i.endsWith("…") ? setInput("Build a full deck about ") : send(i)} className="rounded-full border border-border px-3 py-1 text-xs hover:border-primary hover:text-primary">{i}</button>)}
          </div>
        )}
        <div ref={endRef} />
      </div>
      <form className="flex gap-2 border-t border-border p-3" onSubmit={(e) => { e.preventDefault(); send(); }}>
        <Textarea ref={inputRef} value={input} onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }} placeholder="Describe what to build or change…" rows={2} className="resize-none" />
        <Button type="submit" size="icon" disabled={busy} className="self-end" aria-label="Send"><Send /></Button>
      </form>
    </div>
  );
}

function AiProgressCard({ progress, seconds }: { progress: { slides: number; elements: number; title: string; thinking: string }; seconds: number }) {
  const step = progress.slides > 0
    ? `Building slide ${progress.slides}${progress.title ? ` — “${progress.title}”` : ""}`
    : progress.thinking ? "Planning the design…" : "Starting up…";
  const thought = progress.thinking.replace(/\*\*/g, "").trim().split("\n").filter(Boolean).slice(-3).join("\n");
  return (
    <div className="mr-6 space-y-2 rounded-xl bg-secondary px-3 py-2 text-sm" aria-live="polite">
      <div className="flex items-center gap-2 font-medium">
        <Loader2 className="h-4 w-4 shrink-0 animate-spin" />
        <span className="flex-1 truncate">{step}</span>
        <span className="text-xs tabular-nums text-muted-foreground">{seconds}s</span>
      </div>
      {progress.slides > 0 && (
        <div className="space-y-1">
          <div className="h-1.5 overflow-hidden rounded-full bg-border">
            <div className="h-full rounded-full bg-primary transition-all duration-500" style={{ width: `${Math.min(95, 15 + progress.slides * 12)}%` }} />
          </div>
          <p className="text-xs text-muted-foreground">{progress.slides} slide{progress.slides === 1 ? "" : "s"} · {progress.elements} item{progress.elements === 1 ? "" : "s"} so far</p>
        </div>
      )}
      {thought && <p className="line-clamp-3 whitespace-pre-wrap text-xs italic text-muted-foreground">{thought}</p>}
    </div>
  );
}
