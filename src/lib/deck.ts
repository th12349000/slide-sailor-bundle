export const SLIDE_W = 1920;
export const SLIDE_H = 1080;

export type ShapeKind = "rect" | "ellipse" | "triangle" | "star" | "arrow" | "line" | "diamond";

export type SlideLink =
  | { kind: "url"; url: string }
  | { kind: "slide"; slideIndex: number }
  | { kind: "deck"; deckId: string; slideIndex: number };

export interface SlideElement {
  id: string;
  type: "text" | "shape" | "image";
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number;
  opacity: number;
  text: string;
  fontSize: number;
  fontFamily: "display" | "sans" | "serif" | "mono";
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  align: "left" | "center" | "right";
  fill: string;
  borderColor: string;
  borderWidth: number;
  radius: number;
  shape: ShapeKind;
  src: string;
  link: SlideLink | null;
}

export type TransitionKind = "none" | "fade" | "slide" | "zoom" | "wipe";

export interface Slide {
  id: string;
  background: string;
  notes: string;
  transition: TransitionKind;
  elements: SlideElement[];
}

export interface Deck {
  id: string;
  title: string;
  updatedAt: number;
  slides: Slide[];
}

export const uid = () => Math.random().toString(36).slice(2, 10);

export function makeElement(type: SlideElement["type"], patch: Partial<SlideElement> = {}): SlideElement {
  const base: SlideElement = {
    id: uid(),
    type,
    x: 660,
    y: 390,
    w: 600,
    h: 300,
    rotation: 0,
    opacity: 1,
    text: type === "text" ? "Double-click to edit" : "",
    fontSize: 56,
    fontFamily: "sans",
    color: "#1b1b1f",
    bold: false,
    italic: false,
    underline: false,
    align: "left",
    fill: type === "shape" ? "#c6f432" : "transparent",
    borderColor: "#1b1b1f",
    borderWidth: 0,
    radius: 0,
    shape: "rect",
    src: "",
    link: null,
  };
  if (type === "text") {
    base.w = 900;
    base.h = 160;
  }
  return { ...base, ...patch };
}

export function makeSlide(patch: Partial<Slide> = {}): Slide {
  return { id: uid(), background: "#f6f3ec", notes: "", transition: "fade", elements: [], ...patch };
}

// ---- AI chat history (per deck, per browser) ----
export interface AiChatMsg {
  role: "user" | "ai";
  text: string;
}

const aiKey = (deckId: string) => `slideforge.ai.${deckId}`;

export function loadAiChat(deckId: string): AiChatMsg[] {
  if (typeof window === "undefined") return [];
  try {
    const v = JSON.parse(localStorage.getItem(aiKey(deckId)) || "[]");
    return Array.isArray(v) ? v : [];
  } catch {
    return [];
  }
}

export function saveAiChat(deckId: string, msgs: AiChatMsg[]) {
  try {
    localStorage.setItem(aiKey(deckId), JSON.stringify(msgs.slice(-100)));
  } catch {
    /* storage full — history is best-effort */
  }
}

export function clearAiChat(deckId: string) {
  try {
    localStorage.removeItem(aiKey(deckId));
  } catch {
    /* ignore */
  }
}

export function makeDeck(title = "Untitled presentation"): Deck {
  return {
    id: uid(),
    title,
    updatedAt: Date.now(),
    slides: [
      makeSlide({
        elements: [
          makeElement("text", {
            text: title,
            x: 160,
            y: 380,
            w: 1600,
            h: 200,
            fontSize: 120,
            fontFamily: "display",
            bold: true,
          }),
          makeElement("text", { text: "Start building — or ask the AI", x: 160, y: 600, w: 1400, h: 100, fontSize: 44, color: "#5a5a63" }),
        ],
      }),
    ],
  };
}

/** Decide where a link goes. Returns null when it points nowhere valid. */
export function resolveLink(
  link: SlideLink,
  currentDeckId: string,
  decks: Pick<Deck, "id" | "slides">[],
): { type: "external"; url: string } | { type: "internal"; deckId: string; slideIndex: number } | null {
  if (link.kind === "url") {
    const url = link.url.trim();
    if (!url) return null;
    return { type: "external", url: /^(https?:|mailto:)/i.test(url) ? url : `https://${url}` };
  }
  const deckId = link.kind === "deck" ? link.deckId : currentDeckId;
  const deck = decks.find((d) => d.id === deckId);
  if (!deck || deck.slides.length === 0) return null;
  const slideIndex = Math.min(Math.max(0, link.slideIndex), deck.slides.length - 1);
  return { type: "internal", deckId, slideIndex };
}

const KEY = "slideforge.decks.v1";

export function loadDecks(): Deck[] {
  if (typeof window === "undefined") return [];
  try {
    return JSON.parse(localStorage.getItem(KEY) || "[]");
  } catch {
    return [];
  }
}

export function saveDecks(decks: Deck[]) {
  try {
    localStorage.setItem(KEY, JSON.stringify(decks));
  } catch (e) {
    console.error("Could not save — storage may be full (large images).", e);
    throw e;
  }
}

export function getDeck(id: string) {
  return loadDecks().find((d) => d.id === id);
}

export function upsertDeck(deck: Deck) {
  const decks = loadDecks();
  const i = decks.findIndex((d) => d.id === deck.id);
  const next = { ...deck, updatedAt: Date.now() };
  if (i >= 0) decks[i] = next;
  else decks.unshift(next);
  saveDecks(decks);
}

export function deleteDeck(id: string) {
  saveDecks(loadDecks().filter((d) => d.id !== id));
}

export const FONT_STACK: Record<SlideElement["fontFamily"], string> = {
  display: "var(--font-display)",
  sans: "var(--font-sans)",
  serif: "'Fraunces', Georgia, serif",
  mono: "'JetBrains Mono', monospace",
};
