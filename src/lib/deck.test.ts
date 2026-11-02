import { describe, expect, it } from "vitest";
import { resolveLink, makeSlide } from "./deck";

const decks = [
  { id: "a", slides: [makeSlide(), makeSlide(), makeSlide()] },
  { id: "b", slides: [makeSlide(), makeSlide()] },
];

describe("resolveLink", () => {
  it("url links add https when missing", () => {
    expect(resolveLink({ kind: "url", url: "example.com" }, "a", decks)).toEqual({ type: "external", url: "https://example.com" });
  });
  it("slide links stay in the current presentation", () => {
    expect(resolveLink({ kind: "slide", slideIndex: 2 }, "a", decks)).toEqual({ type: "internal", deckId: "a", slideIndex: 2 });
  });
  it("links to another presentation go to that presentation's slide", () => {
    expect(resolveLink({ kind: "deck", deckId: "b", slideIndex: 1 }, "a", decks)).toEqual({ type: "internal", deckId: "b", slideIndex: 1 });
  });
  it("links to a deleted presentation go nowhere", () => {
    expect(resolveLink({ kind: "deck", deckId: "zzz", slideIndex: 0 }, "a", decks)).toBeNull();
  });
});
