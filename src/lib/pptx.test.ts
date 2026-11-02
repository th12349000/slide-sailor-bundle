import { describe, expect, it } from "vitest";
import { sanitizeElement, toHex6 } from "@/lib/pptx";
import { makeElement } from "@/lib/deck";

describe("PowerPoint-safe values", () => {
  it("expands short hex colors to 6 digits", () => expect(toHex6("#fff")).toBe("FFFFFF"));
  it("drops the alpha part of 8-digit hex", () => expect(toHex6("#11223344")).toBe("112233"));
  it("converts rgb() colors", () => expect(toHex6("rgb(255, 0, 16)")).toBe("FF0010"));
  it("falls back for unknown colors like 'transparent'", () => expect(toHex6("transparent", "FFFFFF")).toBe("FFFFFF"));
  it("never produces a negative or zero size", () => {
    const el = sanitizeElement(makeElement("text", { w: -5, h: 0 }));
    expect(el.w).toBeGreaterThan(0);
    expect(el.h).toBeGreaterThan(0);
  });
  it("strips XML-illegal control characters from text", () =>
    expect(sanitizeElement(makeElement("text", { text: "a\u0007b" })).text).toBe("ab"));
});
