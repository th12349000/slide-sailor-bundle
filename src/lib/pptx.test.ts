import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Blob as NodeBlob } from "node:buffer";
import JSZip from "jszip";
import { buildPresentationFile, exportDeck, sanitizeElement, toHex6 } from "@/lib/pptx";
import { makeDeck, makeElement } from "@/lib/deck";

const browserBlob = globalThis.Blob;

beforeEach(() => {
  globalThis.Blob = NodeBlob as unknown as typeof Blob;
});

afterEach(() => {
  vi.unstubAllGlobals();
  globalThis.Blob = browserBlob;
});

describe("PowerPoint-safe values", () => {
  it("expands short hex colors to 6 digits", () => expect(toHex6("#fff")).toBe("FFFFFF"));
  it("drops the alpha part of 8-digit hex", () => expect(toHex6("#11223344")).toBe("112233"));
  it("converts rgb() colors", () => expect(toHex6("rgb(255, 0, 16)")).toBe("FF0010"));
  it("falls back for unknown colors like 'transparent'", () =>
    expect(toHex6("transparent", "FFFFFF")).toBe("FFFFFF"));
  it("never produces a negative or zero size", () => {
    const el = sanitizeElement(makeElement("text", { w: -5, h: 0 }));
    expect(el.w).toBeGreaterThan(0);
    expect(el.h).toBeGreaterThan(0);
  });
  it("strips XML-illegal control characters from text", () =>
    expect(sanitizeElement(makeElement("text", { text: "a\u0007b" })).text).toBe("ab"));

  it("builds a PowerPoint-readable OOXML package with the correct MIME type", async () => {
    const blob = await buildPresentationFile(makeDeck("PowerPoint export"), false);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const contentTypes = await zip.file("[Content_Types].xml")!.async("text");
    const slideXml = await zip.file("ppt/slides/slide1.xml")!.async("text");
    const parsedSlide = new DOMParser().parseFromString(slideXml, "application/xml");

    expect(blob.type).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    );
    expect(contentTypes).toContain("presentationml.presentation.main+xml");
    expect(parsedSlide.querySelector("parsererror")).toBeNull();
    expect(zip.file("ppt/presentation.xml")).not.toBeNull();
  });

  it("uses the slideshow content type for PPSX exports", async () => {
    const blob = await buildPresentationFile(makeDeck("Slideshow export"), true);
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const contentTypes = await zip.file("[Content_Types].xml")!.async("text");

    expect(blob.type).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
    );
    expect(contentTypes).toContain("presentationml.slideshow.main+xml");
  });

  it("shares a standalone PPTX file instead of wrapping a single export in ZIP", async () => {
    const share = vi.fn().mockResolvedValue(undefined);
    vi.stubGlobal("navigator", { share, canShare: () => true });

    const filename = await exportDeck(makeDeck("Direct export"), "pptx");
    const sharedFile = share.mock.calls[0]?.[0].files[0] as File;

    expect(filename).toBe("Direct export.pptx");
    expect(sharedFile.name).toBe(filename);
    expect(sharedFile.type).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    );
  });
});
