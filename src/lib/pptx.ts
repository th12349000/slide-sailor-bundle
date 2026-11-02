import PptxGenJS from "pptxgenjs";
import JSZip from "jszip";
import { SLIDE_H, SLIDE_W, type Deck, type Slide, type SlideElement, type TransitionKind } from "@/lib/deck";
import { writeZip } from "@/lib/zip";

// Canvas is 1920x1080 px; PowerPoint widescreen is 13.333x7.5 in → 144 px per inch.
const PX = 144;
const ix = (px: number) => px / PX;
// PowerPoint font sizes are points; slide px ≈ pt * (144/96) at 1080p → pt = px * 0.75
const pt = (px: number) => Math.max(8, Math.round(px * 0.75));

const FONT_FACE: Record<SlideElement["fontFamily"], string> = {
  display: "Trebuchet MS",
  sans: "Calibri",
  serif: "Georgia",
  mono: "Consolas",
};

const hex = (c: string) => (c.startsWith("#") ? c.slice(1) : c);

/** Fetch a remote image into a data URL so it can be embedded. Returns null on failure. */
async function toDataUrl(src: string): Promise<string | null> {
  if (src.startsWith("data:")) return src;
  try {
    const res = await fetch(src);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((resolve) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result as string);
      r.onerror = () => resolve(null);
      r.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

function shapeType(el: SlideElement): string {
  switch (el.shape) {
    case "ellipse": return "ellipse";
    case "triangle": return "triangle";
    case "diamond": return "diamond";
    case "star": return "star5"; // OOXML preset name; pptxgenjs's "star5Point" writes an invalid prst
    case "arrow": return "rightArrow";
    case "rect": return el.radius > 0 ? "roundRect" : "rect";
    default: return "rect";
  }
}

async function addElement(slide: PptxGenJS.Slide, el: SlideElement, slideCount: number) {
  const box = { x: ix(el.x), y: ix(el.y), w: ix(el.w), h: ix(el.h), rotate: el.rotation || 0 };
  const hyperlink =
    el.link?.kind === "url" && el.link.url.trim()
      ? { url: /^(https?:|mailto:)/i.test(el.link.url) ? el.link.url : `https://${el.link.url}` }
      : el.link?.kind === "slide" && el.link.slideIndex >= 0 && el.link.slideIndex < slideCount
        ? { slide: el.link.slideIndex + 1 }
        : undefined;

  if (el.type === "image") {
    const data = el.src ? await toDataUrl(el.src) : null;
    if (data) {
      slide.addImage({ data, ...box, transparency: Math.round((1 - el.opacity) * 100) });
    }
    return;
  }

  const transparency = Math.round((1 - el.opacity) * 100);
  const hasFill = el.fill && el.fill !== "transparent";
  const line = el.borderWidth > 0 ? { color: hex(el.borderColor), width: pt(el.borderWidth) / 8 } : { type: "none" as const };

  if (el.type === "shape" && el.shape === "line") {
    slide.addShape("line", {
      ...box,
      line: { color: hex(el.fill === "transparent" ? el.borderColor : el.fill), width: Math.max(1, pt(Math.max(4, el.borderWidth)) / 8) },
    });
    return;
  }

  if (el.type === "shape") {
    slide.addShape(shapeType(el) as PptxGenJS.ShapeType, {
      ...box,
      fill: hasFill ? { color: hex(el.fill), transparency } : { type: "none" },
      line,
      ...(el.radius > 0 && el.shape === "rect" ? { rectRadius: ix(el.radius) } : {}),
    });
    if (el.text) {
      slide.addText(el.text, {
        ...box,
        fontFace: FONT_FACE[el.fontFamily],
        fontSize: pt(el.fontSize),
        color: hex(el.color),
        bold: el.bold,
        italic: el.italic,
        ...(el.underline ? { underline: { style: "sng" as const } } : {}),
        align: el.align,
        valign: "middle",
        margin: 12,
        ...(hyperlink ? { hyperlink } : {}),
      });
    }
    return;
  }

  // text element
  slide.addText(el.text || " ", {
    ...box,
    fontFace: FONT_FACE[el.fontFamily],
    fontSize: pt(el.fontSize),
    color: hex(el.color),
    bold: el.bold,
    italic: el.italic,
    ...(el.underline ? { underline: { style: "sng" as const } } : {}),
    align: el.align,
    valign: "top",
    margin: 0,
    ...(hasFill ? { fill: { color: hex(el.fill), transparency } } : {}),
    ...(el.borderWidth > 0 ? { line } : {}),
    ...(hyperlink ? { hyperlink } : {}),
  });
}

// PowerPoint transition XML for each of our transition kinds.
const TRANSITION_XML: Record<Exclude<TransitionKind, "none">, string> = {
  fade: '<p:transition spd="med"><p:fade/></p:transition>',
  slide: '<p:transition spd="med"><p:push dir="l"/></p:transition>',
  wipe: '<p:transition spd="med"><p:wipe dir="l"/></p:transition>',
  zoom: '<p:transition spd="med"><p:fade/></p:transition>',
};

/**
 * Inject slide transitions into the generated package and optionally flip it to a .ppsx slideshow.
 * Rebuilds the zip with writeZip: JSZip's generateAsync emits zero-length directory entries and
 * stores parts uncompressed, which strict readers (PowerPoint mobile) reject.
 */
async function postProcess(buf: ArrayBuffer, deck: Deck, slideshow: boolean): Promise<Blob> {
  const src = await JSZip.loadAsync(buf);
  const dec = new TextDecoder();
  const enc = new TextEncoder();

  // Read every real file entry (skip directory entries) as text for patching.
  const entries: { name: string; data: Uint8Array }[] = [];
  for (const f of Object.values(src.files)) {
    if (f.dir) continue;
    entries.push({ name: f.name, data: await f.async("uint8array") });
  }
  const patch = (name: string, fn: (xml: string) => string) => {
    const e = entries.find((x) => x.name === name);
    if (e) e.data = enc.encode(fn(dec.decode(e.data)));
  };

  // 1. Inject the chosen transition into each slide.
  for (let i = 0; i < deck.slides.length; i++) {
    const t = deck.slides[i]!.transition ?? "none";
    if (t === "none") continue;
    patch(`ppt/slides/slide${i + 1}.xml`, (xml) => xml.replace("</p:sld>", `${TRANSITION_XML[t]}</p:sld>`));
  }

  // 2. Fix two pptxgenjs quirks that strict PowerPoint validation rejects:
  //    a. a content-type Override for a slideMaster2.xml part that doesn't exist
  //    b. notesMasterIdLst placed after sldIdLst (schema requires it before)
  patch("[Content_Types].xml", (xml) => {
    let out = xml.replace(/<Override PartName="\/ppt\/slideMasters\/slideMaster2\.xml"[^>]*\/>/, "");
    if (slideshow) out = out.replace("presentationml.presentation.main+xml", "presentationml.slideshow.main+xml");
    return out;
  });
  patch("ppt/presentation.xml", (xml) => {
    const m = xml.match(/<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/);
    if (m && xml.indexOf(m[0]) > xml.indexOf("<p:sldIdLst>")) {
      return xml.replace(m[0], "").replace("</p:sldMasterIdLst>", `</p:sldMasterIdLst>${m[0]}`);
    }
    return xml;
  });

  const bytes = await writeZip(entries);
  return new Blob([bytes as BlobPart], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

export async function buildPresentationFile(deck: Deck, slideshow: boolean): Promise<Blob> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SF", width: 13.333, height: 7.5 });
  pptx.layout = "SF";
  pptx.title = deck.title;

  for (const s of deck.slides) {
    const slide = pptx.addSlide();
    slide.background = { color: hex(s.background) };
    if (s.notes) slide.addNotes(s.notes);
    for (const el of s.elements) await addElement(slide, el, deck.slides.length);
  }

  const buf = (await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer;
  return postProcess(buf, deck, slideshow);
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

export async function exportDeck(deck: Deck, format: "pptx" | "ppsx") {
  const blob = await buildPresentationFile(deck, format === "ppsx");
  const safe = deck.title.replace(/[^\w\- ]+/g, "").trim() || "presentation";
  downloadBlob(blob, `${safe}.${format}`);
}

// re-export type used above
export type { Slide };
