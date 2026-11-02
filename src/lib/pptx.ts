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

const NAMED: Record<string, string> = {
  white: "FFFFFF", black: "000000", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
  orange: "FFA500", purple: "800080", gray: "808080", grey: "808080", pink: "FFC0CB", navy: "000080",
  teal: "008080", silver: "C0C0C0", gold: "FFD700", brown: "A52A2A", cyan: "00FFFF", magenta: "FF00FF",
};
const h2 = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0");

/** Normalize any CSS-ish color to the 6-digit RRGGBB PowerPoint requires (it rejects files with anything else). */
export function toHex6(c: string | undefined | null, fallback = "000000"): string {
  const v = (c ?? "").trim().toLowerCase();
  let m = v.match(/^#?([0-9a-f]{6})([0-9a-f]{2})?$/);
  if (m) return m[1]!.toUpperCase();
  m = v.match(/^#?([0-9a-f])([0-9a-f])([0-9a-f])([0-9a-f])?$/);
  if (m) return (m[1]! + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = v.match(/^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/);
  if (m) return (h2(+m[1]!) + h2(+m[2]!) + h2(+m[3]!)).toUpperCase();
  if (NAMED[v]) return NAMED[v]!;
  return fallback;
}
const hex = (c: string, fallback?: string) => toHex6(c, fallback);
const isClear = (c: string | undefined) => !c || /^(transparent|none)$/i.test(c.trim()) || /^rgba\(.*,\s*0(\.0+)?\s*\)$/i.test(c.trim());

// Characters that are illegal in XML 1.0 — one stray control char makes PowerPoint refuse the whole file.
// eslint-disable-next-line no-control-regex
const cleanText = (t: string) => (t ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "").replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "");
const num = (n: unknown, d: number) => (typeof n === "number" && Number.isFinite(n) ? n : d);

/** Make every value safe for strict OOXML readers (PowerPoint) — Google Slides tolerates junk that PowerPoint rejects. */
export function sanitizeElement(el: SlideElement): SlideElement {
  return {
    ...el,
    x: num(el.x, 0),
    y: num(el.y, 0),
    w: Math.max(2, num(el.w, 100)),
    h: Math.max(2, num(el.h, 100)),
    rotation: ((num(el.rotation, 0) % 360) + 360) % 360,
    opacity: Math.min(1, Math.max(0, num(el.opacity, 1))),
    fontSize: Math.min(800, Math.max(8, num(el.fontSize, 40))),
    borderWidth: Math.min(100, Math.max(0, num(el.borderWidth, 0))),
    radius: Math.max(0, num(el.radius, 0)),
    text: cleanText(el.text ?? ""),
    color: hex(el.color, "1B1B1F"),
    fill: isClear(el.fill) ? "transparent" : hex(el.fill, "C6F432"),
    borderColor: hex(el.borderColor, "1B1B1F"),
    fontFamily: FONT_FACE[el.fontFamily] ? el.fontFamily : "sans",
    align: ["left", "center", "right"].includes(el.align) ? el.align : "left",
  };
}

function safeUrl(u: string): string | null {
  const t = u.trim();
  if (!t) return null;
  const full = /^(https?:|mailto:)/i.test(t) ? t : `https://${t}`;
  try {
    return /^mailto:/i.test(full) ? encodeURI(decodeURI(full)) : new URL(full).href;
  } catch {
    return null;
  }
}

const SAFE_IMAGE = /^data:image\/(png|jpe?g|gif);base64,/i;

function blobToDataUrl(blob: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as string);
    r.onerror = () => resolve(null);
    r.readAsDataURL(blob);
  });
}

/** Re-encode any image (webp, avif, svg, ...) as PNG — PowerPoint refuses files containing those formats. */
function rasterizeToPng(blob: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      try {
        const w = img.naturalWidth || 1600;
        const h = img.naturalHeight || 900;
        const c = document.createElement("canvas");
        c.width = w;
        c.height = h;
        c.getContext("2d")!.drawImage(img, 0, 0, w, h);
        resolve(c.toDataURL("image/png"));
      } catch {
        resolve(null);
      } finally {
        URL.revokeObjectURL(url);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve(null);
    };
    img.src = url;
  });
}

/** Load an image as a PNG/JPEG/GIF data URL so it can be embedded. Returns null on failure. */
async function toDataUrl(src: string): Promise<string | null> {
  if (SAFE_IMAGE.test(src)) return src;
  try {
    const res = await fetch(src);
    if (!res.ok) return null;
    const blob = await res.blob();
    if (/^image\/(png|jpe?g|gif)$/i.test(blob.type)) {
      const data = await blobToDataUrl(blob);
      if (data && SAFE_IMAGE.test(data)) return data;
    }
    return await rasterizeToPng(blob);
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
  const linkUrl = el.link?.kind === "url" ? safeUrl(el.link.url) : null;
  const hyperlink =
    linkUrl
      ? { url: linkUrl }
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
  const hasFill = !isClear(el.fill);
  const line = el.borderWidth > 0 ? { color: hex(el.borderColor), width: pt(el.borderWidth) / 8 } : { type: "none" as const };

  if (el.type === "shape" && el.shape === "line") {
    slide.addShape("line", {
      ...box,
      line: { color: hex(isClear(el.fill) ? el.borderColor : el.fill), width: Math.max(1, pt(Math.max(4, el.borderWidth)) / 8) },
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

  // 2. Fix pptxgenjs quirks that strict PowerPoint readers (iOS/Mac) reject:
  //    a. a content-type Override for a slideMaster2.xml part that doesn't exist
  //    b. notesMasterIdLst placed after sldIdLst (schema requires it before)
  //    c. notes master sharing theme1.xml with the slide master — PowerPoint always gives it its own theme
  //    d. empty hyperlink attributes + whitespace-padded extLst inside hlinkClick
  const theme1 = entries.find((e) => e.name === "ppt/theme/theme1.xml");
  const hasNotesMaster = entries.some((e) => e.name === "ppt/notesMasters/notesMaster1.xml");
  if (theme1 && hasNotesMaster && !entries.some((e) => e.name === "ppt/theme/theme2.xml")) {
    entries.push({ name: "ppt/theme/theme2.xml", data: theme1.data.slice() });
    patch("ppt/notesMasters/_rels/notesMaster1.xml.rels", (xml) =>
      xml.replace("../theme/theme1.xml", "../theme/theme2.xml").replace(/>\s+</g, "><"),
    );
  }
  patch("[Content_Types].xml", (xml) => {
    let out = xml.replace(/<Override PartName="\/ppt\/slideMasters\/slideMaster2\.xml"[^>]*\/>/, "");
    if (slideshow) out = out.replace("presentationml.presentation.main+xml", "presentationml.slideshow.main+xml");
    if (hasNotesMaster && !out.includes("/ppt/theme/theme2.xml")) {
      out = out.replace(
        "</Types>",
        '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
      );
    }
    return out.replace(/>\s+</g, "><");
  });
  patch("ppt/presentation.xml", (xml) => {
    const m = xml.match(/<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/);
    if (m && xml.indexOf(m[0]) > xml.indexOf("<p:sldIdLst>")) {
      return xml.replace(m[0], "").replace("</p:sldMasterIdLst>", `</p:sldMasterIdLst>${m[0]}`);
    }
    return xml;
  });
  for (const e of entries) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(e.name)) continue;
    patch(e.name, (xml) =>
      xml
        .replace(/<a:hlinkClick([^>]*?)>\s*<a:extLst>[\s\S]*?<\/a:extLst>\s*<\/a:hlinkClick>/g, "<a:hlinkClick$1/>")
        .replace(/<a:hlinkClick[^>]*\/>/g, (tag) => tag.replace(/\s(invalidUrl|action|tgtFrame|tooltip)=""/g, "")),
    );
  }

  const bytes = await writeZip(entries);
  return new Blob([bytes as BlobPart], {
    type: slideshow
      ? "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
      : "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

export async function buildPresentationFile(deck: Deck, slideshow: boolean): Promise<Blob> {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SF", width: 13.333, height: 7.5 });
  pptx.layout = "SF";
  pptx.title = cleanText(deck.title) || "Presentation";

  for (const s of deck.slides) {
    const slide = pptx.addSlide();
    slide.background = { color: hex(s.background, "FFFFFF") };
    const notes = cleanText(s.notes ?? "");
    if (notes) slide.addNotes(notes);
    for (const el of s.elements ?? []) await addElement(slide, sanitizeElement(el), deck.slides.length);
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

export type ExportFormat = "pptx" | "ppsx" | "both";

const safeName = (deck: Deck) => deck.title.replace(/[^\w\- ]+/g, "").trim() || "presentation";

/**
 * Single .pptx without a zip — for iPhone/iPad, where PowerPoint can't open a .zip.
 * Uses the share sheet when available (lets the user pick PowerPoint directly), else downloads.
 */
export async function exportSinglePptx(deck: Deck): Promise<string> {
  const name = `${safeName(deck)}.pptx`;
  const blob = await buildPresentationFile(deck, false);
  const file = new File([blob], name, { type: blob.type });
  const nav = navigator as Navigator & { canShare?: (d: ShareData) => boolean };
  if (nav.share && nav.canShare?.({ files: [file] })) {
    try {
      await nav.share({ files: [file], title: name });
      return name;
    } catch (e) {
      if ((e as Error).name === "AbortError") return name;
      // share blocked (e.g. inside an iframe) — fall through to download
    }
  }
  downloadBlob(file, name);
  return name;
}

async function deliverPresentation(blob: Blob, filename: string, share: boolean) {
  const file = new File([blob], filename, { type: blob.type });
  const nav = navigator as Navigator & { canShare?: (data: ShareData) => boolean };
  if (share && nav.share && nav.canShare?.({ files: [file] })) {
    try {
      await nav.share({ files: [file], title: filename });
      return;
    } catch (error) {
      if ((error as Error).name === "AbortError") return;
    }
  }
  downloadBlob(file, filename);
}

/** Single-format exports are real Office files; the combined option remains a ZIP. */
export async function exportDeck(deck: Deck, format: ExportFormat) {
  const safe = safeName(deck);
  if (format !== "both") {
    const filename = `${safe}.${format}`;
    await deliverPresentation(await buildPresentationFile(deck, format === "ppsx"), filename, true);
    return filename;
  }

  const formats: ("pptx" | "ppsx")[] = format === "both" ? ["pptx", "ppsx"] : [format];
  const entries = [];
  for (const f of formats) {
    const blob = await buildPresentationFile(deck, f === "ppsx");
    entries.push({ name: `${safe}/${safe}.${f}`, data: new Uint8Array(await blob.arrayBuffer()) });
  }
  const bytes = await writeZip(entries);
  downloadBlob(new Blob([bytes as BlobPart], { type: "application/zip" }), `${safe}.zip`);
  return `${safe}.zip`;
}

// re-export type used above
export type { Slide };
