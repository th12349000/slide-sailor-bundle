import { makeElement, makeSlide, type Slide, type SlideElement } from "@/lib/deck";
import type { AiResult } from "@/lib/ai.types";

function mapElement(element: AiResult["slides"][number]["elements"][number]): SlideElement {
  return makeElement(element.type, {
    x: element.x,
    y: element.y,
    w: element.w,
    h: element.h,
    text: element.text,
    fontSize: element.fontSize,
    fontFamily: element.fontFamily,
    color: element.color,
    bold: element.bold,
    italic: element.italic,
    align: element.align,
    fill: element.fill,
    borderColor: element.borderColor,
    borderWidth: element.borderWidth,
    radius: element.radius,
    shape: element.shape,
    src: element.imageUrl,
    link: element.linkUrl
      ? { kind: "url", url: element.linkUrl }
      : element.linkSlideIndex !== null
        ? { kind: "slide", slideIndex: element.linkSlideIndex }
        : element.type === "image" && element.imageSourceUrl
          ? { kind: "url", url: element.imageSourceUrl }
          : null,
  });
}

export function makeSlidesFromAi(aiSlides: AiResult["slides"]): Slide[] {
  return aiSlides.map((slide) => {
    const elements = slide.elements.map(mapElement);
    const credits: SlideElement[] = [];
    const notes = [
      ...slide.elements.flatMap((element) => {
        if (element.type !== "image" || !element.imageAttribution || !element.imageSourceUrl)
          return [];
        const license = element.imageLicense || "Licensed image";
        const attribution = `${element.imageAttribution} · ${license}`;
        credits.push(
          makeElement("text", {
            x: element.x + 8,
            y: Math.min(1040, element.y + element.h - 34),
            w: Math.max(100, element.w - 16),
            h: 26,
            text: attribution.slice(0, 180),
            fontSize: 18,
            color: "#FFFFFF",
            fill: "#18211F",
            opacity: 0.82,
            link: { kind: "url", url: element.imageSourceUrl },
          }),
        );
        return [
          `${element.imageAttribution}. ${license}. ${element.imageSourceUrl}${element.imageLicenseUrl ? ` (${element.imageLicenseUrl})` : ""}`,
        ];
      }),
    ];
    return makeSlide({
      background: slide.background,
      notes: [slide.notes, ...(notes.length ? [`Image credits:\n${notes.join("\n")}`] : [])]
        .filter(Boolean)
        .join("\n\n"),
      elements: [...elements, ...credits],
    });
  });
}
