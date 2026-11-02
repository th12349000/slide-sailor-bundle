import { describe, expect, it } from "vitest";
import { makeSlidesFromAi } from "@/lib/ai-slide-mapper";
import type { AiResult } from "@/lib/ai.types";

describe("AI slide image attribution", () => {
  it("keeps image credits visible and in speaker notes", () => {
    const slides: AiResult["slides"] = [
      {
        background: "#ffffff",
        notes: "",
        elements: [
          {
            type: "image",
            x: 120,
            y: 120,
            w: 900,
            h: 500,
            text: "",
            fontSize: 24,
            fontFamily: "sans",
            color: "#1b1b1f",
            bold: false,
            italic: false,
            align: "left",
            fill: "transparent",
            borderColor: "#1b1b1f",
            borderWidth: 0,
            radius: 0,
            shape: "rect",
            imageUrl: "https://upload.wikimedia.org/example.jpg",
            imageAttribution: "A. Photographer",
            imageSourceUrl: "https://commons.wikimedia.org/wiki/File:Example.jpg",
            imageLicense: "CC BY-SA 4.0",
            imageLicenseUrl: "https://creativecommons.org/licenses/by-sa/4.0/",
            linkUrl: null,
            linkSlideIndex: null,
          },
        ],
      },
    ];

    const [slide] = makeSlidesFromAi(slides);

    expect(slide!.elements).toHaveLength(2);
    expect(slide!.elements[0]!.link).toEqual({
      kind: "url",
      url: "https://commons.wikimedia.org/wiki/File:Example.jpg",
    });
    expect(slide!.elements[1]!.text).toContain("A. Photographer");
    expect(slide!.notes).toContain("CC BY-SA 4.0");
  });
});
