import { afterEach, describe, expect, it, vi } from "vitest";
import { searchLicensedImages } from "@/lib/image-search";

afterEach(() => vi.unstubAllGlobals());

describe("licensed image search", () => {
  it("returns only freely licensed images with attribution and source metadata", async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        new Response(
          JSON.stringify({
            query: {
              pages: {
                "1": {
                  title: "File:Solar panels.jpg",
                  imageinfo: [
                    {
                      thumburl: "https://upload.wikimedia.org/solar.jpg",
                      descriptionurl: "https://commons.wikimedia.org/wiki/File:Solar_panels.jpg",
                      extmetadata: {
                        LicenseShortName: { value: "CC BY-SA 4.0" },
                        Attribution: { value: "Alex Example" },
                        LicenseUrl: { value: "https://creativecommons.org/licenses/by-sa/4.0/" },
                      },
                    },
                  ],
                },
                "2": {
                  title: "File:Restricted image.jpg",
                  imageinfo: [
                    {
                      thumburl: "https://upload.wikimedia.org/restricted.jpg",
                      descriptionurl: "https://commons.wikimedia.org/wiki/File:Restricted.jpg",
                      extmetadata: {
                        LicenseShortName: { value: "CC BY-NC 4.0" },
                        Attribution: { value: "Not free for commercial use" },
                        LicenseUrl: { value: "https://creativecommons.org/licenses/by-nc/4.0/" },
                      },
                    },
                  ],
                },
              },
            },
          }),
        ),
      )
      .mockName("Wikimedia Commons search");
    vi.stubGlobal("fetch", fetchMock);

    const images = await searchLicensedImages("create a solar panel presentation with photos");

    expect(images).toHaveLength(1);
    expect(images[0]).toMatchObject({ attribution: "Alex Example", license: "CC BY-SA 4.0" });
    expect(images[0]!.sourceUrl).toContain("commons.wikimedia.org");
  });
});
