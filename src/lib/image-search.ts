export interface LicensedImage {
  title: string;
  imageUrl: string;
  sourceUrl: string;
  attribution: string;
  license: string;
  licenseUrl: string;
}

function plainText(value: string | undefined): string {
  return (value ?? "")
    .replace(/<[^>]*>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function searchTerms(prompt: string): string {
  const omit = new Set([
    "about",
    "image",
    "images",
    "photo",
    "photos",
    "visual",
    "visuals",
    "create",
    "build",
    "make",
    "include",
    "relevant",
    "please",
    "slide",
    "slides",
    "presentation",
    "presentations",
  ]);
  return prompt
    .replace(/https?:\/\/\S+/gi, " ")
    .split(/[^\p{L}\p{N}-]+/u)
    .filter((term) => term.length > 2 && !omit.has(term.toLowerCase()))
    .slice(0, 8)
    .join(" ");
}

export async function searchLicensedImages(prompt: string): Promise<LicensedImage[]> {
  const terms = searchTerms(prompt);
  if (!terms) return [];
  const url = new URL("https://commons.wikimedia.org/w/api.php");
  url.search = new URLSearchParams({
    action: "query",
    generator: "search",
    gsrsearch: `filetype:bitmap ${terms}`,
    gsrnamespace: "6",
    gsrlimit: "16",
    prop: "imageinfo",
    iiprop: "url|extmetadata",
    iiurlwidth: "1200",
    format: "json",
    origin: "*",
  }).toString();
  try {
    const response = await fetch(url);
    if (!response.ok) return [];
    const payload = (await response.json()) as {
      query?: {
        pages?: Record<
          string,
          {
            title?: string;
            imageinfo?: {
              thumburl?: string;
              descriptionurl?: string;
              extmetadata?: Record<string, { value?: string }>;
            }[];
          }
        >;
      };
    };
    return Object.values(payload.query?.pages ?? [])
      .flatMap((page) => {
        const info = page.imageinfo?.[0];
        if (!info) return [];
        const metadata = info.extmetadata ?? {};
        const license = plainText(metadata["LicenseShortName"]?.value);
        if (!/^(?:CC0|Public domain|PD|CC BY(?:-SA)?\s+[1-4](?:\.0)?)/i.test(license)) return [];
        const attribution =
          plainText(metadata["Attribution"]?.value) || plainText(metadata["Artist"]?.value);
        if (!attribution || !info.thumburl || !info.descriptionurl) return [];
        return [
          {
            title: plainText(page.title),
            imageUrl: info.thumburl,
            sourceUrl: info.descriptionurl,
            attribution,
            license,
            licenseUrl: metadata["LicenseUrl"]?.value ?? "https://creativecommons.org/",
          },
        ];
      })
      .slice(0, 6);
  } catch {
    return [];
  }
}
