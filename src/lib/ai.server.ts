import { createOpenAI } from "@ai-sdk/openai";
import { Output, parsePartialJson, streamText } from "ai";
import { z } from "zod";

const elementSchema = z.object({
  type: z.enum(["text", "shape", "image"]),
  x: z.number(),
  y: z.number(),
  w: z.number(),
  h: z.number(),
  text: z.string(),
  fontSize: z.number(),
  fontFamily: z.enum(["display", "sans", "serif", "mono"]),
  color: z.string(),
  bold: z.boolean(),
  italic: z.boolean(),
  align: z.enum(["left", "center", "right"]),
  fill: z.string(),
  borderColor: z.string(),
  borderWidth: z.number(),
  radius: z.number(),
  shape: z.enum(["rect", "ellipse", "triangle", "star", "arrow", "line", "diamond"]),
  imageUrl: z.string(),
  linkUrl: z.string().nullable(),
  linkSlideIndex: z.number().nullable(),
});

export const aiResultSchema = z.object({
  reply: z.string(),
  mode: z.enum(["none", "append", "replace_current", "replace_all"]),
  slides: z.array(
    z.object({
      background: z.string(),
      notes: z.string(),
      elements: z.array(elementSchema),
    }),
  ),
});

export type AiResult = z.infer<typeof aiResultSchema>;

const INSTRUCTIONS = `You are a slide design assistant inside a presentation editor.
Canvas is 1920x1080 px. Coordinates are px from top-left. Keep content inside 80px margins.
Design rules: one idea per slide, max 4 bullets, titles 80-120px, body 36-48px, captions 24-28px. Use bold, tasteful color palettes with strong contrast; hex colors only (or "transparent" for fill).
Text elements: put bullets as lines starting with "• " separated by \\n. Set fill "transparent" unless a text box background is wanted.
Shapes: decorative blocks, dividers, callouts. A shape may also contain text.
Images: only use imageUrl when the user gave a URL; otherwise use "" and avoid image elements.
Links: linkUrl for websites, linkSlideIndex (0-based) to jump to a slide in this deck; null otherwise.
mode: "append" adds new slides after the current one; "replace_current" replaces the current slide with slides[0]; "replace_all" replaces the entire deck; "none" when just answering a question (slides = []).
reply: a short friendly sentence about what you did.`;

export type AiProgress =
  | { type: "reasoning"; text: string }
  | { type: "progress"; slides: number; elements: number; currentTitle: string };

export async function runSlideAssistant(
  input: { prompt: string; deckJson: string; currentIndex: number },
  onProgress?: (p: AiProgress) => void,
): Promise<AiResult> {
  const apiKey = process.env["LOVABLE_API_KEY"];
  if (!apiKey) throw new Error("AI is not configured.");
  const provider = createOpenAI({
    baseURL: "https://ai.gateway.lovable.dev/v1",
    apiKey,
    headers: { "Lovable-API-Key": apiKey, "X-Lovable-AIG-SDK": "vercel-ai-sdk" },
  });
  let streamError: unknown;
  const result = streamText({
    onError: ({ error }) => { streamError = error; },
    model: provider.responses("openai/gpt-6-astra"),
    instructions: INSTRUCTIONS,
    messages: [
      {
        role: "user",
        content: `Current deck (JSON, images stripped):\n${input.deckJson}\n\nCurrent slide index: ${input.currentIndex}\n\nRequest: ${input.prompt}`,
      },
    ],
    output: Output.object({ schema: aiResultSchema }),
    providerOptions: {
      openai: {
        forceReasoning: true,
        reasoningEffort: "low",
        reasoningSummary: "auto",
        store: false,
        include: ["reasoning.encrypted_content"],
      },
    },
  } as Parameters<typeof streamText>[0]);
  let raw = "";
  let lastKey = "";
  for await (const part of result.fullStream) {
    if (part.type === "error") streamError = part.error;
    else if (part.type === "reasoning-delta" && part.text) onProgress?.({ type: "reasoning", text: part.text });
    else if (part.type === "text-delta" && onProgress) {
      raw += part.text;
      const { value } = await parsePartialJson(raw);
      const slides = (value as { slides?: { elements?: { text?: string }[] }[] } | undefined)?.slides ?? [];
      const elements = slides.reduce((n, s) => n + (s?.elements?.length ?? 0), 0);
      const last = slides[slides.length - 1];
      const currentTitle = (last?.elements?.find((e) => e?.text)?.text ?? "").split("\n")[0]!.slice(0, 60);
      const key = `${slides.length}:${elements}:${currentTitle}`;
      if (key !== lastKey) {
        lastKey = key;
        onProgress({ type: "progress", slides: slides.length, elements, currentTitle });
      }
    }
  }
  if (streamError) throw streamError;
  try {
    return (await result.output) as AiResult;
  } catch (e) {
    throw streamError ?? e;
  }
}
