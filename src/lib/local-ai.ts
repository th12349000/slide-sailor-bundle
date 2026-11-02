import {
  officeAiResultSchema,
  slideAiResultSchema,
  type AiResult,
  type OfficeAiResult,
} from "@/lib/ai.types";
import { searchLicensedImages } from "@/lib/image-search";

const MODEL_ID = "onnx-community/SmolLM2-360M-Instruct-ONNX";

export interface LocalAiStatus {
  phase: "downloading" | "generating" | "ready";
  progress?: number;
  file?: string;
}

export function localAiStatusLabel(status: LocalAiStatus): string {
  if (status.phase === "generating") return "AI is working on this device…";
  if (status.phase === "ready") return "Local AI is ready";
  return status.progress === undefined
    ? "Downloading the local AI model…"
    : `Downloading local AI model… ${Math.round(status.progress)}%`;
}

export function localAiErrorMessage(error: unknown): string {
  const message = (error as Error)?.message ?? String(error);
  if (/fetch|network|connection|offline|load failed|failed to load/i.test(message)) {
    return `Couldn't download the local AI model. Connect once to the internet, then retry. The model will be cached in this browser. (${message.slice(0, 180)})`;
  }
  if (/memory|webgpu|wasm|allocation/i.test(message)) {
    return `This browser couldn't start the local model. Try a recent Safari or Chrome and close other tabs. (${message.slice(0, 180)})`;
  }
  return message
    ? `Local AI failed: ${message.slice(0, 240)}`
    : "The local AI request failed. Please try again.";
}

type LocalMessage = { role: "system" | "user"; content: string };
type ModelReply = { generated_text: { role: string; content: string }[] }[];
type LocalGenerator = (
  messages: LocalMessage[],
  options: { max_new_tokens: number; do_sample: boolean },
) => Promise<ModelReply>;
type ProgressInfo = { status?: string; file?: string; progress?: number };

let generatorPromise: Promise<LocalGenerator> | undefined;
let generatorDevice: "webgpu" | "wasm" = "wasm";

const slideInstructions = `You are a presentation design assistant. Return exactly one valid JSON object, no markdown. Use this schema: {"reply":string,"mode":"none|append|replace_current|replace_all","slides":[{"background":"#RRGGBB","notes":string,"elements":[{"type":"text|shape|image","x":number,"y":number,"w":number,"h":number,"text":string,"fontSize":number,"fontFamily":"display|sans|serif|mono","color":"#RRGGBB","bold":boolean,"italic":boolean,"align":"left|center|right","fill":"#RRGGBB|transparent","borderColor":"#RRGGBB","borderWidth":number,"radius":number,"shape":"rect|ellipse|triangle|star|arrow|line|diamond","imageUrl":string,"imageAttribution":string,"imageSourceUrl":string,"imageLicense":string,"imageLicenseUrl":string,"linkUrl":string|null,"linkSlideIndex":number|null}]}]}. Canvas is 1920x1080. Keep elements within bounds and use readable contrast. For images, only select an exact imageUrl from the supplied licensed candidates and copy all its attribution/license metadata exactly. Never invent image URLs. If none are relevant, use no image elements. For a question without slide edits, set mode to none and slides to [].`;

function progressFrom(info: ProgressInfo): LocalAiStatus {
  return {
    phase: "downloading",
    ...(typeof info.progress === "number"
      ? { progress: Math.max(0, Math.min(100, info.progress)) }
      : {}),
    ...(info.file ? { file: info.file.split("/").at(-1)! } : {}),
  };
}

async function loadGenerator(
  onStatus?: (status: LocalAiStatus) => void,
  forceWasm = false,
): Promise<LocalGenerator> {
  if (!generatorPromise) {
    generatorPromise = (async () => {
      const { env, pipeline } = await import("@huggingface/transformers");
      env.allowRemoteModels = true;
      env.allowLocalModels = false;
      env.useBrowserCache = true;
      const hasWebGpu =
        !forceWasm && typeof navigator !== "undefined" && Boolean("gpu" in navigator);
      const progress_callback = (info: ProgressInfo) => onStatus?.(progressFrom(info));
      if (hasWebGpu) {
        try {
          const generator = await pipeline("text-generation", MODEL_ID, {
            device: "webgpu",
            dtype: "q4f16",
            progress_callback,
          });
          generatorDevice = "webgpu";
          return generator as unknown as LocalGenerator;
        } catch {
          onStatus?.({ phase: "downloading", file: "Switching to CPU mode" });
        }
      }
      try {
        const generator = await pipeline("text-generation", MODEL_ID, {
          device: "wasm",
          dtype: "q4",
          progress_callback,
        });
        generatorDevice = "wasm";
        return generator as unknown as LocalGenerator;
      } catch (error) {
        if (hasWebGpu) onStatus?.({ phase: "downloading", file: "CPU fallback failed" });
        throw error;
      }
    })();
    generatorPromise.catch(() => {
      generatorPromise = undefined;
    });
  }
  return generatorPromise;
}

function extractJson(text: string): string {
  const start = text.indexOf("{");
  if (start < 0)
    throw new Error("The local model did not return JSON. Try a shorter, more specific request.");
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = start; index < text.length; index++) {
    const character = text[index]!;
    if (inString) {
      if (escaped) escaped = false;
      else if (character === "\\") escaped = true;
      else if (character === '"') inString = false;
      continue;
    }
    if (character === '"') inString = true;
    else if (character === "{") depth++;
    else if (character === "}" && --depth === 0) return text.slice(start, index + 1);
  }
  throw new Error("The local model response was incomplete. Please try again.");
}

async function generateStructured<T>(input: {
  system: string;
  prompt: string;
  maxTokens: number;
  onStatus?: ((status: LocalAiStatus) => void) | undefined;
  schema: { parse(value: unknown): T };
}): Promise<T> {
  input.onStatus?.({ phase: "downloading" });
  const generator = await loadGenerator(input.onStatus);
  input.onStatus?.({ phase: "generating" });
  const messages = [
    { role: "system" as const, content: input.system },
    { role: "user" as const, content: input.prompt },
  ];
  const options = { max_new_tokens: input.maxTokens, do_sample: false };
  let output: ModelReply;
  try {
    output = await generator(messages, options);
  } catch (error) {
    if (generatorDevice !== "webgpu") throw error;
    generatorPromise = undefined;
    const cpuGenerator = await loadGenerator(input.onStatus, true);
    output = await cpuGenerator(messages, options);
  }
  const response = output[0]?.generated_text.at(-1)?.content;
  if (!response) throw new Error("The local model returned an empty response. Please try again.");
  const result = input.schema.parse(JSON.parse(extractJson(response)));
  input.onStatus?.({ phase: "ready" });
  return result;
}

export async function generateSlidesLocally(input: {
  prompt: string;
  deckJson: string;
  currentIndex: number;
  onStatus?: (status: LocalAiStatus) => void;
}): Promise<AiResult> {
  input.onStatus?.({ phase: "downloading" });
  const candidates =
    /image|photo|photograph|visual|picture|illustration|cover|background|build|create|generate/i.test(
      input.prompt,
    )
      ? await searchLicensedImages(input.prompt)
      : [];
  const result = await generateStructured({
    system: slideInstructions,
    prompt: `Current deck: ${input.deckJson.slice(0, 24_000)}\nCurrent slide index: ${input.currentIndex}\nLicensed image candidates: ${JSON.stringify(candidates)}\nRequest: ${input.prompt}`,
    maxTokens: 3072,
    onStatus: input.onStatus,
    schema: slideAiResultSchema,
  });
  const candidateUrls = new Set(candidates.map((candidate) => candidate.imageUrl));
  return {
    ...result,
    slides: result.slides.map((slide) => ({
      ...slide,
      elements: slide.elements.map((element) =>
        element.type === "image" && !candidateUrls.has(element.imageUrl)
          ? {
              ...element,
              imageUrl: "",
              imageAttribution: "",
              imageSourceUrl: "",
              imageLicense: "",
              imageLicenseUrl: "",
            }
          : element,
      ),
    })),
  };
}

export async function generateOfficeLocally(input: {
  kind: "spreadsheet" | "document";
  prompt: string;
  dataJson: string;
  onStatus?: (status: LocalAiStatus) => void;
}): Promise<OfficeAiResult> {
  const system =
    input.kind === "spreadsheet"
      ? `You are an Excel spreadsheet assistant. Return exactly one valid JSON object, no markdown, with schema {"reply":string,"workbook":{"title":string,"sheets":[{"name":string,"cells":string[][],"conditionalFormats":[{"range":string,"operator":"greaterThan|lessThan|equal|containsText","value":string,"color":"#RRGGBB","background":"#RRGGBB"}],"charts":[{"title":string,"range":string,"kind":"bar|line|pie"}]}]},"document":null}. Use real Excel formulas beginning with =, including common functions and references. Use useful headers, totals, and formulas. Conditional formats and charts should use valid A1 ranges. Keep output concise and formula strings intact.`
      : `You are a Word writing assistant. Return exactly one valid JSON object, no markdown, with schema {"reply":string,"workbook":null,"document":{"title":string,"blocks":[{"type":"heading1|heading2|paragraph|bullet|numbered","text":string}]}}. Draft or revise the current document according to the user's request. Organize long documents with headings and concise paragraphs. Use lists when useful.`;
  const context =
    input.dataJson.length > 24_000
      ? `${input.dataJson.slice(0, 24_000)}\n[Current content truncated for local model context]`
      : input.dataJson;
  return generateStructured({
    system,
    prompt: `Current ${input.kind}: ${context}\n\nRequest: ${input.prompt}`,
    maxTokens: 2048,
    onStatus: input.onStatus,
    schema: officeAiResultSchema,
  });
}
