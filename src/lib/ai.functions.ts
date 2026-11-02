import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

export function aiErrorMessage(e: unknown): string {
  const err = e as { statusCode?: number; message?: string; responseBody?: string };
  let gatewayMsg: string | undefined;
  try { gatewayMsg = JSON.parse(err.responseBody ?? "").message; } catch { /* not json */ }
  const status = err.statusCode;
  return status === 402
    ? `AI credits have run out (${gatewayMsg ?? "Not enough credits"}). Add credits in your workspace to keep using the assistant.`
    : status === 429
      ? "The AI is busy right now — please try again in a moment."
      : status === 403
        ? gatewayMsg ?? "AI access is blocked for this workspace."
        : gatewayMsg ?? err.message ?? "The AI request failed.";
}

export const askSlideAiInput = z.object({
  prompt: z.string().min(1).max(4000),
  deckJson: z.string().max(200_000),
  currentIndex: z.number().int().min(0),
});

export const askSlideAi = createServerFn({ method: "POST" })
  .inputValidator(
    z.object({
      prompt: z.string().min(1).max(4000),
      deckJson: z.string().max(200_000),
      currentIndex: z.number().int().min(0),
    }),
  )
  .handler(async ({ data }) => {
    const { runSlideAssistant } = await import("./ai.server");
    try {
      return { ok: true as const, result: await runSlideAssistant(data) };
    } catch (e: unknown) {
      const err = e as { statusCode?: number; message?: string; responseBody?: string };
      let gatewayMsg: string | undefined;
      try { gatewayMsg = JSON.parse(err.responseBody ?? "").message; } catch { /* not json */ }
      const status = err.statusCode;
      const message =
        status === 402
          ? `AI credits have run out (${gatewayMsg ?? "Not enough credits"}). Add credits in your workspace to keep using the assistant.`
          : status === 429
            ? "The AI is busy right now — please try again in a moment."
            : status === 403
              ? gatewayMsg ?? "AI access is blocked for this workspace."
              : gatewayMsg ?? err.message ?? "The AI request failed.";
      return { ok: false as const, error: message };
    }
  });
