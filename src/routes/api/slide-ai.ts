import { createFileRoute } from "@tanstack/react-router";
import { aiErrorMessage, askSlideAiInput } from "@/lib/ai.functions";

// Streams the AI assistant's progress as NDJSON lines:
// {type:"reasoning",text} | {type:"progress",...} | {type:"done",result} | {type:"error",error}
export const Route = createFileRoute("/api/slide-ai")({
  server: {
    handlers: {
      POST: async ({ request }) => {
        const parsed = askSlideAiInput.safeParse(await request.json().catch(() => null));
        if (!parsed.success) return Response.json({ error: "Invalid request" }, { status: 400 });
        const { runSlideAssistant } = await import("@/lib/ai.server");
        const enc = new TextEncoder();
        const stream = new ReadableStream({
          async start(controller) {
            const send = (obj: unknown) => controller.enqueue(enc.encode(JSON.stringify(obj) + "\n"));
            send({ type: "progress", slides: 0, elements: 0, currentTitle: "" });
            try {
              const result = await runSlideAssistant(parsed.data, send);
              send({ type: "done", result });
            } catch (e) {
              send({ type: "error", error: aiErrorMessage(e) });
            } finally {
              controller.close();
            }
          },
        });
        return new Response(stream, {
          headers: { "Content-Type": "application/x-ndjson", "Cache-Control": "no-cache, no-transform" },
        });
      },
    },
  },
});
