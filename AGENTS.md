<!-- LOVABLE:BEGIN -->
> [!IMPORTANT]
> This project is connected to [Lovable](https://lovable.dev). Avoid rewriting
> published git history — force pushing, or rebasing/amending/squashing commits
> that are already pushed — as it rewrites history on Lovable's side and the
> user will likely lose their project history.
>
> Commits you push to the connected branch sync back to Lovable and show up in
> the editor, so keep the branch in a working state.
<!-- LOVABLE:END -->

## Architecture rules
- Presentations are stored in browser localStorage via `src/lib/deck.ts` (no backend yet) — keeps the editor simple; migrate to Lovable Cloud if sharing/sync is needed.
- Slides render at a fixed 1920x1080 and are scaled with `ScaledSlide` everywhere (editor, thumbnails, viewer) so layouts look identical.
- The AI assistant runs server-side (`src/lib/ai.server.ts`) and returns structured slide JSON that the editor applies; keep the API key and prompts on the server.
- PowerPoint export lives in `src/lib/pptx.ts` (pptxgenjs + post-processing). The post-process step is required: it fixes pptxgenjs quirks (bogus slideMaster2 Override, misplaced notesMasterIdLst), injects slide transitions, and flips the content type for .ppsx. The final zip is rebuilt with the custom writer in `src/lib/zip.ts` — JSZip's generateAsync emits zero-length directory entries and stores parts uncompressed, which PowerPoint mobile rejects ("Sorry, can't open file"). Use OOXML preset `star5` for stars — pptxgenjs's `star5Point` writes an invalid prst that PowerPoint/LibreOffice drop silently.
- AI chat history is stored per deck in localStorage (`slideforge.ai.<deckId>`) via `loadAiChat`/`saveAiChat` in `src/lib/deck.ts`; the panel also sends recent history back to the assistant for follow-up context.
- Exports download as a ZIP (built with `writeZip`) holding a folder named after the deck with the .pptx/.ppsx inside. Images are embedded only as PNG/JPEG/GIF; anything else (webp, avif, svg) is re-encoded to PNG via canvas because PowerPoint refuses files containing those formats while Google Slides tolerates them.
