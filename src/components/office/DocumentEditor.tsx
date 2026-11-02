import { useEffect, useState, type FormEvent } from "react";
import { Link } from "@tanstack/react-router";
import {
  ArrowLeft,
  Bold,
  Download,
  Heading1,
  Heading2,
  Italic,
  List,
  ListOrdered,
  Sparkles,
  Strikethrough,
  Table2,
  Wand2,
} from "lucide-react";
import { EditorContent, useEditor } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import { TableKit } from "@tiptap/extension-table";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import type { OfficeDocument } from "@/lib/office";
import { generateOfficeLocally, localAiErrorMessage, localAiStatusLabel, type LocalAiStatus } from "@/lib/local-ai";
import { upsertDocument } from "@/lib/office";
import type { Editor as TiptapEditor, JSONContent } from "@tiptap/core";

type EditorChain = ReturnType<TiptapEditor["chain"]>;

function blocksToContent(
  blocks: { type: "heading1" | "heading2" | "paragraph" | "bullet" | "numbered"; text: string }[],
): JSONContent {
  const content: JSONContent[] = blocks.map((block) => {
    const textNode: JSONContent = { type: "text", text: block.text };
    if (block.type === "heading1" || block.type === "heading2") {
      return {
        type: "heading",
        attrs: { level: block.type === "heading1" ? 1 : 2 },
        content: [textNode],
      };
    }
    if (block.type === "bullet" || block.type === "numbered") {
      return {
        type: block.type === "bullet" ? "bulletList" : "orderedList",
        content: [{ type: "listItem", content: [{ type: "paragraph", content: [textNode] }] }],
      };
    }
    return { type: "paragraph", content: [textNode] };
  });
  return { type: "doc", content: content.length ? content : [{ type: "paragraph" }] };
}

export function DocumentEditor({ initial }: { initial: OfficeDocument }) {
  const [document, setDocument] = useState(initial);
  const [aiPrompt, setAiPrompt] = useState("");
  const [aiBusy, setAiBusy] = useState(false);
  const [aiStatus, setAiStatus] = useState<LocalAiStatus | null>(null);
  const editor = useEditor(
    {
      immediatelyRender: false,
      extensions: [StarterKit, TableKit.configure({ table: { resizable: true } })],
      content: initial.content,
      editorProps: {
        attributes: { class: "document-prose", "aria-label": "Word document content" },
      },
      onUpdate: ({ editor: currentEditor }) =>
        setDocument((current) => ({ ...current, content: currentEditor.getJSON() })),
    },
    [initial.id],
  );

  useEffect(() => {
    const timer = window.setTimeout(() => upsertDocument(document), 450);
    return () => window.clearTimeout(timer);
  }, [document]);

  const runExport = async () => {
    upsertDocument(document);
    try {
      const { exportDocument } = await import("@/lib/office-export");
      await exportDocument(document);
      toast.success("Word document exported.");
    } catch (error) {
      toast.error(`Export failed: ${(error as Error).message}`);
    }
  };

  const applyAi = async (event: FormEvent) => {
    event.preventDefault();
    if (!aiPrompt.trim() || aiBusy) return;
    setAiBusy(true);
    try {
      const result = await generateOfficeLocally({
        kind: "document",
        prompt: aiPrompt.trim(),
        dataJson: JSON.stringify(document),
        onStatus: setAiStatus,
      });
      if (!result.document) {
        toast.info(result.reply || "No document changes were returned.");
        return;
      }
      const content = blocksToContent(result.document.blocks);
      setDocument((current) => ({
        ...current,
        title: result.document!.title || current.title,
        content,
      }));
      editor?.commands.setContent(content);
      setAiPrompt("");
      toast.success(result.reply || "Document updated.");
    } catch (error) {
      toast.error(localAiErrorMessage(error));
    } finally {
      setAiBusy(false);
      setAiStatus(null);
    }
  };

  const run = (action: (chain: EditorChain) => unknown) => {
    if (editor) action(editor.chain().focus());
  };

  return (
    <div className="flex h-screen min-h-[620px] flex-col bg-background">
      <header className="flex min-h-14 shrink-0 items-center gap-2 border-b border-border px-3">
        <Button variant="ghost" size="icon" asChild>
          <Link to="/" aria-label="Back to workspace">
            <ArrowLeft />
          </Link>
        </Button>
        <input
          aria-label="Document title"
          className="h-9 min-w-0 flex-1 border-0 bg-transparent font-display font-bold outline-none md:max-w-80"
          value={document.title}
          onChange={(event) =>
            setDocument((current) => ({ ...current, title: event.target.value }))
          }
        />
        <Button variant="outline" onClick={runExport}>
          <Download /> <span className="hidden sm:inline">Export .docx</span>
        </Button>
      </header>

      <div className="flex shrink-0 flex-wrap items-center gap-1 border-b border-border px-3 py-2">
        <Button
          size="icon"
          variant={editor?.isActive("bold") ? "secondary" : "ghost"}
          aria-label="Bold"
          onClick={() => run((chain) => chain.toggleBold().run())}
        >
          <Bold />
        </Button>
        <Button
          size="icon"
          variant={editor?.isActive("italic") ? "secondary" : "ghost"}
          aria-label="Italic"
          onClick={() => run((chain) => chain.toggleItalic().run())}
        >
          <Italic />
        </Button>
        <Button
          size="icon"
          variant={editor?.isActive("strike") ? "secondary" : "ghost"}
          aria-label="Strikethrough"
          onClick={() => run((chain) => chain.toggleStrike().run())}
        >
          <Strikethrough />
        </Button>
        <span className="mx-1 h-6 w-px bg-border" />
        <Button
          size="icon"
          variant="ghost"
          aria-label="Heading 1"
          onClick={() => run((chain) => chain.toggleHeading({ level: 1 }).run())}
        >
          <Heading1 />
        </Button>
        <Button
          size="icon"
          variant="ghost"
          aria-label="Heading 2"
          onClick={() => run((chain) => chain.toggleHeading({ level: 2 }).run())}
        >
          <Heading2 />
        </Button>
        <Button
          size="icon"
          variant="ghost"
          aria-label="Bulleted list"
          onClick={() => run((chain) => chain.toggleBulletList().run())}
        >
          <List />
        </Button>
        <Button
          size="icon"
          variant="ghost"
          aria-label="Numbered list"
          onClick={() => run((chain) => chain.toggleOrderedList().run())}
        >
          <ListOrdered />
        </Button>
        <Button
          size="sm"
          variant="ghost"
          onClick={() =>
            run((chain) => chain.insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run())
          }
        >
          <Table2 /> Table
        </Button>
        <div className="flex-1" />
        <span className="hidden text-xs text-muted-foreground sm:inline">
          Autosaved in this browser
        </span>
      </div>

      <div className="grid min-h-0 flex-1 grid-cols-1 xl:grid-cols-[minmax(0,1fr)_320px]">
        <main className="min-h-0 overflow-auto bg-[var(--canvas)] px-4 py-8 sm:px-8">
          <div className="mx-auto min-h-[900px] w-full max-w-[820px] bg-white px-8 py-12 text-[#202124] shadow-xl sm:px-16 sm:py-16">
            <EditorContent editor={editor} />
          </div>
        </main>
        <aside className="flex min-h-0 flex-col border-l border-border bg-card/60">
          <div className="flex items-center gap-2 border-b border-border px-4 py-3 text-sm font-semibold">
            <Sparkles className="h-4 w-4 text-primary" /> Writing assistant
          </div>
          <p className="px-4 pt-3 text-xs leading-relaxed text-muted-foreground">
            Runs privately in your browser. First use downloads the model once; then draft,
            rewrite, summarize, or organize your document.
          </p>
          <form onSubmit={applyAi} className="flex flex-1 flex-col gap-3 p-4">
            <Textarea
              value={aiPrompt}
              onChange={(event) => setAiPrompt(event.target.value)}
              placeholder="Draft a project proposal, improve this section, or turn the text into an executive summary..."
              className="min-h-32 resize-y"
            />
            <Button type="submit" disabled={aiBusy || !aiPrompt.trim()}>
              {aiBusy ? <span className="animate-spin">◌</span> : <Wand2 />}{" "}
              {aiBusy ? (aiStatus ? localAiStatusLabel(aiStatus) : "Preparing local AI…") : "Generate or revise"}
            </Button>
          </form>
        </aside>
      </div>
    </div>
  );
}
