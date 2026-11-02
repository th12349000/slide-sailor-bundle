import { z } from "zod";

const slideElementSchema = z.object({
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
  imageAttribution: z.string().default(""),
  imageSourceUrl: z.string().default(""),
  imageLicense: z.string().default(""),
  imageLicenseUrl: z.string().default(""),
  linkUrl: z.string().nullable(),
  linkSlideIndex: z.number().nullable(),
});

export const slideAiResultSchema = z.object({
  reply: z.string(),
  mode: z.enum(["none", "append", "replace_current", "replace_all"]),
  slides: z.array(
    z.object({
      background: z.string(),
      notes: z.string(),
      elements: z.array(slideElementSchema),
    }),
  ),
});

const conditionalFormatSchema = z.object({
  range: z.string(),
  operator: z.enum(["greaterThan", "lessThan", "equal", "containsText"]),
  value: z.string(),
  color: z.string(),
  background: z.string(),
});

const workbookSchema = z.object({
  title: z.string(),
  sheets: z
    .array(
      z.object({
        name: z.string(),
        cells: z.array(z.array(z.string())),
        conditionalFormats: z.array(conditionalFormatSchema),
        charts: z.array(
          z.object({ title: z.string(), range: z.string(), kind: z.enum(["bar", "line", "pie"]) }),
        ),
      }),
    )
    .min(1)
    .max(8),
});

const documentSchema = z.object({
  title: z.string(),
  blocks: z
    .array(
      z.object({
        type: z.enum(["heading1", "heading2", "paragraph", "bullet", "numbered"]),
        text: z.string(),
      }),
    )
    .max(80),
});

export const officeAiResultSchema = z.object({
  reply: z.string(),
  workbook: workbookSchema.nullable(),
  document: documentSchema.nullable(),
});

export type AiResult = z.infer<typeof slideAiResultSchema>;
export type OfficeAiResult = z.infer<typeof officeAiResultSchema>;
