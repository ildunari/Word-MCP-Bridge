import type { BridgeCapability } from "./protocol.js";
import { z } from "zod";

export type WordToolGroup = "read" | "write" | "review" | "advanced";

export interface WordToolContract {
  name: string;
  title: string;
  description: string;
  group: WordToolGroup;
  requiredCapability: BridgeCapability;
  firstClassMcp: boolean;
  readOnly?: boolean;
  destructive?: boolean;
  idempotent?: boolean;
  inputSchema: z.ZodTypeAny;
  parameters: Record<string, unknown>;
}

type WordToolArgRecord = Record<string, unknown>;

function schema(shape: Record<string, z.ZodTypeAny>) {
  return z.object(shape);
}

function parameters(
  properties: Record<string, Record<string, unknown>>,
  required: string[] = [],
) {
  return {
    type: "object",
    properties,
    ...(required.length > 0 ? { required } : {}),
    additionalProperties: false,
  };
}

function defineWordToolContract(contract: WordToolContract): WordToolContract {
  return contract;
}

function asArgRecord(value: unknown): WordToolArgRecord {
  return value && typeof value === "object" && !Array.isArray(value)
    ? { ...(value as WordToolArgRecord) }
    : {};
}

const paragraphRangeSchema = schema({
  paragraphIndex: z.number().int().min(0),
  startOffset: z.number().int().min(0),
  endOffset: z.number().int().min(0),
});

const formatTextRangeSchema = schema({
  target: z.enum(["selection", "paragraphRange"]).optional(),
  paragraphIndex: z.number().int().min(0).optional(),
  startOffset: z.number().int().min(0).optional(),
  endOffset: z.number().int().min(0).optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  highlightColor: z.string().min(1).nullable().optional(),
  fontColor: z.string().min(1).optional(),
}).superRefine((value, ctx) => {
  const target = value.target ?? "selection";
  if (target === "paragraphRange") {
    if (value.paragraphIndex == null) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "paragraphIndex is required when target is paragraphRange.",
      });
    }
    if (value.startOffset == null || value.endOffset == null) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message:
          "startOffset and endOffset are required when target is paragraphRange.",
      });
    }
  }
  if (
    value.bold == null &&
    value.italic == null &&
    value.underline == null &&
    value.highlightColor === undefined &&
    value.fontColor == null
  ) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "At least one inline formatting field must be provided.",
    });
  }
});

const revisionReferenceSchema = schema({
  revisionToken: z.string().min(1).optional(),
  revisionId: z.string().min(1).optional(),
}).superRefine((value, ctx) => {
  if (!value.revisionToken && !value.revisionId) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "revisionToken or revisionId is required.",
    });
  }
});

const searchAndReplaceSchema = schema({
  query: z.string().min(1),
  replacement: z.string(),
  matchCase: z.boolean().optional(),
  wholeWord: z.boolean().optional(),
  maxMatches: z.number().int().min(1).max(100).optional(),
  targetMatchIndexes: z.array(z.number().int().min(0)).optional(),
});

export const WORD_TOOL_CONTRACTS: WordToolContract[] = [
  defineWordToolContract({
    name: "word_get_document_text",
    title: "Word Get Document Text",
    description:
      "Read the active Word document as paragraphs with paragraph indices and style hints.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({
      startParagraph: z.number().int().min(0).optional(),
      endParagraph: z.number().int().min(0).optional(),
    }),
    parameters: parameters({
      startParagraph: {
        type: "integer",
        minimum: 0,
        description: "Optional first paragraph index to include.",
      },
      endParagraph: {
        type: "integer",
        minimum: 0,
        description: "Optional last paragraph index to include.",
      },
    }),
  }),
  defineWordToolContract({
    name: "word_get_text_range",
    title: "Word Get Text Range",
    description:
      "Read exact text from one paragraph using explicit character offsets. Use this before precise edits to verify the intended slice.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: paragraphRangeSchema,
    parameters: parameters(
      {
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description: "Target paragraph index.",
        },
        startOffset: {
          type: "integer",
          minimum: 0,
          description: "Inclusive start character offset inside the paragraph.",
        },
        endOffset: {
          type: "integer",
          minimum: 0,
          description: "Exclusive end character offset inside the paragraph.",
        },
      },
      ["paragraphIndex", "startOffset", "endOffset"],
    ),
  }),
  defineWordToolContract({
    name: "word_get_document_structure",
    title: "Word Get Document Structure",
    description:
      "Inspect headings, tables, content controls, and document counts for the active Word document.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
  defineWordToolContract({
    name: "word_get_selection",
    title: "Word Get Selection",
    description:
      "Read the current Word selection, including selected text and style when available.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
  defineWordToolContract({
    name: "word_get_paragraph_runs",
    title: "Word Get Paragraph Runs",
    description:
      "Read run-level formatting boundaries for a specific paragraph or, when paragraphIndex is omitted, for the paragraph containing the current selection.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({
      paragraphIndex: z.number().int().min(0).optional(),
    }),
    parameters: parameters({
      paragraphIndex: {
        type: "integer",
        minimum: 0,
        description:
          "Optional paragraph index. If omitted, the tool uses the current selection paragraph.",
      },
    }),
  }),
  defineWordToolContract({
    name: "word_get_tables",
    title: "Word Get Tables",
    description:
      "List table summaries from the active Word document, including dimensions and optional cell previews.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({
      includeCellPreview: z.boolean().optional(),
    }),
    parameters: parameters({
      includeCellPreview: {
        type: "boolean",
        description: "When true, include a lightweight text preview for each table.",
      },
    }),
  }),
  defineWordToolContract({
    name: "word_list_comments",
    title: "Word List Comments",
    description:
      "List native Word comments and replies from the active document.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({
      includeReplies: z.boolean().optional(),
    }),
    parameters: parameters({
      includeReplies: {
        type: "boolean",
        description: "When true, include replies for each comment.",
      },
    }),
  }),
  defineWordToolContract({
    name: "word_get_revisions",
    title: "Word Get Revisions",
    description:
      "List tracked revisions from the active Word document and return a stronger revisionToken for validated follow-up actions. revisionId values remain ephemeral debug hints only.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
  defineWordToolContract({
    name: "word_insert_paragraph",
    title: "Word Insert Paragraph",
    description:
      "Insert a paragraph at the end of the document or before/after a specific paragraph.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      text: z.string().min(1),
      location: z.enum(["end", "before", "after"]).optional(),
      paragraphIndex: z.number().int().min(0).optional(),
      style: z.string().min(1).optional(),
    }),
    parameters: parameters(
      {
        text: {
          type: "string",
          description: "Paragraph text to insert.",
        },
        location: {
          type: "string",
          enum: ["end", "before", "after"],
          description: "Where to insert the paragraph. Defaults to end.",
        },
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description:
            "Required when location is before or after; identifies the target paragraph.",
        },
        style: {
          type: "string",
          description: "Optional built-in Word style to apply to the new paragraph.",
        },
      },
      ["text"],
    ),
  }),
  defineWordToolContract({
    name: "word_insert_text",
    title: "Word Insert Text",
    description:
      "Insert text without replacing existing text. target=selection inserts at the end of the current selection, target=documentEnd appends inline at the document end, and paragraphStart/paragraphEnd insert at the start or end of one explicit paragraph.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      text: z.string().min(1),
      target: z
        .enum(["selection", "cursor", "documentEnd", "paragraphStart", "paragraphEnd"])
        .optional(),
      paragraphIndex: z.number().int().min(0).optional(),
    }),
    parameters: parameters(
      {
        text: {
          type: "string",
          description: "Text to insert.",
        },
        target: {
          type: "string",
          enum: [
            "selection",
            "cursor",
            "documentEnd",
            "paragraphStart",
            "paragraphEnd",
          ],
          description:
            "Insertion target. selection requires highlighted text, cursor uses the active caret, and defaults to selection.",
        },
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description:
            "Required when target is paragraphStart or paragraphEnd.",
        },
      },
      ["text"],
    ),
  }),
  defineWordToolContract({
    name: "word_replace_text_range",
    title: "Word Replace Text Range",
    description:
      "Replace exact text inside one paragraph using explicit character offsets. Optionally provide expectedText to fail fast if the paragraph changed since the last read.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: paragraphRangeSchema.extend({
      text: z.string(),
      expectedText: z.string().optional(),
    }),
    parameters: parameters(
      {
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description: "Target paragraph index.",
        },
        startOffset: {
          type: "integer",
          minimum: 0,
          description: "Inclusive start character offset inside the paragraph.",
        },
        endOffset: {
          type: "integer",
          minimum: 0,
          description: "Exclusive end character offset inside the paragraph.",
        },
        text: {
          type: "string",
          description: "Replacement text for the exact target range.",
        },
        expectedText: {
          type: "string",
          description:
            "Optional stale-check text. If provided, the current range text must match exactly before mutation.",
        },
      },
      ["paragraphIndex", "startOffset", "endOffset", "text"],
    ),
  }),
  defineWordToolContract({
    name: "word_replace_selection",
    title: "Word Replace Selection",
    description:
      "Replace only the current Word selection. Fails if no text is selected and returns refreshed local paragraph state after the mutation.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      text: z.string(),
    }),
    parameters: parameters(
      {
        text: {
          type: "string",
          description: "Replacement text for the current selection.",
        },
      },
      ["text"],
    ),
  }),
  defineWordToolContract({
    name: "word_search_text",
    title: "Word Search Text",
    description:
      "Search the active Word document and return candidate edit ranges with paragraph indices, offsets, and short context so the next mutation can target the right match safely.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: schema({
      query: z.string().min(1),
      matchCase: z.boolean().optional(),
      wholeWord: z.boolean().optional(),
      maxMatches: z.number().int().min(1).max(100).optional(),
    }),
    parameters: parameters(
      {
        query: {
          type: "string",
          description: "Text to search for.",
        },
        matchCase: {
          type: "boolean",
          description: "When true, match case exactly.",
        },
        wholeWord: {
          type: "boolean",
          description: "When true, match whole words only.",
        },
        maxMatches: {
          type: "integer",
          minimum: 1,
          maximum: 100,
          description: "Maximum number of matches to return.",
        },
      },
      ["query"],
    ),
  }),
  defineWordToolContract({
    name: "word_search_and_replace",
    title: "Word Search And Replace",
    description:
      "Replace exact text matches in the active Word document with explicit match limits. For repeated phrases, prefer word_search_text followed by word_replace_text_range, or provide targetMatchIndexes.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: searchAndReplaceSchema,
    parameters: parameters(
      {
        query: {
          type: "string",
          description: "Text to search for.",
        },
        replacement: {
          type: "string",
          description: "Replacement text.",
        },
        matchCase: {
          type: "boolean",
          description: "When true, match case exactly.",
        },
        wholeWord: {
          type: "boolean",
          description: "When true, match whole words only.",
        },
        maxMatches: {
          type: "integer",
          minimum: 1,
          maximum: 100,
          description: "Maximum number of matches to replace.",
        },
        targetMatchIndexes: {
          type: "array",
          items: {
            type: "integer",
            minimum: 0,
          },
          description:
            "Optional exact match indexes from word_search_text. Strongly recommended when the query appears more than once.",
        },
      },
      ["query", "replacement"],
    ),
  }),
  defineWordToolContract({
    name: "word_apply_style",
    title: "Word Apply Style",
    description:
      "Apply a built-in Word paragraph style to a specific paragraph.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: schema({
      paragraphIndex: z.number().int().min(0),
      style: z.string().min(1),
    }),
    parameters: parameters(
      {
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description: "Target paragraph index.",
        },
        style: {
          type: "string",
          description: "Built-in or existing paragraph style name.",
        },
      },
      ["paragraphIndex", "style"],
    ),
  }),
  defineWordToolContract({
    name: "word_set_paragraph_format",
    title: "Word Set Paragraph Format",
    description:
      "Update paragraph-level layout such as alignment, spacing, and line spacing for one paragraph.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: schema({
      paragraphIndex: z.number().int().min(0),
      alignment: z
        .enum(["Left", "Centered", "Right", "Justified"])
        .optional(),
      spaceBefore: z.number().min(0).optional(),
      spaceAfter: z.number().min(0).optional(),
      lineSpacing: z.number().positive().max(144).optional(),
    }),
    parameters: parameters(
      {
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description: "Target paragraph index.",
        },
        alignment: {
          type: "string",
          enum: ["Left", "Centered", "Right", "Justified"],
          description: "Paragraph alignment.",
        },
        spaceBefore: {
          type: "number",
          minimum: 0,
          description: "Space before the paragraph in points.",
        },
        spaceAfter: {
          type: "number",
          minimum: 0,
          description: "Space after the paragraph in points.",
        },
        lineSpacing: {
          type: "number",
          exclusiveMinimum: 0,
          maximum: 144,
          description:
            "Line spacing in points. Common values are 12 (single), 18 (1.5x), and 24 (double).",
        },
      },
      ["paragraphIndex"],
    ),
  }),
  defineWordToolContract({
    name: "word_format_text_range",
    title: "Word Format Text Range",
    description:
      "Apply a narrow inline formatting change either to the current selection or to an explicit paragraph text range.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: formatTextRangeSchema,
    parameters: parameters({
      target: {
        type: "string",
        enum: ["selection", "paragraphRange"],
        description:
          "Formatting target. Defaults to selection. Use paragraphRange for explicit offset-based formatting.",
      },
      paragraphIndex: {
        type: "integer",
        minimum: 0,
        description: "Required when target is paragraphRange.",
      },
      startOffset: {
        type: "integer",
        minimum: 0,
        description: "Inclusive start character offset when target is paragraphRange.",
      },
      endOffset: {
        type: "integer",
        minimum: 0,
        description: "Exclusive end character offset when target is paragraphRange.",
      },
      bold: {
        type: "boolean",
        description: "Set bold on or off.",
      },
      italic: {
        type: "boolean",
        description: "Set italic on or off.",
      },
      underline: {
        type: "boolean",
        description: "Set underline on or off.",
      },
      highlightColor: {
        type: ["string", "null"],
        description:
          "Set highlight color using a color name or #RRGGBB. Use null to clear highlight.",
      },
      fontColor: {
        type: "string",
        description: "Set font color using a color name or #RRGGBB.",
      },
    }),
  }),
  defineWordToolContract({
    name: "word_insert_table",
    title: "Word Insert Table",
    description:
      "Insert a table at the document end or before/after a specific paragraph.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      rows: z.number().int().min(1),
      columns: z.number().int().min(1),
      values: z.array(z.array(z.string())).optional(),
      location: z.enum(["end", "before", "after"]).optional(),
      paragraphIndex: z.number().int().min(0).optional(),
    }),
    parameters: parameters(
      {
        rows: {
          type: "integer",
          minimum: 1,
          description: "Number of rows to create.",
        },
        columns: {
          type: "integer",
          minimum: 1,
          description: "Number of columns to create.",
        },
        values: {
          type: "array",
          description: "Optional initial table data by row.",
        },
        location: {
          type: "string",
          enum: ["end", "before", "after"],
          description: "Where to insert the table. Defaults to end.",
        },
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description:
            "Required when location is before or after; identifies the target paragraph.",
        },
      },
      ["rows", "columns"],
    ),
  }),
  defineWordToolContract({
    name: "word_update_table_cell",
    title: "Word Update Table Cell",
    description:
      "Replace the contents of a specific table cell in the active Word document.",
    group: "write",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      tableIndex: z.number().int().min(0),
      rowIndex: z.number().int().min(0),
      columnIndex: z.number().int().min(0),
      text: z.string(),
    }),
    parameters: parameters(
      {
        tableIndex: {
          type: "integer",
          minimum: 0,
          description: "Target table index.",
        },
        rowIndex: {
          type: "integer",
          minimum: 0,
          description: "Target row index.",
        },
        columnIndex: {
          type: "integer",
          minimum: 0,
          description: "Target column index.",
        },
        text: {
          type: "string",
          description: "Replacement cell text.",
        },
      },
      ["tableIndex", "rowIndex", "columnIndex", "text"],
    ),
  }),
  defineWordToolContract({
    name: "word_insert_comment",
    title: "Word Insert Comment",
    description:
      "Insert a native Word comment on the current selection or a specific paragraph and return refreshed comment counts and target details.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      text: z.string().min(1),
      paragraphIndex: z.number().int().min(0).optional(),
    }),
    parameters: parameters(
      {
        text: {
          type: "string",
          description: "Comment text to insert.",
        },
        paragraphIndex: {
          type: "integer",
          minimum: 0,
          description:
            "Optional paragraph index when no selection should be used.",
        },
      },
      ["text"],
    ),
  }),
  defineWordToolContract({
    name: "word_reply_to_comment",
    title: "Word Reply To Comment",
    description:
      "Reply to an existing native Word comment by comment ID and return refreshed reply counts.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    inputSchema: schema({
      commentId: z.string().min(1),
      text: z.string().min(1),
    }),
    parameters: parameters(
      {
        commentId: {
          type: "string",
          description: "Target comment ID returned by word_list_comments.",
        },
        text: {
          type: "string",
          description: "Reply text.",
        },
      },
      ["commentId", "text"],
    ),
  }),
  defineWordToolContract({
    name: "word_resolve_comment",
    title: "Word Resolve Comment",
    description:
      "Mark a native Word comment as resolved by comment ID and return refreshed comment state.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: schema({
      commentId: z.string().min(1),
    }),
    parameters: parameters(
      {
        commentId: {
          type: "string",
          description: "Target comment ID returned by word_list_comments.",
        },
      },
      ["commentId"],
    ),
  }),
  defineWordToolContract({
    name: "word_delete_comment",
    title: "Word Delete Comment",
    description:
      "Delete one native Word comment or reply by comment ID and return refreshed remaining comment counts.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    destructive: true,
    inputSchema: schema({
      commentId: z.string().min(1),
    }),
    parameters: parameters(
      {
        commentId: {
          type: "string",
          description: "Target comment ID returned by word_list_comments.",
        },
      },
      ["commentId"],
    ),
  }),
  defineWordToolContract({
    name: "word_reopen_comment",
    title: "Word Reopen Comment",
    description:
      "Mark a resolved native Word comment as unresolved by comment ID and return refreshed comment state.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: schema({
      commentId: z.string().min(1),
    }),
    parameters: parameters(
      {
        commentId: {
          type: "string",
          description: "Target comment ID returned by word_list_comments.",
        },
      },
      ["commentId"],
    ),
  }),
  defineWordToolContract({
    name: "word_enable_track_changes",
    title: "Word Enable Track Changes",
    description:
      "Turn on track changes for the active Word document.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
  defineWordToolContract({
    name: "word_disable_track_changes",
    title: "Word Disable Track Changes",
    description:
      "Turn off track changes for the active Word document.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    idempotent: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
  defineWordToolContract({
    name: "word_get_revision_scope",
    title: "Word Get Revision Scope",
    description:
      "Read contextual location details for one tracked revision from word_get_revisions. Paragraph matching is best-effort because Word may not expose a stable paragraph anchor.",
    group: "read",
    requiredCapability: "tool_call",
    firstClassMcp: true,
    readOnly: true,
    inputSchema: revisionReferenceSchema,
    parameters: parameters(
      {
        revisionToken: {
          type: "string",
          description:
            "Preferred validated token returned by word_get_revisions for the current document state.",
        },
        revisionId: {
          type: "string",
          description:
            "Ephemeral revision ID returned by word_get_revisions for the current document state.",
        },
      },
      [],
    ),
  }),
  defineWordToolContract({
    name: "word_accept_revision",
    title: "Word Accept Revision",
    description:
      "Accept one tracked revision using a validated revisionToken from word_get_revisions. revisionId is supported only as a legacy fallback and is more brittle.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    destructive: true,
    inputSchema: revisionReferenceSchema,
    parameters: parameters(
      {
        revisionToken: {
          type: "string",
          description:
            "Preferred validated token returned by word_get_revisions for the current document state.",
        },
        revisionId: {
          type: "string",
          description:
            "Legacy ephemeral revision ID returned by word_get_revisions for the current document state.",
        },
      },
      [],
    ),
  }),
  defineWordToolContract({
    name: "word_reject_revision",
    title: "Word Reject Revision",
    description:
      "Reject one tracked revision using a validated revisionToken from word_get_revisions. revisionId is supported only as a legacy fallback and is more brittle.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    destructive: true,
    inputSchema: revisionReferenceSchema,
    parameters: parameters(
      {
        revisionToken: {
          type: "string",
          description:
            "Preferred validated token returned by word_get_revisions for the current document state.",
        },
        revisionId: {
          type: "string",
          description:
            "Legacy ephemeral revision ID returned by word_get_revisions for the current document state.",
        },
      },
      [],
    ),
  }),
  defineWordToolContract({
    name: "word_accept_all_revisions",
    title: "Word Accept All Revisions",
    description:
      "Accept all tracked revisions in the active Word document.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    destructive: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
  defineWordToolContract({
    name: "word_reject_all_revisions",
    title: "Word Reject All Revisions",
    description:
      "Reject all tracked revisions in the active Word document.",
    group: "review",
    requiredCapability: "document_edit",
    firstClassMcp: true,
    destructive: true,
    inputSchema: schema({}),
    parameters: parameters({}),
  }),
];

export const WORD_TOOL_CONTRACTS_BY_NAME = new Map(
  WORD_TOOL_CONTRACTS.map((contract) => [contract.name, contract]),
);

export function getWordToolContract(name: string): WordToolContract | undefined {
  return WORD_TOOL_CONTRACTS_BY_NAME.get(name);
}

export function getFirstClassWordToolContracts(): WordToolContract[] {
  return WORD_TOOL_CONTRACTS.filter((contract) => contract.firstClassMcp);
}

export function getWordToolInputShape(
  contract: WordToolContract,
): Record<string, z.ZodTypeAny> {
  const candidate = contract.inputSchema as z.ZodTypeAny & {
    shape?: Record<string, z.ZodTypeAny>;
  };
  return candidate.shape ? { ...candidate.shape } : {};
}

export function normalizeWordToolArgs(
  toolName: string,
  value: unknown,
): WordToolArgRecord {
  const args = asArgRecord(value);

  const copyStringAlias = (from: string, to: string) => {
    if (args[to] == null && typeof args[from] === "string" && args[from].trim()) {
      args[to] = args[from];
    }
  };

  if (toolName === "word_search_text" || toolName === "word_search_and_replace") {
    copyStringAlias("searchText", "query");
  }

  if (toolName === "word_search_and_replace") {
    copyStringAlias("replaceText", "replacement");
  }

  if (toolName === "word_replace_text_range") {
    copyStringAlias("replacement", "text");
  }

  if (toolName === "word_apply_style" || toolName === "word_insert_paragraph") {
    copyStringAlias("styleName", "style");
  }

  if (
    toolName === "word_reply_to_comment" ||
    toolName === "word_resolve_comment" ||
    toolName === "word_delete_comment" ||
    toolName === "word_reopen_comment"
  ) {
    if (args.commentId != null && typeof args.commentId === "number") {
      args.commentId = String(args.commentId);
    }
  }

  if (toolName === "word_format_text_range") {
    copyStringAlias("color", "fontColor");
  }

  return args;
}
