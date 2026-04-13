import {
  normalizeWordToolArgs,
  WORD_TOOL_CONTRACTS,
  type WordToolContract,
} from "@word-mcp-bridge/bridge/word-tool-contracts";

declare const Word: any;

type ToolContentPart = { type: "text"; text: string };

interface ExecutableWordTool {
  name: string;
  label?: string;
  description?: string;
  parameters?: unknown;
  requiredCapability?: "tool_call" | "document_edit";
  execute: (_toolCallId: string, args: unknown) => Promise<unknown>;
}

type WordToolExecutor = (args: Record<string, unknown>) => Promise<unknown>;

function toolSuccess(summary: string, structuredContent: Record<string, unknown>) {
  return {
    content: [{ type: "text", text: summary } satisfies ToolContentPart],
    structuredContent: {
      success: true,
      ...structuredContent,
    },
  };
}

function toolError(message: string, extra: Record<string, unknown> = {}) {
  return {
    success: false,
    error: message,
    ...extra,
  };
}

async function runWordTool<T>(work: (context: any) => Promise<T>) {
  return Word.run(async (context: any) => work(context));
}

function asObject(args: unknown): Record<string, unknown> {
  return args && typeof args === "object" && !Array.isArray(args)
    ? (args as Record<string, unknown>)
    : {};
}

const FORMAT_TEXT_RANGE_ALLOWED_KEYS = new Set([
  "target",
  "paragraphIndex",
  "startOffset",
  "endOffset",
  "bold",
  "italic",
  "underline",
  "highlightColor",
  "fontColor",
]);

function unsupportedFormatTextRangeKeys(args: Record<string, unknown>) {
  return Object.keys(args).filter(
    (key) => !FORMAT_TEXT_RANGE_ALLOWED_KEYS.has(key) && key !== "color",
  );
}

function summarizeWordFailure(error: unknown, fallback: string) {
  if (error instanceof Error && error.message.trim()) {
    return error.message.trim();
  }
  return fallback;
}

function readParagraphStyle(paragraph: any): string | null {
  try {
    return typeof paragraph.style === "string" ? paragraph.style : null;
  } catch {
    return null;
  }
}

function getSafeString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0
    ? value.trim()
    : undefined;
}

function getSafeNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function getRunText(node: Element, namespaceUri: string): string {
  return Array.from(node.getElementsByTagNameNS(namespaceUri, "t"))
    .map((child) => child.textContent ?? "")
    .join("");
}

function getRunFormatting(node: Element, namespaceUri: string) {
  const properties = node.getElementsByTagNameNS(namespaceUri, "rPr")[0];
  if (!properties) {
    return {};
  }
  const has = (name: string) =>
    properties.getElementsByTagNameNS(namespaceUri, name).length > 0;
  const fonts = properties.getElementsByTagNameNS(namespaceUri, "rFonts")[0];
  const size = properties.getElementsByTagNameNS(namespaceUri, "sz")[0];
  const color = properties.getElementsByTagNameNS(namespaceUri, "color")[0];

  return {
    bold: has("b"),
    italic: has("i"),
    underline: has("u"),
    fontName:
      fonts?.getAttributeNS(namespaceUri, "ascii") ??
      fonts?.getAttribute("w:ascii") ??
      null,
    fontSizeHalfPoints:
      size?.getAttributeNS(namespaceUri, "val") ??
      size?.getAttribute("w:val") ??
      null,
    color:
      color?.getAttributeNS(namespaceUri, "val") ??
      color?.getAttribute("w:val") ??
      null,
  };
}

function extractRunsFromOoxml(ooxml: string) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(ooxml, "application/xml");
  const namespaceUri =
    xml.documentElement.lookupNamespaceURI("w") ??
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const runs = Array.from(xml.getElementsByTagNameNS(namespaceUri, "r"));
  let cursor = 0;

  return runs
    .map((run, index) => {
      const text = getRunText(run, namespaceUri);
      const start = cursor;
      cursor += text.length;
      return {
        runIndex: index,
        text,
        start,
        end: cursor,
        formatting: getRunFormatting(run, namespaceUri),
      };
    })
    .filter((run) => run.text.length > 0);
}

async function getParagraphs(context: any) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  for (const paragraph of paragraphs.items) {
    paragraph.load("text,style");
  }
  await context.sync();
  return paragraphs.items as any[];
}

async function getParagraphByIndex(context: any, paragraphIndex: number) {
  const paragraphs = await getParagraphs(context);
  const paragraph = paragraphs[paragraphIndex];
  if (!paragraph) {
    throw new Error(`Paragraph ${paragraphIndex} does not exist.`);
  }
  return { paragraph, paragraphs };
}

async function readParagraphSnapshot(
  context: any,
  paragraphIndex: number,
  paragraph?: any,
) {
  const target = paragraph ?? (await getParagraphByIndex(context, paragraphIndex)).paragraph;
  target.load("text,style");
  await context.sync();
  return {
    paragraphIndex,
    text: target.text ?? "",
    style: readParagraphStyle(target),
    summary: summarizeParagraph(target.text ?? ""),
  };
}

async function tryReadParagraphSnapshot(
  context: any,
  paragraphIndex: number,
  paragraph?: any,
) {
  try {
    return await readParagraphSnapshot(context, paragraphIndex, paragraph);
  } catch {
    return null;
  }
}

function summarizeParagraph(text: string) {
  const trimmed = text.trim();
  if (!trimmed) return "(empty paragraph)";
  return trimmed.length > 80 ? `${trimmed.slice(0, 77)}...` : trimmed;
}

function matchingParagraphIndexes(
  allParagraphs: any[],
  paragraphText: string,
  paragraphStyle: string,
) {
  return allParagraphs.flatMap((candidate, index) =>
    (candidate.text ?? "") === paragraphText &&
    readParagraphStyle(candidate) === paragraphStyle
      ? [index]
      : [],
  );
}

function normalizeOffsets(
  paragraphText: string,
  startOffset: number,
  endOffset: number,
) {
  if (!Number.isInteger(startOffset) || !Number.isInteger(endOffset)) {
    throw new Error("startOffset and endOffset must be integers.");
  }
  if (startOffset < 0 || endOffset < 0) {
    throw new Error("startOffset and endOffset must be at least 0.");
  }
  if (endOffset < startOffset) {
    throw new Error("endOffset must be greater than or equal to startOffset.");
  }
  if (endOffset > paragraphText.length) {
    throw new Error(
      `Requested range ${startOffset}-${endOffset} is outside paragraph length ${paragraphText.length}.`,
    );
  }
  return {
    startOffset,
    endOffset,
    text: paragraphText.slice(startOffset, endOffset),
  };
}

function countOccurrencesBefore(text: string, needle: string, beforeOffset: number) {
  if (!needle) return 0;
  let count = 0;
  let cursor = text.indexOf(needle);
  while (cursor >= 0 && cursor < beforeOffset) {
    count += 1;
    cursor = text.indexOf(needle, cursor + needle.length);
  }
  return count;
}

async function resolveParagraphTextRange(
  context: any,
  args: Record<string, unknown>,
  options: { requireNonEmpty?: boolean; allowExpectedText?: boolean } = {},
) {
  const paragraphIndex = Math.trunc(getSafeNumber(args.paragraphIndex) ?? -1);
  const startOffset = Math.trunc(getSafeNumber(args.startOffset) ?? -1);
  const endOffset = Math.trunc(getSafeNumber(args.endOffset) ?? -1);
  if (paragraphIndex < 0 || startOffset < 0 || endOffset < 0) {
    throw new Error("paragraphIndex, startOffset, and endOffset are required.");
  }
  const { paragraph } = await getParagraphByIndex(context, paragraphIndex);
  paragraph.load("text,style");
  await context.sync();
  const paragraphText = paragraph.text ?? "";
  const normalized = normalizeOffsets(paragraphText, startOffset, endOffset);
  if (options.requireNonEmpty && normalized.startOffset === normalized.endOffset) {
    throw new Error("The requested range must contain at least one character.");
  }
  const expectedText =
    options.allowExpectedText && typeof args.expectedText === "string"
      ? args.expectedText
      : undefined;
  if (expectedText !== undefined && normalized.text !== expectedText) {
    throw new Error(
      `The requested range is stale. Expected "${expectedText}" but found "${normalized.text}".`,
    );
  }
  return {
    paragraph,
    paragraphIndex,
    paragraphText,
    paragraphStyle: readParagraphStyle(paragraph),
    startOffset: normalized.startOffset,
    endOffset: normalized.endOffset,
    rangeText: normalized.text,
  };
}

async function locateParagraphRange(
  context: any,
  paragraph: any,
  paragraphText: string,
  startOffset: number,
  endOffset: number,
) {
  const targetText = paragraphText.slice(startOffset, endOffset);
  if (!targetText) {
    throw new Error("The requested range must contain at least one character.");
  }
  const occurrenceIndex = countOccurrencesBefore(paragraphText, targetText, startOffset);
  const matches = paragraph.search(targetText, {
    matchCase: true,
    matchWholeWord: false,
  });
  matches.load("items");
  await context.sync();
  const range = matches.items[occurrenceIndex];
  if (!range) {
    throw new Error(
      "Could not resolve the requested text range in Word. Re-read the paragraph and try again.",
    );
  }
  range.load("text");
  await context.sync();
  if ((range.text ?? "") !== targetText) {
    throw new Error(
      "Resolved range text does not match the requested slice. Re-read the paragraph and try again.",
    );
  }
  return range;
}

function normalizeMaxMatches(value: unknown, fallback = 20) {
  const next = getSafeNumber(value);
  if (next == null) return fallback;
  return Math.max(1, Math.min(100, Math.trunc(next)));
}

function resolveSearchOptions(args: Record<string, unknown>) {
  return {
    matchCase: Boolean(args.matchCase),
    matchWholeWord: Boolean(args.wholeWord),
  };
}

async function resolveComments(context: any) {
  const comments = context.document.body.getComments();
  comments.load("items");
  await context.sync();
  for (const comment of comments.items) {
    comment.load("id,content,resolved,authorName,creationDate");
    comment.replies.load("items");
  }
  await context.sync();
  return comments.items as any[];
}

export async function resolveTableDimensions(context: any, table: any) {
  table.load("rowCount,columnCount");
  const rows = table.rows;
  rows.load("items");
  await context.sync();

  const rowCountFromRows = Array.isArray(rows.items) ? rows.items.length : 0;
  const rowCount =
    typeof table.rowCount === "number" && table.rowCount > 0
      ? table.rowCount
      : rowCountFromRows;

  let columnCount =
    typeof table.columnCount === "number" && table.columnCount > 0
      ? table.columnCount
      : 0;

  if (columnCount <= 0 && rowCountFromRows > 0 && rows.items[0]) {
    const firstRow = rows.items[0];
    const cells = firstRow.cells;
    cells.load("items");
    await context.sync();
    columnCount = Array.isArray(cells.items) ? cells.items.length : 0;
  }

  return {
    rowCount,
    columnCount,
  };
}

async function findCommentById(context: any, commentId: string) {
  const comments = await resolveComments(context);
  const comment = comments.find((entry) => String(entry.id) === commentId);
  if (!comment) {
    throw new Error(`Comment ${commentId} was not found.`);
  }
  return comment;
}

async function resolveTrackedChanges(context: any) {
  const changes = context.document.body.getTrackedChanges();
  changes.load("items");
  await context.sync();
  for (const [index, change] of changes.items.entries()) {
    change.load("author,date,text,type");
    (change as any).__bridgeRevisionId = String(index);
  }
  await context.sync();
  return changes.items as any[];
}

function encodeRevisionToken(payload: Record<string, unknown>) {
  const json = JSON.stringify(payload);
  const bytes = new TextEncoder().encode(json);
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return `word-revision:${btoa(binary)}`;
}

function decodeRevisionToken(token: string): Record<string, unknown> {
  if (!token.startsWith("word-revision:")) {
    throw new Error("revisionToken is not in the expected format.");
  }
  try {
    const binary = atob(token.slice("word-revision:".length));
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return JSON.parse(new TextDecoder().decode(bytes)) as Record<string, unknown>;
  } catch {
    throw new Error("revisionToken could not be decoded.");
  }
}

function normalizeRevisionScalar(value: unknown): string | null {
  if (value == null) return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "string") return value;
  return String(value);
}

function summarizeContext(text: string, startOffset: number, endOffset: number, radius = 24) {
  const before = text.slice(Math.max(0, startOffset - radius), startOffset);
  const match = text.slice(startOffset, endOffset);
  const after = text.slice(endOffset, Math.min(text.length, endOffset + radius));
  return {
    before,
    match,
    after,
    preview: `${before}[${match}]${after}`,
  };
}

function isWordChar(value: string | undefined) {
  return value != null && /[\p{L}\p{N}_]/u.test(value);
}

function isWholeWordBoundary(text: string, start: number, end: number) {
  return !isWordChar(text[start - 1]) && !isWordChar(text[end]);
}

function collectParagraphSearchMatches(
  paragraphText: string,
  query: string,
  options: { matchCase: boolean; wholeWord: boolean },
) {
  const haystack = options.matchCase ? paragraphText : paragraphText.toLocaleLowerCase();
  const needle = options.matchCase ? query : query.toLocaleLowerCase();
  const matches: { startOffset: number; endOffset: number; text: string }[] = [];
  let cursor = haystack.indexOf(needle);
  while (cursor >= 0) {
    const endOffset = cursor + needle.length;
    if (!options.wholeWord || isWholeWordBoundary(paragraphText, cursor, endOffset)) {
      matches.push({
        startOffset: cursor,
        endOffset,
        text: paragraphText.slice(cursor, endOffset),
      });
    }
    cursor = haystack.indexOf(needle, cursor + Math.max(needle.length, 1));
  }
  return matches;
}

function findSearchCandidates(
  paragraphs: any[],
  query: string,
  options: { matchCase: boolean; wholeWord: boolean },
  maxMatches: number,
) {
  const matches: Array<{
    matchIndex: number;
    paragraphIndex: number;
    startOffset: number;
    endOffset: number;
    text: string;
    contextPreview: string;
    paragraphSummary: string;
  }> = [];
  let totalMatches = 0;

  paragraphs.forEach((paragraph, paragraphIndex) => {
    const paragraphText = paragraph.text ?? "";
    const paragraphMatches = collectParagraphSearchMatches(paragraphText, query, options);
    for (const paragraphMatch of paragraphMatches) {
      const matchIndex = totalMatches;
      totalMatches += 1;
      if (matches.length >= maxMatches) {
        continue;
      }
      matches.push({
        matchIndex,
        paragraphIndex,
        startOffset: paragraphMatch.startOffset,
        endOffset: paragraphMatch.endOffset,
        text: paragraphMatch.text,
        contextPreview: summarizeContext(
          paragraphText,
          paragraphMatch.startOffset,
          paragraphMatch.endOffset,
        ).preview,
        paragraphSummary: summarizeParagraph(paragraphText),
      });
    }
  });

  return { matches, totalMatches };
}

async function buildRevisionDescriptor(
  context: any,
  revision: any,
  revisionId: string,
  allParagraphs?: any[],
) {
  const range = revision.getRange("Whole");
  range.load("text");
  const paragraphs = range.paragraphs;
  paragraphs.load("items");
  await context.sync();

  let paragraphSummary: string | null = null;
  let paragraphIndex: number | null = null;
  if (paragraphs.items[0]) {
    paragraphs.items[0].load("text,style");
    await context.sync();
    const paragraphText = paragraphs.items[0].text ?? "";
    const paragraphStyle = readParagraphStyle(paragraphs.items[0]);
    paragraphSummary = summarizeParagraph(paragraphText);
    const paragraphPool = allParagraphs ?? (await getParagraphs(context));
    const matches = matchingParagraphIndexes(
      paragraphPool,
      paragraphText,
      paragraphStyle,
    );
    paragraphIndex = matches.length === 1 ? matches[0] : null;
  }

  const fingerprint = {
    author: normalizeRevisionScalar(revision.author),
    date: normalizeRevisionScalar(revision.date),
    text: normalizeRevisionScalar(revision.text) ?? "",
    type: normalizeRevisionScalar(revision.type),
    scopeText: range.text ?? "",
    paragraphSummary,
  };

  return {
    revisionId,
    revisionToken: encodeRevisionToken(fingerprint),
    ephemeralId: true,
    author: fingerprint.author,
    date: fingerprint.date,
    text: fingerprint.text,
    type: fingerprint.type,
    scopeText: fingerprint.scopeText,
    paragraphSummary,
    paragraphIndex,
  };
}

async function findTrackedChangeByReference(context: any, args: Record<string, unknown>) {
  const revisionToken = getSafeString(args.revisionToken);
  const revisionId = getSafeString(args.revisionId);

  if (revisionToken) {
    const expected = decodeRevisionToken(revisionToken);
    const changes = await resolveTrackedChanges(context);
    const allParagraphs = await getParagraphs(context);
    const candidates: Array<{
      revision: any;
      reference: Awaited<ReturnType<typeof buildRevisionDescriptor>>;
    }> = [];
    for (const change of changes) {
      const descriptor = await buildRevisionDescriptor(
        context,
        change,
        String((change as any).__bridgeRevisionId),
        allParagraphs,
      );
      if (
        descriptor.author === (expected.author ?? null) &&
        descriptor.date === (expected.date ?? null) &&
        descriptor.text === (expected.text ?? "") &&
        descriptor.type === (expected.type ?? null) &&
        descriptor.scopeText === (expected.scopeText ?? "") &&
        descriptor.paragraphSummary === (expected.paragraphSummary ?? null)
      ) {
        candidates.push({
          revision: change,
          reference: descriptor,
        });
      }
    }
    if (candidates.length === 1) {
      return {
        revision: candidates[0].revision,
        reference: {
          ...candidates[0].reference,
          revisionToken,
        },
        matchedBy: "revisionToken" as const,
      };
    }
    if (candidates.length > 1) {
      throw new Error(
        "The requested revisionToken matches more than one current revision. Re-read revisions and choose a more specific target before mutating.",
      );
    }
    throw new Error(
      "The requested revisionToken is stale or no longer matches the current revision set. Re-read revisions before mutating.",
    );
  }

  if (!revisionId) {
    throw new Error("revisionToken or revisionId is required.");
  }
  const revision = await findTrackedChangeById(context, revisionId);
  const descriptor = await buildRevisionDescriptor(
    context,
    revision,
    revisionId,
    await getParagraphs(context),
  );
  return {
    revision,
    reference: descriptor,
    matchedBy: "revisionId" as const,
  };
}

async function findTrackedChangeById(context: any, revisionId: string) {
  const changes = await resolveTrackedChanges(context);
  const revision = changes.find(
    (entry) => String((entry as any).__bridgeRevisionId) === revisionId,
  );
  if (!revision) {
    throw new Error(`Revision ${revisionId} was not found.`);
  }
  return revision;
}

function getSelectionText(selection: any) {
  return typeof selection.text === "string" ? selection.text : "";
}

const WORD_TOOL_EXECUTORS: Record<string, WordToolExecutor> = {
  async word_get_document_text(args) {
    return runWordTool(async (context) => {
      const paragraphs = await getParagraphs(context);
      const startParagraph = Math.max(0, Math.trunc(getSafeNumber(args.startParagraph) ?? 0));
      const endParagraph = Math.min(
        paragraphs.length - 1,
        Math.trunc(getSafeNumber(args.endParagraph) ?? (paragraphs.length - 1)),
      );

      if (paragraphs.length === 0) {
        return toolSuccess("Read 0 paragraphs.", { paragraphCount: 0, paragraphs: [] });
      }
      if (endParagraph < startParagraph) {
        return toolError("endParagraph must be greater than or equal to startParagraph.");
      }

      const result = paragraphs.slice(startParagraph, endParagraph + 1).map((paragraph, offset) => ({
        paragraphIndex: startParagraph + offset,
        text: paragraph.text ?? "",
        style: readParagraphStyle(paragraph),
      }));

      return toolSuccess(`Read ${result.length} paragraph${result.length === 1 ? "" : "s"}.`, {
        paragraphCount: paragraphs.length,
        paragraphs: result,
      });
    });
  },

  async word_get_text_range(args) {
    return runWordTool(async (context) => {
      const resolved = await resolveParagraphTextRange(context, args);
      return toolSuccess(
        `Read ${resolved.rangeText.length} character${resolved.rangeText.length === 1 ? "" : "s"} from paragraph ${resolved.paragraphIndex}.`,
        {
          paragraphIndex: resolved.paragraphIndex,
          startOffset: resolved.startOffset,
          endOffset: resolved.endOffset,
          text: resolved.rangeText,
          paragraph: {
            paragraphIndex: resolved.paragraphIndex,
            text: resolved.paragraphText,
            style: resolved.paragraphStyle,
            summary: summarizeParagraph(resolved.paragraphText),
          },
        },
      );
    });
  },

  async word_get_document_structure() {
    return runWordTool(async (context) => {
      const paragraphs = await getParagraphs(context);
      const sections = context.document.sections;
      const tables = context.document.body.tables;
      const contentControls = context.document.body.contentControls;
      sections.load("items");
      tables.load("items");
      contentControls.load("items");
      await context.sync();
      for (const control of contentControls.items) {
        control.load("title,tag,type");
      }
      await context.sync();

      const headings = paragraphs
        .map((paragraph, paragraphIndex) => ({
          paragraphIndex,
          text: paragraph.text ?? "",
          style: readParagraphStyle(paragraph),
        }))
        .filter((paragraph) =>
          typeof paragraph.style === "string" &&
          paragraph.style.toLowerCase().startsWith("heading"),
        );

      const tableSummaries = [];
      for (const [tableIndex, table] of tables.items.entries()) {
        const dimensions = await resolveTableDimensions(context, table);
        tableSummaries.push({
          tableIndex,
          rowCount: dimensions.rowCount,
          columnCount: dimensions.columnCount,
        });
      }

      return toolSuccess("Read document structure.", {
        paragraphCount: paragraphs.length,
        sectionCount: sections.items.length,
        headings,
        tables: tableSummaries,
        contentControls: contentControls.items.map((control: any, controlIndex: number) => ({
          controlIndex,
          title: control.title ?? null,
          tag: control.tag ?? null,
          type: control.type ?? null,
        })),
      });
    });
  },

  async word_get_selection() {
    return runWordTool(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text,style");
      await context.sync();
      const selectedText = getSelectionText(selection);
      return toolSuccess(
        selectedText
          ? `Read the current selection (${selectedText.length} characters).`
          : "There is no active text selection.",
        {
          hasSelection: selectedText.length > 0,
          selectedText,
          selectedStyle: selection.style ?? null,
        },
      );
    });
  },

  async word_get_paragraph_runs(args) {
    return runWordTool(async (context) => {
      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      let paragraph: any;
      let resolvedParagraphIndex = paragraphIndex ?? null;

      if (paragraphIndex == null) {
        const selection = context.document.getSelection();
        const selectionParagraphs = selection.paragraphs;
        selectionParagraphs.load("items");
        await context.sync();
        paragraph = selectionParagraphs.items[0];
        if (!paragraph) {
          return toolError("Could not resolve the current selection paragraph.");
        }
        paragraph.load("text,style");
        await context.sync();
        const allParagraphs = await getParagraphs(context);
        const matchingIndices = allParagraphs.flatMap((candidate, index) =>
          (candidate.text ?? "") === (paragraph.text ?? "") &&
          readParagraphStyle(candidate) === readParagraphStyle(paragraph)
            ? [index]
            : [],
        );
        const matchedIndex =
          matchingIndices.length === 1 ? matchingIndices[0] : -1;
        resolvedParagraphIndex = matchedIndex >= 0 ? matchedIndex : null;
      } else {
        ({ paragraph } = await getParagraphByIndex(context, paragraphIndex));
      }
      const ooxml = paragraph.getRange().getOoxml();
      await context.sync();
      const runs = extractRunsFromOoxml(ooxml.value ?? "");
      return toolSuccess(
        `Read ${runs.length} run${runs.length === 1 ? "" : "s"} from ${resolvedParagraphIndex != null ? `paragraph ${resolvedParagraphIndex}` : "the current selection paragraph"}.`,
        {
          paragraphIndex: resolvedParagraphIndex,
          runs,
        },
      );
    });
  },

  async word_get_tables(args) {
    return runWordTool(async (context) => {
      const includeCellPreview = Boolean(args.includeCellPreview);
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      const summaries = [];
      for (const [tableIndex, table] of tables.items.entries()) {
        const dimensions = await resolveTableDimensions(context, table);
        let preview: string[][] | null = null;
        if (includeCellPreview) {
          const rowsToRead = Math.min(2, dimensions.rowCount);
          const colsToRead = Math.min(3, dimensions.columnCount);
          preview = [];
          for (let rowIndex = 0; rowIndex < rowsToRead; rowIndex += 1) {
            const rowValues: string[] = [];
            for (let columnIndex = 0; columnIndex < colsToRead; columnIndex += 1) {
              const cell = table.getCell(rowIndex, columnIndex);
              cell.body.load("text");
              await context.sync();
              rowValues.push(cell.body.text ?? "");
            }
            preview.push(rowValues);
          }
        }
        summaries.push({
          tableIndex,
          rowCount: dimensions.rowCount,
          columnCount: dimensions.columnCount,
          preview,
        });
      }

      return toolSuccess(
        `Read ${summaries.length} table${summaries.length === 1 ? "" : "s"}.`,
        { tables: summaries },
      );
    });
  },

  async word_list_comments(args) {
    return runWordTool(async (context) => {
      const includeReplies = Boolean(args.includeReplies);
      const comments = await resolveComments(context);
      const result = [];

      for (const comment of comments) {
        const replies = [];
        if (includeReplies) {
          for (const reply of comment.replies.items) {
            reply.load("content,authorName,creationDate");
          }
          await context.sync();
          for (const reply of comment.replies.items) {
            replies.push({
              authorName: reply.authorName ?? null,
              content: reply.content ?? "",
              createdDate: reply.creationDate ?? null,
            });
          }
        }
        result.push({
          commentId: String(comment.id),
          authorName: comment.authorName ?? null,
          content: comment.content ?? "",
          createdDate: comment.creationDate ?? null,
          resolved: Boolean(comment.resolved),
          replies,
        });
      }

      return toolSuccess(
        `Read ${result.length} comment${result.length === 1 ? "" : "s"}.`,
        { comments: result },
      );
    });
  },

  async word_get_revisions() {
    return runWordTool(async (context) => {
      const changes = await resolveTrackedChanges(context);
      const allParagraphs = await getParagraphs(context);
      const revisions = [];
      for (const [index, change] of changes.entries()) {
        revisions.push(
          await buildRevisionDescriptor(
            context,
            change,
            String((change as any).__bridgeRevisionId ?? index),
            allParagraphs,
          ),
        );
      }

      return toolSuccess(
        `Read ${revisions.length} revision${revisions.length === 1 ? "" : "s"}.`,
        {
          revisions,
          note: "Use revisionToken for follow-up actions. revisionId remains an ephemeral debug hint and should not be trusted after intervening edits.",
        },
      );
    });
  },

  async word_get_revision_scope(args) {
    return runWordTool(async (context) => {
      const { revision, reference } = await findTrackedChangeByReference(context, args);
      const revisionId = reference.revisionId;
      const range = revision.getRange("Whole");
      range.load("text");
      const paragraphs = range.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let paragraphSnapshot: Awaited<ReturnType<typeof readParagraphSnapshot>> | null = null;
      if (paragraphs.items[0]) {
        paragraphs.items[0].load("text,style");
        await context.sync();
        const allParagraphs = await getParagraphs(context);
        const matchingIndices = matchingParagraphIndexes(
          allParagraphs,
          paragraphs.items[0].text ?? "",
          readParagraphStyle(paragraphs.items[0]),
        );
        if (matchingIndices.length === 1) {
          paragraphSnapshot = await readParagraphSnapshot(
            context,
            matchingIndices[0],
            allParagraphs[matchingIndices[0]],
          );
        }
      }

      return toolSuccess(`Read scope details for revision ${revisionId}.`, {
        revisionId,
        revisionToken: reference.revisionToken,
        ephemeralId: true,
        revision: {
          author: reference.author,
          date: reference.date,
          text: reference.text,
          type: reference.type,
        },
        scope: {
          text: range.text ?? "",
          paragraph: paragraphSnapshot,
          paragraphMatchIsBestEffort: true,
        },
        note:
          "Revision tokens are safer than revisionId for follow-up actions. Paragraph matching is best-effort because Word does not always expose a stable paragraph anchor for tracked changes.",
      });
    });
  },

  async word_insert_paragraph(args) {
    return runWordTool(async (context) => {
      const text = getSafeString(args.text);
      if (!text) {
        return toolError("text is required.");
      }
      const location = getSafeString(args.location) ?? "end";
      const style = getSafeString(args.style);

      if (location === "end") {
        try {
          const paragraph = context.document.body.insertParagraph(text, "End");
          if (style) paragraph.style = style;
          await context.sync();
          const paragraphs = await getParagraphs(context);
          const insertedIndex = Math.max(0, paragraphs.length - 1);
          return toolSuccess("Inserted a paragraph at the end of the document.", {
            location,
            text,
            style: style ?? null,
            paragraph: await readParagraphSnapshot(
              context,
              insertedIndex,
              paragraphs[insertedIndex],
            ),
          });
        } catch (error) {
          return toolError(
            style
              ? `Inserted paragraph text could not be styled with "${style}". Make sure the style exists in this document.`
              : "Word could not insert the requested paragraph.",
            {
              cause: summarizeWordFailure(error, "Word rejected the paragraph insertion."),
              location,
              style: style ?? null,
            },
          );
        }
      }

      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      if (paragraphIndex == null) {
        return toolError("paragraphIndex is required when location is before or after.");
      }

      try {
        const { paragraph } = await getParagraphByIndex(context, paragraphIndex);
        const inserted = paragraph.insertParagraph(
          text,
          location === "before" ? "Before" : "After",
        );
        if (style) inserted.style = style;
        await context.sync();
        const insertedIndex = paragraphIndex + (location === "before" ? 0 : 1);
        return toolSuccess(`Inserted a paragraph ${location} paragraph ${paragraphIndex}.`, {
          location,
          paragraphIndex,
          text,
          style: style ?? null,
          paragraph: await readParagraphSnapshot(context, insertedIndex, inserted),
        });
      } catch (error) {
        return toolError(
          style
            ? `Inserted paragraph text could not be styled with "${style}". Make sure the style exists in this document.`
            : `Word could not insert a paragraph ${location} paragraph ${paragraphIndex}.`,
          {
            cause: summarizeWordFailure(error, "Word rejected the paragraph insertion."),
            location,
            paragraphIndex,
            style: style ?? null,
          },
        );
      }
    });
  },

  async word_insert_text(args) {
    return runWordTool(async (context) => {
      const text = getSafeString(args.text);
      if (!text) {
        return toolError("text is required.");
      }
      const target = getSafeString(args.target) ?? "selection";
      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      let affectedParagraph: Record<string, unknown> | null = null;

      if (target === "documentEnd") {
        context.document.body.insertText(text, "End");
        await context.sync();
        const paragraphs = await getParagraphs(context);
        if (paragraphs.length > 0) {
          affectedParagraph = await readParagraphSnapshot(
            context,
            paragraphs.length - 1,
            paragraphs[paragraphs.length - 1],
          );
        }
      } else if (target === "selection" || target === "cursor") {
        const selection = context.document.getSelection();
        selection.load("text");
        const selectionParagraphs = selection.paragraphs;
        selectionParagraphs.load("items");
        await context.sync();
        if (target === "selection" && !getSelectionText(selection)) {
          return toolError(
            "There is no active text selection to insert into. Use target=\"cursor\" to insert at the caret.",
          );
        }
        selection.insertText(text, "End");
        await context.sync();
        if (selectionParagraphs.items[0]) {
          selectionParagraphs.items[0].load("text,style");
          await context.sync();
          affectedParagraph = {
            paragraphIndex: null,
            text: selectionParagraphs.items[0].text ?? "",
            style: readParagraphStyle(selectionParagraphs.items[0]),
            summary: summarizeParagraph(selectionParagraphs.items[0].text ?? ""),
          };
        }
      } else {
        if (paragraphIndex == null) {
          return toolError("paragraphIndex is required for paragraphStart and paragraphEnd.");
        }
        const { paragraph } = await getParagraphByIndex(context, paragraphIndex);
        paragraph.insertText(text, target === "paragraphStart" ? "Start" : "End");
        await context.sync();
        affectedParagraph = await readParagraphSnapshot(context, paragraphIndex, paragraph);
      }

      return toolSuccess("Inserted text into the active document.", {
        target,
        paragraphIndex: paragraphIndex ?? null,
        textLength: text.length,
        affectedParagraph,
      });
    });
  },

  async word_replace_text_range(args) {
    return runWordTool(async (context) => {
      const replacementText = typeof args.text === "string" ? args.text : "";
      const resolved = await resolveParagraphTextRange(context, args, {
        requireNonEmpty: true,
        allowExpectedText: true,
      });
      const range = await locateParagraphRange(
        context,
        resolved.paragraph,
        resolved.paragraphText,
        resolved.startOffset,
        resolved.endOffset,
      );
      range.insertText(replacementText, "Replace");
      await context.sync();
      const paragraph = await readParagraphSnapshot(
        context,
        resolved.paragraphIndex,
        resolved.paragraph,
      );
      const replacedRange = {
        paragraphIndex: resolved.paragraphIndex,
        startOffset: resolved.startOffset,
        endOffset: resolved.startOffset + replacementText.length,
        previousText: resolved.rangeText,
        text: replacementText,
      };
      return toolSuccess(
        `Replaced ${resolved.rangeText.length} character${resolved.rangeText.length === 1 ? "" : "s"} in paragraph ${resolved.paragraphIndex}.`,
        {
          paragraph,
          replacedRange,
        },
      );
    });
  },

  async word_replace_selection(args) {
    return runWordTool(async (context) => {
      const text = typeof args.text === "string" ? args.text : "";
      const selection = context.document.getSelection();
      selection.load("text");
      const selectionParagraphs = selection.paragraphs;
      selectionParagraphs.load("items");
      await context.sync();
      const selectedText = getSelectionText(selection);
      if (!selectedText) {
        return toolError("There is no active text selection to replace.");
      }
      selection.insertText(text, "Replace");
      await context.sync();
      let affectedParagraph: Record<string, unknown> | null = null;
      let affectedParagraphs: Record<string, unknown>[] = [];
      let verificationError: string | null = null;
      try {
        for (const [index, paragraph] of selectionParagraphs.items.entries()) {
          paragraph.load("text,style");
          (paragraph as any).__bridgeSelectionParagraphIndex = index;
        }
        if (selectionParagraphs.items.length > 0) {
          await context.sync();
          affectedParagraphs = selectionParagraphs.items.map((paragraph: any, index: number) => ({
            selectionParagraphIndex:
              typeof paragraph.__bridgeSelectionParagraphIndex === "number"
                ? paragraph.__bridgeSelectionParagraphIndex
                : index,
            text: paragraph.text ?? "",
            style: readParagraphStyle(paragraph),
            summary: summarizeParagraph(paragraph.text ?? ""),
          }));
          affectedParagraph = affectedParagraphs[0] ?? null;
        }
      } catch (error) {
        affectedParagraphs = [];
        affectedParagraph = null;
        verificationError =
          error instanceof Error && error.message.trim()
            ? error.message
            : "The selection was replaced, but the affected paragraphs could not be reread afterward.";
      }
      return toolSuccess(
        verificationError
          ? "Replaced the active selection, but post-write verification could not confirm the updated paragraphs."
          : "Replaced the active selection.",
        {
        previousText: selectedText,
        newText: text,
        previousLength: selectedText.length,
        textLength: text.length,
        affectedParagraph,
        affectedParagraphs,
        affectedParagraphCount: affectedParagraphs.length,
        verificationState:
          verificationError == null
            ? affectedParagraphs.length > 0
              ? "verified"
              : "best-effort"
            : "verification_failed",
        verificationError,
        verificationScope:
          affectedParagraphs.length > 1
            ? "selection-paragraphs"
            : affectedParagraphs.length === 1
              ? "single-paragraph"
              : "best-effort",
        },
      );
    });
  },

  async word_search_text(args) {
    return runWordTool(async (context) => {
      const query = getSafeString(args.query);
      if (!query) {
        return toolError("query is required.");
      }
      const maxMatches = normalizeMaxMatches(args.maxMatches);
      const paragraphs = await getParagraphs(context);
      const { matches, totalMatches } = findSearchCandidates(
        paragraphs,
        query,
        resolveSearchOptions(args),
        maxMatches,
      );
      return toolSuccess(
        `Found ${totalMatches} match${totalMatches === 1 ? "" : "es"} for "${query}".`,
        {
          query,
          totalMatches,
          truncated: totalMatches > maxMatches,
          matches,
        },
      );
    });
  },

  async word_search_and_replace(args) {
    return runWordTool(async (context) => {
      const query = getSafeString(args.query);
      if (!query) {
        return toolError("query is required.");
      }
      const replacement = typeof args.replacement === "string" ? args.replacement : "";
      const maxMatches = normalizeMaxMatches(args.maxMatches);
      const paragraphs = await getParagraphs(context);
      const searchOptions = resolveSearchOptions(args);
      const { matches, totalMatches } = findSearchCandidates(
        paragraphs,
        query,
        searchOptions,
        maxMatches,
      );
      const targetMatchIndexes = Array.isArray(args.targetMatchIndexes)
        ? (args.targetMatchIndexes as unknown[])
            .filter((value): value is number => typeof value === "number" && Number.isInteger(value))
        : [];
      const replaceable =
        targetMatchIndexes.length > 0
          ? matches.filter((match) => targetMatchIndexes.includes(match.matchIndex))
          : matches;
      const uniqueRequestedMatchIndexes = [...new Set(targetMatchIndexes)].sort((a, b) => a - b);
      if (totalMatches > 1 && targetMatchIndexes.length === 0) {
        return toolError(
          "The query matched more than one location. Call word_search_text first and pass targetMatchIndexes, or narrow the query.",
          {
            query,
            totalMatches,
            matches,
          },
        );
      }
      if (uniqueRequestedMatchIndexes.length > 0) {
        const availableIndexes = new Set(matches.map((match) => match.matchIndex));
        const missingMatchIndexes = uniqueRequestedMatchIndexes.filter(
          (index) => !availableIndexes.has(index),
        );
        if (missingMatchIndexes.length > 0) {
          return toolError(
            "One or more requested targetMatchIndexes were not available in the current candidate set. Re-run word_search_text with a higher maxMatches or choose from the returned matches.",
            {
              query,
              totalMatches,
              truncated: totalMatches > maxMatches,
              requestedMatchIndexes: uniqueRequestedMatchIndexes,
              missingMatchIndexes,
              matches,
            },
          );
        }
      }
      if (replaceable.length === 0) {
        return toolError(
          "No replaceable matches were found for the requested targetMatchIndexes.",
          {
            query,
            totalMatches,
            truncated: totalMatches > maxMatches,
            requestedMatchIndexes: uniqueRequestedMatchIndexes,
            matches,
          },
        );
      }
      const affectedParagraphIndexes = new Set<number>();
      const replaceableInApplyOrder = [...replaceable].sort((left, right) => {
        if (left.paragraphIndex !== right.paragraphIndex) {
          return left.paragraphIndex - right.paragraphIndex;
        }
        if (left.startOffset !== right.startOffset) {
          return right.startOffset - left.startOffset;
        }
        return right.endOffset - left.endOffset;
      });
      for (const match of replaceableInApplyOrder) {
        const { paragraph } = await getParagraphByIndex(context, match.paragraphIndex);
        const range = await locateParagraphRange(
          context,
          paragraph,
          paragraph.text ?? "",
          match.startOffset,
          match.endOffset,
        );
        range.insertText(replacement, "Replace");
        affectedParagraphIndexes.add(match.paragraphIndex);
      }
      await context.sync();
      const affectedParagraphs = await Promise.all(
        [...affectedParagraphIndexes].map((paragraphIndex) =>
          readParagraphSnapshot(context, paragraphIndex),
        ),
      );
      return toolSuccess(
        `Replaced ${replaceable.length} match${replaceable.length === 1 ? "" : "es"} for "${query}".`,
        {
          query,
          replacement,
          replacedMatches: replaceable.map((match) => ({
            ...match,
            previousText: match.text,
            newText: replacement,
          })),
          replacedCount: replaceable.length,
          totalMatches,
          truncated: totalMatches > maxMatches,
          affectedParagraphs,
        },
      );
    });
  },

  async word_apply_style(args) {
    return runWordTool(async (context) => {
      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      const style = getSafeString(args.style);
      if (paragraphIndex == null || !style) {
        return toolError("paragraphIndex and style are required.");
      }
      try {
        const { paragraph } = await getParagraphByIndex(context, paragraphIndex);
        paragraph.style = style;
        await context.sync();
        return toolSuccess(`Applied style ${style} to paragraph ${paragraphIndex}.`, {
          paragraph: await readParagraphSnapshot(context, paragraphIndex, paragraph),
          style,
        });
      } catch (error) {
        return toolError(
          `Style "${style}" could not be applied. Make sure the style exists in this document and retry.`,
          {
            cause: summarizeWordFailure(error, "Word rejected the requested style."),
            style,
            paragraphIndex,
          },
        );
      }
    });
  },

  async word_set_paragraph_format(args) {
    return runWordTool(async (context) => {
      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      if (paragraphIndex == null) {
        return toolError("paragraphIndex is required.");
      }
      const { paragraph } = await getParagraphByIndex(context, paragraphIndex);
      const alignment = getSafeString(args.alignment);
      if (alignment) paragraph.alignment = alignment;
      const spaceBefore = getSafeNumber(args.spaceBefore);
      if (spaceBefore != null) paragraph.spaceBefore = spaceBefore;
      const spaceAfter = getSafeNumber(args.spaceAfter);
      if (spaceAfter != null) paragraph.spaceAfter = spaceAfter;
      const lineSpacing = getSafeNumber(args.lineSpacing);
      if (lineSpacing != null) paragraph.lineSpacing = lineSpacing;
      await context.sync();
      return toolSuccess(`Updated paragraph formatting for paragraph ${paragraphIndex}.`, {
        paragraph: await readParagraphSnapshot(context, paragraphIndex, paragraph),
        alignment: alignment ?? null,
        spaceBefore: spaceBefore ?? null,
        spaceAfter: spaceAfter ?? null,
        lineSpacing: lineSpacing ?? null,
      });
    });
  },

  async word_format_text_range(args) {
    return runWordTool(async (context) => {
      const target = getSafeString(args.target) ?? "selection";
      let range: any;
      let paragraphInfo: Awaited<ReturnType<typeof readParagraphSnapshot>> | null = null;
      if (target === "selection") {
        range = context.document.getSelection();
        range.load("text");
        await context.sync();
        if (!getSelectionText(range)) {
          return toolError("There is no active text selection to format.");
        }
      } else {
        const resolved = await resolveParagraphTextRange(context, args, {
          requireNonEmpty: true,
        });
        range = await locateParagraphRange(
          context,
          resolved.paragraph,
          resolved.paragraphText,
          resolved.startOffset,
          resolved.endOffset,
        );
        paragraphInfo = await readParagraphSnapshot(
          context,
          resolved.paragraphIndex,
          resolved.paragraph,
        );
      }

      if (typeof args.bold === "boolean") {
        range.font.bold = args.bold;
      }
      if (typeof args.italic === "boolean") {
        range.font.italic = args.italic;
      }
      if (typeof args.underline === "boolean") {
        range.font.underline = args.underline ? "Single" : "None";
      }
      if (args.highlightColor === null || typeof args.highlightColor === "string") {
        range.font.highlightColor = args.highlightColor;
      }
      if (typeof args.fontColor === "string") {
        range.font.color = args.fontColor;
      }
      await context.sync();

      if (target === "selection") {
        range.load("text");
        await context.sync();
      }

      return toolSuccess("Applied inline formatting to the requested text range.", {
        target,
        formattedText: range.text ?? null,
        paragraph: paragraphInfo,
        applied: {
          bold: typeof args.bold === "boolean" ? args.bold : undefined,
          italic: typeof args.italic === "boolean" ? args.italic : undefined,
          underline: typeof args.underline === "boolean" ? args.underline : undefined,
          highlightColor:
            args.highlightColor === null || typeof args.highlightColor === "string"
              ? args.highlightColor
              : undefined,
          fontColor: typeof args.fontColor === "string" ? args.fontColor : undefined,
        },
      });
    });
  },

  async word_insert_table(args) {
    return runWordTool(async (context) => {
      const rows = Math.trunc(getSafeNumber(args.rows) ?? 0);
      const columns = Math.trunc(getSafeNumber(args.columns) ?? 0);
      if (rows < 1 || columns < 1) {
        return toolError("rows and columns must both be at least 1.");
      }
      const values = Array.isArray(args.values) ? (args.values as string[][]) : undefined;
      const location = getSafeString(args.location) ?? "end";
      if (location === "end") {
        try {
          context.document.body.insertTable(rows, columns, "End", values);
          await context.sync();
          return toolSuccess("Inserted a table at the end of the document.", {
            rows,
            columns,
            location,
          });
        } catch (error) {
          return toolError(
            `Word rejected a table with ${rows} row(s) and ${columns} column(s). Reduce the requested dimensions and retry.`,
            {
              cause: summarizeWordFailure(error, "Word rejected the requested table dimensions."),
              rows,
              columns,
              location,
            },
          );
        }
      }
      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      if (paragraphIndex == null) {
        return toolError("paragraphIndex is required when location is before or after.");
      }
      try {
        const { paragraph } = await getParagraphByIndex(context, paragraphIndex);
        paragraph.insertTable(
          rows,
          columns,
          location === "before" ? "Before" : "After",
          values,
        );
        await context.sync();
        return toolSuccess(`Inserted a table ${location} paragraph ${paragraphIndex}.`, {
          rows,
          columns,
          location,
          paragraphIndex,
        });
      } catch (error) {
        return toolError(
          `Word rejected a table with ${rows} row(s) and ${columns} column(s). Reduce the requested dimensions and retry.`,
          {
            cause: summarizeWordFailure(error, "Word rejected the requested table dimensions."),
            rows,
            columns,
            location,
            paragraphIndex,
          },
        );
      }
    });
  },

  async word_update_table_cell(args) {
    return runWordTool(async (context) => {
      const tableIndex = Math.trunc(getSafeNumber(args.tableIndex) ?? -1);
      const rowIndex = Math.trunc(getSafeNumber(args.rowIndex) ?? -1);
      const columnIndex = Math.trunc(getSafeNumber(args.columnIndex) ?? -1);
      if (tableIndex < 0 || rowIndex < 0 || columnIndex < 0) {
        return toolError("tableIndex, rowIndex, and columnIndex are required.");
      }
      const text = typeof args.text === "string" ? args.text : "";
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      const table = tables.items[tableIndex];
      if (!table) {
        return toolError(`Table ${tableIndex} does not exist.`);
      }
      const dimensions = await resolveTableDimensions(context, table);
      if (rowIndex >= dimensions.rowCount || columnIndex >= dimensions.columnCount) {
        return toolError(
          `Cell (${rowIndex}, ${columnIndex}) is outside table ${tableIndex} with ${dimensions.rowCount} row(s) and ${dimensions.columnCount} column(s).`,
        );
      }
      const cell = table.getCell(rowIndex, columnIndex);
      cell.body.clear();
      cell.body.insertText(text, "Start");
      await context.sync();
      return toolSuccess(`Updated table ${tableIndex}, row ${rowIndex}, column ${columnIndex}.`, {
        tableIndex,
        rowIndex,
        columnIndex,
        textLength: text.length,
      });
    });
  },

  async word_insert_comment(args) {
    return runWordTool(async (context) => {
      const text = getSafeString(args.text);
      if (!text) {
        return toolError("text is required.");
      }
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      const selectedText = getSelectionText(selection);
      let range = selection;
      const paragraphIndex = getSafeNumber(args.paragraphIndex);
      if (!selectedText) {
        if (paragraphIndex == null) {
          return toolError("There is no active selection. Pass paragraphIndex to target a paragraph.");
        }
        const resolved = await getParagraphByIndex(context, paragraphIndex);
        range = resolved.paragraph.getRange();
      }
      const comment = range.insertComment(text);
      comment.load("id,content,resolved,authorName,creationDate");
      await context.sync();
      const comments = await resolveComments(context);
      return toolSuccess("Inserted a native Word comment.", {
        commentId: String(comment.id),
        content: comment.content ?? text,
        resolved: Boolean(comment.resolved),
        authorName: comment.authorName ?? null,
        createdDate: comment.creationDate ?? null,
        totalComments: comments.length,
        targetParagraphIndex: paragraphIndex ?? null,
      });
    });
  },

  async word_reply_to_comment(args) {
    return runWordTool(async (context) => {
      const commentId = getSafeString(args.commentId);
      const text = getSafeString(args.text);
      if (!commentId || !text) {
        return toolError("commentId and text are required.");
      }
      const comment = await findCommentById(context, commentId);
      comment.reply(text);
      await context.sync();
      const refreshed = await findCommentById(context, commentId);
      return toolSuccess(`Replied to comment ${commentId}.`, {
        commentId,
        text,
        replyCount: refreshed.replies.items.length,
      });
    });
  },

  async word_resolve_comment(args) {
    return runWordTool(async (context) => {
      const commentId = getSafeString(args.commentId);
      if (!commentId) {
        return toolError("commentId is required.");
      }
      const comment = await findCommentById(context, commentId);
      comment.resolved = true;
      await context.sync();
      const refreshed = await findCommentById(context, commentId);
      return toolSuccess(`Resolved comment ${commentId}.`, {
        commentId,
        resolved: Boolean(refreshed.resolved),
      });
    });
  },

  async word_delete_comment(args) {
    return runWordTool(async (context) => {
      const commentId = getSafeString(args.commentId);
      if (!commentId) {
        return toolError("commentId is required.");
      }
      const comment = await findCommentById(context, commentId);
      comment.delete();
      await context.sync();
      const comments = await resolveComments(context);
      return toolSuccess(`Deleted comment ${commentId}.`, {
        commentId,
        deleted: true,
        totalComments: comments.length,
      });
    });
  },

  async word_reopen_comment(args) {
    return runWordTool(async (context) => {
      const commentId = getSafeString(args.commentId);
      if (!commentId) {
        return toolError("commentId is required.");
      }
      const comment = await findCommentById(context, commentId);
      comment.resolved = false;
      await context.sync();
      const refreshed = await findCommentById(context, commentId);
      return toolSuccess(`Reopened comment ${commentId}.`, {
        commentId,
        resolved: Boolean(refreshed.resolved),
      });
    });
  },

  async word_enable_track_changes() {
    return runWordTool(async (context) => {
      context.document.changeTrackingMode = "TrackAll";
      await context.sync();
      return toolSuccess("Enabled track changes.", { trackingMode: "TrackAll" });
    });
  },

  async word_disable_track_changes() {
    return runWordTool(async (context) => {
      context.document.changeTrackingMode = "Off";
      await context.sync();
      return toolSuccess("Disabled track changes.", { trackingMode: "Off" });
    });
  },

  async word_accept_revision(args) {
    return runWordTool(async (context) => {
      const { revision, reference, matchedBy } = await findTrackedChangeByReference(context, args);
      const revisionId = reference.revisionId;
      revision.accept();
      await context.sync();
      const remainingChanges = context.document.body.getTrackedChanges();
      remainingChanges.load("items");
      await context.sync();
      return toolSuccess(`Accepted revision ${revisionId}.`, {
        revisionId,
        revisionToken: reference.revisionToken,
        matchedBy,
        affectedRangeText: reference.scopeText,
        affectedParagraph:
          reference.paragraphIndex != null
            ? await tryReadParagraphSnapshot(context, reference.paragraphIndex)
            : null,
        remainingRevisionCount: remainingChanges.items.length,
        note:
          matchedBy === "revisionId"
            ? "Used legacy revisionId matching. Prefer revisionToken to avoid drift."
            : reference.paragraphIndex != null
              ? undefined
              : "Paragraph localization after the mutation is best-effort and may be unavailable for shifted or multi-paragraph revisions.",
      });
    });
  },

  async word_reject_revision(args) {
    return runWordTool(async (context) => {
      const { revision, reference, matchedBy } = await findTrackedChangeByReference(context, args);
      const revisionId = reference.revisionId;
      revision.reject();
      await context.sync();
      const remainingChanges = context.document.body.getTrackedChanges();
      remainingChanges.load("items");
      await context.sync();
      return toolSuccess(`Rejected revision ${revisionId}.`, {
        revisionId,
        revisionToken: reference.revisionToken,
        matchedBy,
        affectedRangeText: reference.scopeText,
        affectedParagraph:
          reference.paragraphIndex != null
            ? await tryReadParagraphSnapshot(context, reference.paragraphIndex)
            : null,
        remainingRevisionCount: remainingChanges.items.length,
        note:
          matchedBy === "revisionId"
            ? "Used legacy revisionId matching. Prefer revisionToken to avoid drift."
            : reference.paragraphIndex != null
              ? undefined
              : "Paragraph localization after the mutation is best-effort and may be unavailable for shifted or multi-paragraph revisions.",
      });
    });
  },

  async word_accept_all_revisions() {
    return runWordTool(async (context) => {
      const changes = await resolveTrackedChanges(context);
      const revisionCount = changes.length;
      context.document.body.getTrackedChanges().acceptAll();
      await context.sync();
      return toolSuccess(
        revisionCount > 0
          ? `Accepted ${revisionCount} tracked revision${revisionCount === 1 ? "" : "s"}.`
          : "There were no tracked revisions to accept.",
        {
          revisionCount,
          changed: revisionCount > 0,
        },
      );
    });
  },

  async word_reject_all_revisions() {
    return runWordTool(async (context) => {
      const changes = await resolveTrackedChanges(context);
      const revisionCount = changes.length;
      context.document.body.getTrackedChanges().rejectAll();
      await context.sync();
      return toolSuccess(
        revisionCount > 0
          ? `Rejected ${revisionCount} tracked revision${revisionCount === 1 ? "" : "s"}.`
          : "There were no tracked revisions to reject.",
        {
          revisionCount,
          changed: revisionCount > 0,
        },
      );
    });
  },
};

export function createWordBridgeTools(): ExecutableWordTool[] {
  return WORD_TOOL_CONTRACTS.map((contract: WordToolContract) => {
    const execute = WORD_TOOL_EXECUTORS[contract.name];
    if (!execute) {
      throw new Error(`Missing Word tool executor for ${contract.name}`);
    }
    return {
      name: contract.name,
      label: contract.title,
      description: contract.description,
      parameters: contract.parameters,
      requiredCapability: contract.requiredCapability,
      execute: async (_toolCallId, args) => {
        try {
          const rawArgs = normalizeWordToolArgs(contract.name, asObject(args));
          if (contract.name === "word_format_text_range") {
            const unsupportedKeys = unsupportedFormatTextRangeKeys(rawArgs);
            if (unsupportedKeys.length > 0) {
              return toolError(
                `Unsupported inline formatting fields: ${unsupportedKeys.join(", ")}. Supported fields: bold, italic, underline, highlightColor, fontColor.`,
              );
            }
          }
          const parsedArgs = contract.inputSchema.safeParse(rawArgs);
          if (!parsedArgs.success) {
            return toolError(parsedArgs.error.message);
          }
          return await execute(parsedArgs.data);
        } catch (error) {
          return toolError(
            error instanceof Error ? error.message : "Word tool execution failed.",
          );
        }
      },
    };
  });
}

export function getWordBridgeToolNames() {
  return WORD_TOOL_CONTRACTS.map((tool) => tool.name);
}
