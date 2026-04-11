const GENERIC_PAGE_TITLES = new Set(["Word MCP Bridge"]);

function decodeDocumentBasename(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

export function resolveWordDocumentTitle(options: {
  documentUrl?: string | null;
  pageTitle?: string | null;
}): string | null {
  const documentUrl = options.documentUrl?.trim();
  if (documentUrl) {
    const basename = documentUrl.split("/").pop()?.trim();
    if (basename) return decodeDocumentBasename(basename);
  }

  const pageTitle = options.pageTitle?.trim();
  if (!pageTitle || GENERIC_PAGE_TITLES.has(pageTitle)) {
    return null;
  }

  return pageTitle;
}
