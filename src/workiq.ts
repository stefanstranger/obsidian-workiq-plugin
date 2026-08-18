export const DEFAULT_GRAPH_SEARCH_ENDPOINT = "https://graph.microsoft.com/v1.0/search/query";

export interface WorkIqSearchSettings {
  accessToken: string;
  graphSearchEndpoint: string;
  entityTypes: string[];
  maxResults: number;
}

export interface WorkIqSearchHit {
  title: string;
  summary?: string;
  url?: string;
  resource?: Record<string, unknown>;
}

export interface WorkIqSearchResponse {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: WorkIqSearchHit[];
    }>;
  }>;
}

export const DEFAULT_SETTINGS: WorkIqSearchSettings = {
  accessToken: "",
  graphSearchEndpoint: DEFAULT_GRAPH_SEARCH_ENDPOINT,
  entityTypes: ["driveItem", "message", "event"],
  maxResults: 10
};

export function parseWorkIqSearchResponse(value: unknown): WorkIqSearchResponse {
  if (!isRecord(value)) {
    throw new Error("Microsoft Graph returned an unexpected response.");
  }

  const graphError = getGraphErrorMessage(value);

  if (graphError) {
    throw new Error(graphError);
  }

  if (value.value !== undefined && !Array.isArray(value.value)) {
    throw new Error("Microsoft Graph returned an unexpected search result shape.");
  }

  return value as WorkIqSearchResponse;
}

export function buildWorkIqSearchRequest(
  query: string,
  settings: Pick<WorkIqSearchSettings, "entityTypes" | "maxResults">
): object {
  const trimmedQuery = query.trim();

  if (!trimmedQuery) {
    throw new Error("Enter a WorkIQ search query.");
  }

  return {
    requests: [
      {
        entityTypes: settings.entityTypes,
        from: 0,
        query: {
          queryString: trimmedQuery
        },
        size: settings.maxResults
      }
    ]
  };
}

export function flattenWorkIqHits(response: WorkIqSearchResponse): WorkIqSearchHit[] {
  return (
    response.value?.flatMap((value) =>
      value.hitsContainers?.flatMap((container) => container.hits ?? []) ?? []
    ) ?? []
  );
}

export function formatWorkIqHits(query: string, hits: WorkIqSearchHit[]): string {
  const lines = [`## WorkIQ search: ${query.trim()}`, ""];

  if (hits.length === 0) {
    lines.push("_No Microsoft 365 results found._");
    return lines.join("\n");
  }

  hits.forEach((hit, index) => {
    const title = escapeMarkdownText(hit.title || "Untitled result");
    const summary = escapeMarkdownText(normalizeWhitespace(hit.summary ?? getStringResourceValue(hit.resource, "summary")));
    const url = getSafeUrl(hit.url ?? getStringResourceValue(hit.resource, "webUrl"));

    lines.push(`${index + 1}. ${url ? `[${title}](${url})` : title}`);

    if (summary) {
      lines.push(`   - ${summary}`);
    }
  });

  return lines.join("\n");
}

function normalizeWhitespace(value?: string): string {
  return value?.replace(/\s+/g, " ").trim() ?? "";
}

function getStringResourceValue(resource: Record<string, unknown> | undefined, key: string): string | undefined {
  const value = resource?.[key];
  return typeof value === "string" ? value : undefined;
}

function escapeMarkdownText(value: string): string {
  return value.replace(/([\\`*_[\]{}()#+\-.!|>])/g, "\\$1");
}

function getSafeUrl(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  try {
    const url = new URL(value);
    return url.protocol === "https:" ? value : undefined;
  } catch {
    return undefined;
  }
}

function getGraphErrorMessage(value: Record<string, unknown>): string | undefined {
  const error = value.error;

  if (!isRecord(error)) {
    return undefined;
  }

  const message = error.message;
  return typeof message === "string" && message.trim() ? message : "Microsoft Graph returned an error.";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
