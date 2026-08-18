export const DEFAULT_WORK_IQ_ENDPOINT = "https://workiq.svc.cloud.microsoft/rest/conversations";

export interface WorkIqSettings {
  workIqExecutablePath: string;
}

export interface WorkIqAttribution {
  providerDisplayName?: string;
  seeMoreWebUrl?: string;
}

export interface WorkIqConversationMessage {
  text: string;
  attributions?: WorkIqAttribution[];
}

export interface WorkIqConversation {
  id: string;
  state?: string;
  turnCount?: number;
  messages: WorkIqConversationMessage[];
}

export interface WorkIqCliAnswer {
  conversationId: string;
  response: string;
}

export const DEFAULT_SETTINGS: WorkIqSettings = {
  workIqExecutablePath: ""
};

export function parseWorkIqCliAnswer(value: unknown): WorkIqCliAnswer {
  const result = requireRecord(value);
  const response = result.response;
  const conversationId = result.conversationId;

  if (result.isError === true) {
    throw new Error(typeof response === "string" && response.trim() ? response : "Work IQ returned an error.");
  }

  if (typeof response !== "string" || !response.trim() || typeof conversationId !== "string") {
    throw new Error("Work IQ CLI returned an unexpected response.");
  }

  return {
    conversationId,
    response: response.trim()
  };
}

export function formatWorkIqCliAnswer(prompt: string, answer: WorkIqCliAnswer): string {
  return [`## Work IQ: ${escapeMarkdownText(prompt.trim())}`, "", answer.response].join("\n");
}

export function buildWorkIqChatRequest(
  prompt: string,
  settings: { timeZone: string; webSearchEnabled: boolean }
): object {
  const trimmedPrompt = prompt.trim();

  if (!trimmedPrompt) {
    throw new Error("Enter a Work IQ prompt.");
  }

  return {
    message: {
      text: trimmedPrompt
    },
    locationHint: {
      timeZone: settings.timeZone.trim() || "UTC"
    },
    contextualResources: {
      webContext: {
        isWebEnabled: settings.webSearchEnabled
      }
    }
  };
}

export function getWorkIqConversationId(value: unknown): string {
  const conversation = requireRecord(value);
  throwIfWorkIqError(conversation);

  const id = conversation.id;

  if (typeof id !== "string" || !id.trim()) {
    throw new Error("Work IQ did not return a conversation ID.");
  }

  return id;
}

export function parseWorkIqConversation(value: unknown): WorkIqConversation {
  const conversation = requireRecord(value);
  throwIfWorkIqError(conversation);

  const id = conversation.id;
  const messages = conversation.messages;

  if (typeof id !== "string" || !id.trim() || !Array.isArray(messages)) {
    throw new Error("Work IQ returned an unexpected conversation response.");
  }

  return {
    id,
    state: typeof conversation.state === "string" ? conversation.state : undefined,
    turnCount: typeof conversation.turnCount === "number" ? conversation.turnCount : undefined,
    messages: messages.map((message) => parseConversationMessage(message))
  };
}

export function formatWorkIqAnswer(prompt: string, conversation: WorkIqConversation): string {
  const answer = [...conversation.messages].reverse().find((message) => message.text.trim());

  if (!answer) {
    throw new Error("Work IQ returned a conversation without an answer.");
  }

  const lines = [`## Work IQ: ${escapeMarkdownText(prompt.trim())}`, "", answer.text.trim()];
  const sources = getUniqueSafeAttributions(answer.attributions ?? []);

  if (sources.length > 0) {
    lines.push("", "### Sources", "");
    sources.forEach((source, index) => {
      lines.push(`${index + 1}. [${escapeMarkdownText(source.label)}](${source.url})`);
    });
  }

  return lines.join("\n");
}

function parseConversationMessage(value: unknown): WorkIqConversationMessage {
  if (!isRecord(value) || typeof value.text !== "string") {
    throw new Error("Work IQ returned an unexpected conversation message.");
  }

  const attributions = Array.isArray(value.attributions)
    ? value.attributions.flatMap((attribution) => {
        if (!isRecord(attribution)) {
          return [];
        }

        return [
          {
            providerDisplayName:
              typeof attribution.providerDisplayName === "string" ? attribution.providerDisplayName : undefined,
            seeMoreWebUrl: typeof attribution.seeMoreWebUrl === "string" ? attribution.seeMoreWebUrl : undefined
          }
        ];
      })
    : undefined;

  return {
    text: value.text,
    attributions
  };
}

function getUniqueSafeAttributions(attributions: WorkIqAttribution[]): Array<{ label: string; url: string }> {
  const sources = new Map<string, string>();

  attributions.forEach((attribution) => {
    const url = getSafeUrl(attribution.seeMoreWebUrl);

    if (url && !sources.has(url)) {
      sources.set(url, attribution.providerDisplayName?.trim() || "Work IQ source");
    }
  });

  return Array.from(sources, ([url, label]) => ({ label, url }));
}

function getLocalTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
  } catch {
    return "UTC";
  }
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

function escapeMarkdownText(value: string): string {
  return value.replace(/([\\`*_[\]{}()#+!|>~])/g, "\\$1");
}

function throwIfWorkIqError(value: Record<string, unknown>): void {
  const error = value.error;

  if (!isRecord(error)) {
    return;
  }

  const message = error.message;
  throw new Error(typeof message === "string" && message.trim() ? message : "Work IQ returned an error.");
}

function requireRecord(value: unknown): Record<string, unknown> {
  if (!isRecord(value)) {
    throw new Error("Work IQ returned an unexpected response.");
  }

  return value;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
