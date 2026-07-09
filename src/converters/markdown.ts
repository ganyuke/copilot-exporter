/**
 * For M365 Copilot conversation to Markdown format conversion.
 */
import type { CopilotConversation, CopilotMessage } from "../api";

const SOURCE_URL = "https://m365.cloud.microsoft/chat/conversation";
const CITATION_MARKER_RE = /【([^】]+)】/g;

type CitationReference = {
    targetLink?: string;
    isCitedInResponse?: boolean;
    displayData?: {
        content?: string;
    };
};

type CitationContent = {
    label?: string;
    Title?: string;
};

type MessageWithCitations = CopilotMessage & {
    adaptiveCards?: {
        body?: {
            type?: string;
            text?: string;
        }[];
    }[];
    references?: Record<string, CitationReference>;
};

function millisToIsoOffset(ms: number): string {
    return new Date(ms).toISOString().replace("Z", "+00:00");
}

function parseMessageDate(iso: string | undefined | null): Date | null {
    if (!iso) return null;
    const ms = Date.parse(iso);
    return Number.isFinite(ms) ? new Date(ms) : null;
}

function formatTimeElement(date: Date): string {
    const datetime = date.toISOString();
    const title = formatTimeTitle(date);
    const label = formatTimeShort(date);
    return `<time datetime="${datetime}" title="${title}">${label}</time>`;
}

function speakerLabel(author: string): string {
    return author === "user" ? "You" : "Copilot";
}

function formatTimeTitle(date: Date): string {
    return new Intl.DateTimeFormat("en-US", {
        dateStyle: "short",
        timeStyle: "medium",
    }).format(date);
}

function formatTimeShort(date: Date): string {
    return new Intl.DateTimeFormat("en-US", {
        hour: "numeric",
        minute: "2-digit",
    }).format(date);
}

function parseCitationContent(content: string | undefined): CitationContent | null {
    if (!content) return null;
    try {
        return JSON.parse(content) as CitationContent;
    } catch {
        return null;
    }
}

function resolveMessageBody(message: MessageWithCitations): string {
    const rawText = message.text ?? "";
    try {
        const cardText = message.adaptiveCards?.[0]?.body?.[0]?.text;
        if (typeof cardText !== "string") return rawText;

        const references = message.references ?? {};
        const used = new Map<string, { n: string; title: string; url: string }>();

        const rewritten = cardText.replace(CITATION_MARKER_RE, (_match, key: string) => {
            const ref = references[key];
            if (!ref?.targetLink) throw new Error(`missing reference for ${key}`);

            const parsed = parseCitationContent(ref.displayData?.content);
            const n = parsed?.label;
            const title = parsed?.Title;
            if (!n || !title) throw new Error(`incomplete citation metadata for ${key}`);

            if (!used.has(key)) {
                used.set(key, { n, title, url: ref.targetLink });
            }
            return `[${n}]`;
        });

        if (used.size === 0) return rewritten;

        const sources = [...used.values()]
            .sort((a, b) => Number(a.n) - Number(b.n))
            .map(({ n, title, url }) => `[${n}] [${title}](${url})`)
            .join("\n");

        return `${rewritten}\n\n**Sources:**\n\n${sources}`;
    } catch {
        return rawText;
    }
}

function formatMessage(message: CopilotMessage): string {
    const date = parseMessageDate(message.createdAt ?? message.timestamp);
    const lines = [`## ${speakerLabel(message.author)}:`];
    if (date) lines.push(formatTimeElement(date));
    lines.push("", resolveMessageBody(message));
    return lines.join("\n");
}

export function mapToMarkdown(source: CopilotConversation): string {
    const frontmatter = [
        "---",
        `title: ${JSON.stringify(source.chatName)}`,
        `createdAt: ${JSON.stringify(millisToIsoOffset(source.createTimeUtc))}`,
        `updatedAt: ${JSON.stringify(millisToIsoOffset(source.updateTimeUtc))}`,
        `source: ${SOURCE_URL}/${source.conversationId}`,
        "---",
    ].join("\n");

    const body = [
        `# ${source.chatName}`,
        ...source.messages.map((message) => formatMessage(message)),
    ].join("\n\n");

    return `${frontmatter}\n\n${body}\n`;
}
