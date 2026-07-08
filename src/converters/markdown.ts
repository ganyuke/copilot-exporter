/**
 * For M365 Copilot conversation to Markdown format conversion.
 */
import type { CopilotConversation, CopilotMessage } from "../api";

const SOURCE_URL = "https://m365.cloud.microsoft/chat/conversation";

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

function formatMessage(message: CopilotMessage): string {
    const date = parseMessageDate(message.createdAt ?? message.timestamp);
    const lines = [`## ${speakerLabel(message.author)}:`];
    if (date) lines.push(formatTimeElement(date));
    lines.push("", message.text ?? "");
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
