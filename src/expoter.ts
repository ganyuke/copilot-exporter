import sanitize from "sanitize-filename";
import type { CopilotConversation } from "./api";
import { deleteCopilotConversation, fetchCopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { mapToConversationJson } from "./converters/chatgpt";
import { mapToMarkdown } from "./converters/markdown";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";

const FETCH_DELAY = 1500;

export type ExportFormat = "json" | "markdown" | "chatgpt";

export type ExportCallback = (event:
    | { phase: 'start'; index: number }
    | { phase: 'success'; index: number }
    | { phase: 'error'; index: number; error: string }
) => void;

function sanitizeFilename(name: string): string {
    const sanitized = sanitize(name, { replacement: '_' });
    return sanitized;
}

function exportFilename(conversation: CopilotConversation, format: ExportFormat): string {
    const base = sanitizeFilename(conversation.chatName) || conversation.conversationId;
    switch (format) {
        case "markdown":
            return `m365-copilot-${base}.md`;
        case "chatgpt":
            return `m365-copilot-as-chatgpt-${base}.json`;
        default:
            return `m365-copilot-${base}.json`;
    }
}

function conversationToBlob(conversation: CopilotConversation, format: ExportFormat): Blob {
    switch (format) {
        case "markdown":
            return new Blob([mapToMarkdown(conversation)], { type: "text/markdown" });
        case "chatgpt":
            return new Blob(
                [JSON.stringify(mapToConversationJson(conversation), null, 2)],
                { type: "application/json" },
            );
        default:
            return new Blob([JSON.stringify(conversation, null, 2)], { type: "application/json" });
    }
}

async function getTokenAndIds() {
    console.log(`${APP_TAG} Getting MSAL ids...`);
    const msalIds = getMsalIds();
    console.log(`${APP_TAG} Getting access token...`);
    const accessToken = await getAccessToken(msalIds);
    return {
        token: accessToken,
        ...msalIds
    };
}

export async function exportBulkDirect(
    conversationIds: string[],
    callback: ExportCallback,
    format: ExportFormat = "json",
): Promise<void> {
    const { token, localAccountId, tenantId } = await getTokenAndIds();

    for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        callback({ phase: 'start', index: i });
        try {
            const blob = await fetchCopilotConversation(token, localAccountId, tenantId, conversationId);
            const conversation = JSON.parse(await blob.text()) as CopilotConversation;
            const exportBlob = conversationToBlob(conversation, format);
            downloadBlobAsFile(exportBlob, exportFilename(conversation, format));
            console.log(`${APP_TAG} Completed download for conversation ${conversationId}`);
            callback({ phase: 'success', index: i });
        } catch (err) {
            callback({
                phase: 'error',
                index: i,
                error: err instanceof Error ? err.message : String(err),
            });
        }
        await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
    }
}

export async function deleteBulk(conversationIds: string[], callback: (progress: number) => void) {
    const { token, localAccountId, tenantId } = await getTokenAndIds();
    // afaik you can only do one-by-one in the M365 dashboard.
    // maybe it's a little suspicious to hit the Substrate API
    // with multiple like this...
    await deleteCopilotConversation(token, localAccountId, tenantId, conversationIds);
    callback(conversationIds.length - 1);
    console.log(`${APP_TAG} Completed deletion for conversations ${conversationIds.join()}`)
    /*for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        await deleteCopilotConversation(token, localAccountId, tenantId, [conversationId]);
        console.log(`${APP_TAG} Completed deletion for conversation ${conversationId}`)
        callback(i);
        await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
    }*/
}
