import { createZip, type ZipFile } from "@litejs/zip";
import type { CopilotConversation } from "./api";
import { deleteCopilotConversation, fetchCopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { mapToConversationJson } from "./converters/chatgpt";
import { mapToMarkdown } from "./converters/markdown";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";
import sanitize from "sanitize-filename";

const FETCH_DELAY = 1500;

export type ExportFormat = "json" | "markdown" | "chatgpt";
export type OutputMode = "individual" | "combined" | "zip";

export type ExportCallback = (event:
    | { phase: 'start'; index: number }
    | { phase: 'success'; index: number }
    | { phase: 'error'; index: number; error: string }
) => void;

export type ExportResult = {
    successCount: number;
    totalCount: number;
};

function sanitizeFilename(name: string): string {
    const sanitized = sanitize(name, { replacement: '_' });
    return sanitized;
}

export function exportFilename(conversation: CopilotConversation, format: ExportFormat): string {
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

function zipArchiveFilename(format: ExportFormat): string {
    switch (format) {
        case "markdown":
            return "m365-copilot-export-markdown.zip";
        case "chatgpt":
            return "m365-copilot-export-chatgpt-json.zip";
        default:
            return "m365-copilot-export-json.zip";
    }
}

function combinedArchiveFilename(format: ExportFormat): string {
    switch (format) {
        case "chatgpt":
            return "conversations.json";
        default:
            return "m365-copilot-conversations.json";
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

function conversationToCombinedEntry(conversation: CopilotConversation, format: ExportFormat): unknown {
    switch (format) {
        case "chatgpt":
            return mapToConversationJson(conversation);
        default:
            return conversation;
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

async function fetchConversation(
    token: string,
    localAccountId: string,
    tenantId: string,
    conversationId: string,
): Promise<CopilotConversation> {
    const blob = await fetchCopilotConversation(token, localAccountId, tenantId, conversationId);
    return JSON.parse(await blob.text()) as CopilotConversation;
}

async function exportIndividual(
    conversationIds: string[],
    callback: ExportCallback,
    format: ExportFormat,
    token: string,
    localAccountId: string,
    tenantId: string,
): Promise<ExportResult> {
    let successCount = 0;

    for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        callback({ phase: 'start', index: i });
        try {
            const conversation = await fetchConversation(token, localAccountId, tenantId, conversationId);
            const exportBlob = conversationToBlob(conversation, format);
            downloadBlobAsFile(exportBlob, exportFilename(conversation, format));
            console.log(`${APP_TAG} Completed download for conversation ${conversationId}`);
            callback({ phase: 'success', index: i });
            successCount++;
        } catch (err) {
            callback({
                phase: 'error',
                index: i,
                error: err instanceof Error ? err.message : String(err),
            });
        }
        await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
    }

    return { successCount, totalCount: conversationIds.length };
}

async function exportCombined(
    conversationIds: string[],
    callback: ExportCallback,
    format: ExportFormat,
    token: string,
    localAccountId: string,
    tenantId: string,
): Promise<ExportResult> {
    const combined: unknown[] = [];
    let successCount = 0;

    for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        callback({ phase: 'start', index: i });
        try {
            const conversation = await fetchConversation(token, localAccountId, tenantId, conversationId);
            combined.push(conversationToCombinedEntry(conversation, format));
            console.log(`${APP_TAG} Completed export for conversation ${conversationId}`);
            callback({ phase: 'success', index: i });
            successCount++;
        } catch (err) {
            callback({
                phase: 'error',
                index: i,
                error: err instanceof Error ? err.message : String(err),
            });
        }
        await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
    }

    if (combined.length > 0) {
        const blob = new Blob([JSON.stringify(combined, null, 2)], { type: "application/json" });
        downloadBlobAsFile(blob, combinedArchiveFilename(format));
    }

    return { successCount, totalCount: conversationIds.length };
}

async function exportZip(
    conversationIds: string[],
    callback: ExportCallback,
    format: ExportFormat,
    token: string,
    localAccountId: string,
    tenantId: string,
): Promise<ExportResult> {
    let successCount = 0;
    const files: ZipFile[] = [];

    for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        callback({ phase: 'start', index: i });
        try {
            const conversation = await fetchConversation(token, localAccountId, tenantId, conversationId);
            const blob = conversationToBlob(conversation, format);
            files.push({
                name: exportFilename(conversation, format),
                content: new TextEncoder().encode(await blob.text()),
                time: conversation.createTimeUtc,
            });
            console.log(`${APP_TAG} Completed export for conversation ${conversationId}`);
            callback({ phase: 'success', index: i });
            successCount++;
        } catch (err) {
            callback({
                phase: 'error',
                index: i,
                error: err instanceof Error ? err.message : String(err),
            });
        }
        await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
    }

    if (successCount > 0) {
        const zipUint8Array = await createZip(files);
        const zipBlob = new Blob([zipUint8Array as BlobPart], { type: "application/zip" });
        downloadBlobAsFile(zipBlob, zipArchiveFilename(format));
    }

    return { successCount, totalCount: conversationIds.length };
}

export async function exportBulkDirect(
    conversationIds: string[],
    callback: ExportCallback,
    format: ExportFormat = "json",
    outputMode: OutputMode = "individual",
): Promise<ExportResult> {
    const { token, localAccountId, tenantId } = await getTokenAndIds();

    switch (outputMode) {
        case "combined":
            return exportCombined(conversationIds, callback, format, token, localAccountId, tenantId);
        case "zip":
            return exportZip(conversationIds, callback, format, token, localAccountId, tenantId);
        default:
            return exportIndividual(conversationIds, callback, format, token, localAccountId, tenantId);
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
