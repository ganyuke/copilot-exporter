import { deleteCopilotConversation, fetchCopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";

const FETCH_DELAY = 1500;

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

export async function exportBulkDirect(conversationIds: string[], callback: (progress: number) => void) {
    const { token, localAccountId, tenantId } = await getTokenAndIds();
    
    for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        const blob = await fetchCopilotConversation(token, localAccountId, tenantId, conversationId);
        console.log(`${APP_TAG} Completed download for conversation ${conversationId}`)
        callback(i);
        downloadBlobAsFile(blob, `m365-copilot-${conversationId}.json`)
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