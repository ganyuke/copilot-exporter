import { fetchCopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";

export async function exportAllDirect(conversationIds: string[], callback: (progress: number) => void) {
    console.log(`${APP_TAG} Getting MSAL ids...`);
    const msalIds = getMsalIds();
    console.log(`${APP_TAG} Getting access token...`);
    const accessToken = await getAccessToken(msalIds);
    const delay = 1500;
    for (let i = 0; i < conversationIds.length; i++) {
        const conversationId = conversationIds[i];
        const blob = await fetchCopilotConversation(accessToken, msalIds.localAccountId, msalIds.tenantId, conversationId);
        console.log(`${APP_TAG} Completed download for conversation ${conversationId}`)
        callback(i);
        downloadBlobAsFile(blob, `m365-copilot-${conversationId}.json`)
        await new Promise((resolve) => setTimeout(resolve, delay));
    }
}