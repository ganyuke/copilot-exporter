import { fetchCopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";

export async function exportAllDirect(conversationIds: string[], callback: (progress: number) => void) {
    console.log(`${APP_TAG} Getting MSAL ids...`);
    const msalIds = getMsalIds();
    console.log(`${APP_TAG} Getting access token...`);
    const accessToken =  await getAccessToken(msalIds);
    const delay = 1500;
    let progress = 0;
    for (const conversationId of conversationIds) {
        const blob = await fetchCopilotConversation(accessToken, msalIds.localAccountId, msalIds.tenantId, conversationId);
        downloadBlobAsFile(blob, `m365-copilot-${conversationId}.json`)
        await new Promise((resolve) => setTimeout(resolve, delay));
        progress++;
        callback(progress);
    }
}