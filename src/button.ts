import { APP_TAG } from "./main";

async function hookIntoSidebar(callback: Function) {
    const sidebarRoot = await waitForElement('[data-testid="appbar-v2"]'); // this should always be here

    if (!sidebarRoot) {
        const log = `${APP_TAG} Could not hook into sidebar root!`;
        console.error(log);
        throw new Error(log);
    }

    const observer = new MutationObserver(() => {
        const allConversationsBtn = document.getElementById("all-history") as HTMLButtonElement;

        if (allConversationsBtn && allConversationsBtn.dataset.hooked !== "1") {
            allConversationsBtn.dataset.hooked = "1"
            createExportButton(allConversationsBtn, callback);
        }
    });

    observer.observe(sidebarRoot, {
        childList: true,
        subtree: true,
    });

}

function waitForElement(selector: string, timeout = 10000): Promise<HTMLElement> {
    return new Promise((resolve, reject) => {
        const found = document.querySelector<HTMLElement>(selector);
        if (found) return resolve(found);

        const observer = new MutationObserver(() => {
            const el = document.querySelector<HTMLElement>(selector);
            if (el) {
                observer.disconnect();
                resolve(el);
            }
        });

        observer.observe(document.body, { childList: true, subtree: true });

        setTimeout(() => {
            observer.disconnect();
            reject(new Error(`Element not found for selector ${selector}`));
        }, timeout);
    });
}

function createExportButton(baseBtn: HTMLButtonElement, callback: Function) {
    // Steal the Fluent UI "All conversations" button
    const exportBtn = baseBtn.cloneNode(true) as HTMLButtonElement;
    exportBtn.id = 'export-conversations';
    exportBtn.setAttribute('aria-label', 'Export conversations');
    exportBtn.value = 'export-conversations';

    const span = exportBtn.querySelector('span');
    if (span) {
        span.textContent = 'Export conversations';
        span.setAttribute('aria-label', 'Export conversations');
    }

    exportBtn.addEventListener('click', () => {
        callback();
    });

    // I sure hope this still has a parent!
    baseBtn.parentElement?.append(exportBtn);

}
/*
async function sidebarOpened(callback: Function) {
    // the "Chat" section isn't loaded unless the sidebar is expanded
    const chatBtn = document.getElementById("chat");

    // do nothing until the sidebar is opened
    if (chatBtn) {
        return;
    }

    // this is the "All conversations" button
    const baseSelector = 'button#all-history';

    // this is the expand fold labeled "Conversations"
    // under the M365 sidebar "Chat" heading. It needs
    // to be expanded for "All conversations" to appear.
    const conversationBtn = await waitForElement('button[aria-label="Conversations"]');

    if (conversationBtn && !conversationBtn.dataset.hooked) {
        conversationBtn.dataset.hooked = "true";

        // we'll listen in for when "Conversations"
        // is opened so we can inject the "Export" button.
        conversationBtn.addEventListener("click", () => {
            const allConversationsBtn = document.getElementById("all-history") as HTMLButtonElement;
            if (!allConversationsBtn) {
                throw new Error(`${APP_TAG} Failed to find "All conversations" button.`);
            }
            createExportButton(allConversationsBtn, callback);
        });
    }

    if (conversationBtn) {
        conversationBtn.addEventListener()
    } else {
        // If this doesn't exist... um. We'll just wait
        // until the "All conversations" button exists
        // (if ever...)
        try {
            const allConversationsBtn = await waitForElement(baseSelector) as HTMLButtonElement;
            createExportButton(allConversationsBtn, callback);
        } catch (err) {
            console.error(`${APP_TAG} Failed to insert export button with error: ${err}`);
        }
    }
}

export async function injectExportButton(callback: Function) {
    const sidebarBtn = await waitForElement('button[aria-label="Expand side navigation panel"]') as HTMLButtonElement;
    sidebarBtn.addEventListener('click', () => sidebarOpened(callback));

    // in case it's already open
    sidebarOpened(callback);
}*/

export function injectExportButton(callback: Function) {
    hookIntoSidebar(callback);
}