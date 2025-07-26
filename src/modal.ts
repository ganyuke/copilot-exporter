import { fetchCopilotChats } from "./api";
import { deleteBulk, exportBulkDirect } from "./expoter";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";

type TransportObject = {
    id: string;
    title: string;
}

export function showExportModal() {
    if (document.getElementById('copilotExportOverlay')) return;

    const overlay = document.createElement('div');
    overlay.id = 'copilotExportOverlay';
    overlay.style.cssText = `
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.5);
    display: flex; align-items: center; justify-content: center;
    z-index: 9999;
  `;

    // allow closing the modal by
    // clicking outside the modal
    overlay.addEventListener("click", () => {
        overlay.remove();
    })

    const modal = document.createElement('div');

    // prevent clicking the modal
    // itself from closing it
    modal.addEventListener("click", (e) => {
        e.stopPropagation();
    });

    modal.style.cssText = `
    background: white; padding: 20px; border-radius: 8px;
    min-width: 400px; max-width: 90%;
    box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    font-family: sans-serif;
  `;

    modal.innerHTML = `
    <h2 style="margin-top:0;">Export conversations</h2>
    <p style="margin-bottom: 1em;">Export from API</p>

    <div style="display:flex;column-gap:0.5em;">
      <label style="flex-grow:1;" for="conversation-fetch-list-max">Max conversations to fetch</label>
      <input type="number" id="conversation-fetch-list-max" name="quantity" min="0">
      <button id="conversation-refetch">Refetch</button>
    </div>

    <div style="margin: 1em 0; border: 1px solid #ccc; padding: 0.5em; max-height: 200px; overflow-y: auto;">
      <label><input type="checkbox" id="selectAllCheckbox"> Select All</label>
      <div id="chatList" style="margin-top: 0.5em; color: #666">Loading…</div>
    </div>

    <div style="display: flex; justify-content: space-between; align-items: center;">
      <select>
        <option>JSON</option>
      </select>
      <div>
        <button id="delete-conversations-button">Delete</button>
        <button id="export-conversations-button">Export</button>
      </div>
    </div>
  `;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    function createProgressBar(idsToExport: TransportObject[], initalString: string) {
        if (idsToExport.length < 1) {
            return;
        }
        const existingProgressBarContainer = document.querySelector("#chat-export-progress-bar-container");
        if (existingProgressBarContainer) {
            return;
            // existingProgressBarContainer.remove();
        }

        const progressBarContainer = document.createElement("div") as HTMLDivElement;
        progressBarContainer.id = "chat-export-progress-bar-container"
        progressBarContainer.style = "display: flex;flex-direction: column;margin-top: 0.5em;"
        const progressBar = document.createElement("progress") as HTMLProgressElement;
        const label = document.createElement("label") as HTMLLabelElement;
        label.style = "display:flex;"
        const titleSpan = document.createElement("span") as HTMLSpanElement;
        const progressTextSpan = document.createElement("span") as HTMLSpanElement;
        titleSpan.style = "flex-grow:1;"
        progressBar.id = "chat-export-progress-bar";
        progressBar.max = idsToExport.length;
        progressBar.value = 0;
        label.htmlFor = "chat-export-progress-bar";
        titleSpan.textContent = initalString;
        progressTextSpan.textContent = `0/${idsToExport.length}`
        label.append(titleSpan, progressTextSpan);
        progressBarContainer.append(label, progressBar);
        modal.append(progressBarContainer);

        const progressUpdater = (progress: number) => {
            titleSpan.textContent = idsToExport[progress].title;
            progressTextSpan.textContent = `${progress + 1}/${idsToExport.length}`;
            progressBar.value = progress + 1;

            if (progressBar.value === progressBar.max) {
                setTimeout(() => {
                    progressBarContainer.remove();
                }, 3000);
            };
        }

        return progressUpdater;
    }

    async function fetchChats() {
        const inputNumber = document.getElementById("conversation-fetch-list-max")! as HTMLInputElement;
        const n = inputNumber.valueAsNumber;
        const maxChats = isNaN(n) ? 15 : n; // default: 15 chats

        console.log(`${APP_TAG} Getting MSAL ids...`);
        const msalIds = getMsalIds();
        console.log(`${APP_TAG} Getting access token...`);
        const accessToken = await getAccessToken(msalIds);
        const copilotChatList = await fetchCopilotChats(accessToken, msalIds.localAccountId, msalIds.tenantId, maxChats);

        const chatList = document.getElementById('chatList')!;
        chatList.innerText = "";

        copilotChatList.chats.forEach((data) => {
            const label = document.createElement("label");
            label.style = "column-gap:0.5em;display:flex;"
            const checkbox = document.createElement("input");
            const span = document.createElement("span");
            checkbox.type = "checkbox";
            checkbox.dataset["id"] = data.conversationId;
            checkbox.dataset["title"] = data.chatName;
            span.innerText = data.chatName;
            label.append(checkbox);
            label.append(span);
            chatList.appendChild(label);
        })

        const selectAll = document.getElementById("selectAllCheckbox")! as HTMLInputElement;
        selectAll.addEventListener("change", () => {
            const checkboxes = document.querySelectorAll('#chatList input[type="checkbox"]') as NodeListOf<HTMLInputElement>;
            checkboxes.forEach(cb => {
                cb.checked = selectAll.checked;
            });
        });

    }

    function getSelectedChats(): TransportObject[] {
        const checkboxes = document.querySelectorAll('#chatList input[type="checkbox"]:checked') as NodeListOf<HTMLInputElement>;
        const listToExport: TransportObject[] = [];
        checkboxes.forEach((c) => {
            const uuid = c.dataset["id"]!;
            const title = c.dataset["title"]!;
            listToExport.push({
                id: uuid,
                title: title,
            })
        });
        return listToExport;
    }

    function exportChats() {
        const idsToExport = getSelectedChats();
        const progressUpdater = createProgressBar(idsToExport, "Exporting...");
        if (!progressUpdater) {
            return;
        }
        exportBulkDirect(idsToExport.map((obj) => obj.id), progressUpdater);
    }

    function deleteChats() {
        const idsToDelete = getSelectedChats();
        const progressUpdater = createProgressBar(idsToDelete, "Deleting...");
        if (!progressUpdater) {
            return;
        }
        deleteBulk(idsToDelete.map((obj) => obj.id), progressUpdater);
    }

    // hook up export button
    const exportBtn = document.getElementById("export-conversations-button")! as HTMLButtonElement;
    exportBtn.addEventListener("click", exportChats)

    // hook up delete button
    const deleteBtn = document.getElementById("delete-conversations-button")! as HTMLButtonElement;
    deleteBtn.addEventListener("click", deleteChats)

    // hook up refetch button
    const refetchButton = document.getElementById("conversation-refetch")! as HTMLButtonElement;
    refetchButton.addEventListener("click", fetchChats);

    // it looks nicer if we populate the list on load
    fetchChats();
}