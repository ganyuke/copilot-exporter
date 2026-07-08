import { CopilotConversationOverview, fetchCopilotChats } from "./api";
import type { CopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { mapToConversationJson } from "./converters/chatgpt";
import { mapToMarkdown } from "./converters/markdown";
import { deleteBulk, exportBulkDirect, ExportCallback, ExportFormat, OutputMode } from "./expoter";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";
import project from '../package.json' with { type: 'json' };

type TransportObject = {
    id: string;
    title: string;
}

type RowStatus = 'exporting' | 'exported' | 'deleting' | 'deleted' | 'error';

const STATUS_COLORS: Record<RowStatus, string> = {
    exporting: '#ca8a04',
    exported: '#16a34a',
    deleting: '#ca8a04',
    deleted: '#6b7280',
    error: '#dc2626',
};

const STATUS_LABELS: Record<RowStatus, string> = {
    exporting: 'exporting…',
    exported: 'exported',
    deleting: 'deleting…',
    deleted: 'deleted',
    error: 'error',
};

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

    overlay.addEventListener("click", () => {
        overlay.remove();
    })

    const modal = document.createElement('div');

    modal.addEventListener("click", (e) => {
        e.stopPropagation();
    });

    modal.style.cssText = `
    background: white; padding: 20px; border-radius: 8px;
    width: 90vw; max-width: 800px;
    box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    font-family: sans-serif;
  `;

    modal.innerHTML = `
    <h2 style="margin:0;">Export conversations</h2>
    <p style="margin: 0.5rem 0;color: darkorchid;"><a style="color: inherit;" href="${project.repository.url}" target="_blank">M365 Copilot Exporter</a> v${project.version} by <a style="color: inherit;" href="${project.author.url}" target="_blank">${project.author.name}</a></p>

    <div id="chatTableContainer" style="margin: 1em 0; border: 1px solid #ccc; padding: 0.5em;">
      <div id="chatTableToolbar" style="margin-bottom: 0.5em;">
        <div style="display: flex; align-items: center; justify-content: space-between;">
          <label style="font-size: 0.875em;"><input type="checkbox" id="selectAllCheckbox"> Select All</label>
          <span id="selectedCount" style="color: #666; font-size: 0.875em;">(0/0)</span>
        </div>
        <div style="display: flex; align-items: center; gap: 0.5em; margin-top: 0.5em; font-size: 0.875em;">
          <label for="conversation-fetch-list-max" style="flex: 1;">Max conversations</label>
          <input type="number" id="conversation-fetch-list-max" name="quantity" min="0" placeholder="15">
          <button id="conversation-refetch">Refetch</button>
        </div>
      </div>
      <div id="chatTableScroll" style="max-height: 50vh; overflow-y: auto; overflow-x: hidden;">
      <table id="chatTable" style="width: 100%; border-collapse: collapse; table-layout: fixed;">
        <colgroup>
          <col style="width: 32px">
          <col style="width: 38%">
          <col style="width: 22%">
          <col style="width: 22%">
          <col style="width: 18%">
        </colgroup>
        <thead style="position: sticky;top: 0;">
          <tr style="background: lavender; font-size: 0.875em;">
            <th></th>
            <th style="text-align: left; padding: 4px 8px;">Name</th>
            <th style="text-align: left; padding: 4px 8px;">Created</th>
            <th style="text-align: left; padding: 4px 8px;">Updated</th>
            <th style="text-align: left; padding: 4px 8px;">Status</th>
          </tr>
        </thead>
        <tbody id="chatTableBody">
          <tr><td colspan="5" style="color: #666; padding: 8px;">Loading…</td></tr>
        </tbody>
      </table>
      </div>
    </div>

    <div style="display: flex; justify-content: space-between; align-items: center; gap: 0.5em;">
      <div style="display: flex; gap: 0.5em; align-items: center;">
        <select id="export-format-select">
          <option value="json">Copilot JSON</option>
          <option value="markdown">Markdown</option>
          <option value="chatgpt">ChatGPT JSON</option>
        </select>
        <select id="export-output-mode-select">
          <option value="individual">Individual files</option>
          <option value="combined">Combined file</option>
          <option value="zip">Individual files (ZIP)</option>
        </select>
      </div>
      <div>
        <button id="delete-conversations-button">Delete</button>
        <button id="export-conversations-button">Export</button>
      </div>
    </div>

    <div style="margin-top: 1em;">
      <input type="file" id="copilot-json-upload" accept=".json,application/json" multiple hidden>
      <select id="convert-format-select" style="margin-right: 0.5em;">
        <option value="chatgpt">ChatGPT JSON</option>
        <option value="markdown">Markdown</option>
      </select>
      <button id="convert-uploaded-button">Convert uploaded JSON</button>
    </div>
  `;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    function formatPrettyDate(ms: number): string {
        return new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(ms));
    }

    function findStatusCell(conversationId: string): HTMLTableCellElement | null {
        const checkbox = document.querySelector(
            `#chatTableBody input[type="checkbox"][data-id="${CSS.escape(conversationId)}"]`
        );
        return checkbox?.closest('tr')?.querySelector('.status-cell') as HTMLTableCellElement | null ?? null;
    }

    function setRowStatus(conversationId: string, status: RowStatus, error?: string): void {
        const cell = findStatusCell(conversationId);
        if (!cell) return;
        cell.textContent = STATUS_LABELS[status];
        cell.style.color = STATUS_COLORS[status];
        if (status === 'error' && error) {
            cell.title = error;
        } else {
            cell.removeAttribute('title');
        }
    }

    function clearRowStatus(conversationIds: string[]): void {
        for (const id of conversationIds) {
            const cell = findStatusCell(id);
            if (!cell) continue;
            cell.textContent = '';
            cell.style.color = '';
            cell.removeAttribute('title');
        }
    }

    function updateSelectedCount(): void {
        const checkboxes = document.querySelectorAll('#chatTableBody input[type="checkbox"]') as NodeListOf<HTMLInputElement>;
        const selected = document.querySelectorAll('#chatTableBody input[type="checkbox"]:checked').length;
        const loaded = checkboxes.length;
        const countEl = document.getElementById('selectedCount')!;
        countEl.textContent = `(${selected}/${loaded})`;
        const selectAll = document.getElementById('selectAllCheckbox')! as HTMLInputElement;
        selectAll.checked = selected > 0 && selected === loaded;
    }

    type TableRowSnapshot = {
        checked: boolean;
        chatName: string;
        createTimeUtc: number;
        updateTimeUtc: number;
        statusText: string;
        statusColor: string;
        statusTitle: string | null;
    };

    function captureTableState(): Map<string, TableRowSnapshot> {
        const state = new Map<string, TableRowSnapshot>();
        const rows = document.querySelectorAll('#chatTableBody tr[data-conversation-id]');

        for (const row of rows) {
            const id = row.getAttribute('data-conversation-id');
            if (!id) continue;

            const checkbox = row.querySelector('input[type="checkbox"]') as HTMLInputElement | null;
            const statusCell = row.querySelector('.status-cell') as HTMLTableCellElement | null;

            state.set(id, {
                checked: checkbox?.checked ?? false,
                chatName: row.getAttribute('data-chat-name') ?? '',
                createTimeUtc: Number(row.getAttribute('data-create-time')),
                updateTimeUtc: Number(row.getAttribute('data-update-time')),
                statusText: statusCell?.textContent ?? '',
                statusColor: statusCell?.style.color ?? '',
                statusTitle: statusCell?.getAttribute('title') ?? null,
            });
        }

        return state;
    }

    function chatDataMatches(snapshot: TableRowSnapshot, data: CopilotConversationOverview): boolean {
        return snapshot.chatName === data.chatName
            && snapshot.createTimeUtc === data.createTimeUtc
            && snapshot.updateTimeUtc === data.updateTimeUtc;
    }

    function renderChatTable(chats: CopilotConversationOverview[], previousState?: Map<string, TableRowSnapshot>): void {
        const tbody = document.getElementById('chatTableBody')!;
        tbody.innerHTML = '';

        const sorted = [...chats].sort((a, b) => b.updateTimeUtc - a.updateTimeUtc);

        if (sorted.length === 0) {
            const row = document.createElement('tr');
            const cell = document.createElement('td');
            cell.colSpan = 5;
            cell.textContent = 'No conversations found.';
            cell.style.padding = '8px';
            row.appendChild(cell);
            tbody.appendChild(row);
        } else {
            for (const data of sorted) {
                const row = document.createElement('tr');
                row.setAttribute('data-conversation-id', data.conversationId);
                row.setAttribute('data-chat-name', data.chatName);
                row.setAttribute('data-create-time', String(data.createTimeUtc));
                row.setAttribute('data-update-time', String(data.updateTimeUtc));
                row.style.borderBottom = '1px solid #e5e7eb';

                const checkboxTd = document.createElement('td');
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.dataset.id = data.conversationId;
                checkbox.dataset.title = data.chatName;

                const previous = previousState?.get(data.conversationId);
                if (previous) {
                    checkbox.checked = previous.checked;
                }

                checkboxTd.appendChild(checkbox);

                const nameTd = document.createElement('td');
                nameTd.className = 'name-cell';
                nameTd.style.cssText = 'overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding: 4px 8px;';
                nameTd.title = data.chatName;
                nameTd.textContent = data.chatName;

                const createdTd = document.createElement('td');
                createdTd.style.cssText = 'font-size: 0.875em; padding: 4px 8px;';
                createdTd.textContent = formatPrettyDate(data.createTimeUtc);

                const updatedTd = document.createElement('td');
                updatedTd.style.cssText = 'font-size: 0.875em; padding: 4px 8px;';
                updatedTd.textContent = formatPrettyDate(data.updateTimeUtc);

                const statusTd = document.createElement('td');
                statusTd.className = 'status-cell';
                statusTd.style.padding = '4px 8px';

                if (previous && chatDataMatches(previous, data)) {
                    statusTd.textContent = previous.statusText;
                    statusTd.style.color = previous.statusColor;
                    if (previous.statusTitle) {
                        statusTd.title = previous.statusTitle;
                    }
                }

                row.append(checkboxTd, nameTd, createdTd, updatedTd, statusTd);
                tbody.appendChild(row);
            }
        }

        updateSelectedCount();
    }

    function removeProgressBar(): void {
        document.querySelector("#chat-export-progress-bar-container")?.remove();
    }

    function createProgressBar(items: TransportObject[], initialString: string) {
        if (items.length < 1) {
            return;
        }
        removeProgressBar();

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
        progressBar.max = items.length;
        progressBar.value = 0;
        label.htmlFor = "chat-export-progress-bar";
        titleSpan.textContent = initialString;
        progressTextSpan.textContent = `0/${items.length}`
        label.append(titleSpan, progressTextSpan);
        progressBarContainer.append(label, progressBar);
        modal.append(progressBarContainer);

        const progressUpdater = (progress: number) => {
            titleSpan.textContent = items[progress].title;
            progressTextSpan.textContent = `${progress + 1}/${items.length}`;
            progressBar.value = progress + 1;

            if (progressBar.value === progressBar.max) {
                setTimeout(() => {
                    progressBarContainer.remove();
                }, 3000);
            };
        }

        return progressUpdater;
    }

    function createExportProgressHandler(items: TransportObject[], initialString: string): ExportCallback | undefined {
        if (items.length < 1) {
            return;
        }
        removeProgressBar();

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
        progressBar.max = items.length;
        progressBar.value = 0;
        label.htmlFor = "chat-export-progress-bar";
        titleSpan.textContent = initialString;
        progressTextSpan.textContent = `0/${items.length}`
        label.append(titleSpan, progressTextSpan);
        progressBarContainer.append(label, progressBar);
        modal.append(progressBarContainer);

        let completed = 0;

        const handler: ExportCallback = (event) => {
            const item = items[event.index];

            if (event.phase === 'start') {
                titleSpan.textContent = item.title;
                setRowStatus(item.id, 'exporting');
            } else if (event.phase === 'success') {
                setRowStatus(item.id, 'exported');
                completed++;
                progressBar.value = completed;
                progressTextSpan.textContent = `${completed}/${items.length}`;
            } else {
                setRowStatus(item.id, 'error', event.error);
                completed++;
                progressBar.value = completed;
                progressTextSpan.textContent = `${completed}/${items.length}`;
            }

            if (completed === items.length) {
                setTimeout(() => {
                    progressBarContainer.remove();
                }, 3000);
            }
        };

        return handler;
    }

    async function fetchChats() {
        const previousState = captureTableState();
        const tbody = document.getElementById('chatTableBody')!;
        tbody.innerHTML = '<tr><td colspan="5" style="color: #666; padding: 8px;">Loading…</td></tr>';

        try {
            const inputNumber = document.getElementById("conversation-fetch-list-max")! as HTMLInputElement;
            const n = inputNumber.valueAsNumber;
            const maxChats = isNaN(n) ? 15 : n;

            console.log(`${APP_TAG} Getting MSAL ids...`);
            const msalIds = getMsalIds();
            console.log(`${APP_TAG} Getting access token...`);
            const accessToken = await getAccessToken(msalIds);
            const copilotChatList = await fetchCopilotChats(accessToken, msalIds.localAccountId, msalIds.tenantId, maxChats);

            renderChatTable(copilotChatList.chats, previousState);
            updateSelectedCount();
        } catch {
            tbody.innerHTML = '<tr><td colspan="5" style="color: #dc2626; padding: 8px;">Failed to load conversations.</td></tr>';
            const selectAll = document.getElementById('selectAllCheckbox')! as HTMLInputElement;
            selectAll.checked = false;
            document.getElementById('selectedCount')!.textContent = '(0/0)';
        }
    }

    function getSelectedChats(): TransportObject[] {
        const checkboxes = document.querySelectorAll('#chatTableBody input[type="checkbox"]:checked') as NodeListOf<HTMLInputElement>;
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

    function getExportFormat(): ExportFormat {
        const select = document.getElementById("export-format-select")! as HTMLSelectElement;
        return select.value as ExportFormat;
    }

    function getOutputMode(): OutputMode {
        const select = document.getElementById("export-output-mode-select")! as HTMLSelectElement;
        return select.value as OutputMode;
    }

    function parseCopilotJsonFile(text: string): CopilotConversation[] {
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed)) {
            return parsed as CopilotConversation[];
        }
        return [parsed as CopilotConversation];
    }

    function showExportResultAlert(successCount: number, totalCount: number): void {
        if (successCount === totalCount) {
            alert(`Successfully exported ${successCount} of ${totalCount} conversations.`);
        } else {
            alert(`Exported ${successCount} of ${totalCount} conversations. Hover over red statuses for error details.`);
        }
    }

    function sanitizeFilename(name: string): string {
        const sanitized = name.replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").trim();
        return sanitized || "conversation";
    }

    async function exportChats() {
        const items = getSelectedChats();
        if (items.length === 0) return;

        clearRowStatus(items.map(i => i.id));

        const handler = createExportProgressHandler(items, "Exporting...");
        if (!handler) {
            return;
        }
        const result = await exportBulkDirect(
            items.map(i => i.id),
            handler,
            getExportFormat(),
            getOutputMode(),
        );
        showExportResultAlert(result.successCount, result.totalCount);
    }

    async function deleteChats() {
        const items = getSelectedChats();
        if (items.length === 0) return;

        const message = items.length === 1
            ? `Permanently delete "${items[0].title}"? This cannot be undone.`
            : `Permanently delete ${items.length} conversations? This cannot be undone.`;
        if (!confirm(message)) return;

        clearRowStatus(items.map(i => i.id));
        items.forEach(i => setRowStatus(i.id, 'deleting'));

        const progressUpdater = createProgressBar(items, "Deleting...");

        try {
            await deleteBulk(
                items.map(i => i.id),
                progressUpdater ?? (() => { }),
            );
            items.forEach(i => setRowStatus(i.id, 'deleted'));
        } catch (err) {
            const msg = err instanceof Error ? err.message : String(err);
            items.forEach(i => setRowStatus(i.id, 'error', msg));
        }
    }

    const selectAllCheckbox = document.getElementById('selectAllCheckbox')! as HTMLInputElement;
    selectAllCheckbox.addEventListener('change', () => {
        const checkboxes = document.querySelectorAll('#chatTableBody input[type="checkbox"]') as NodeListOf<HTMLInputElement>;
        checkboxes.forEach(cb => {
            cb.checked = selectAllCheckbox.checked;
        });
        updateSelectedCount();
    });

    const chatTableBody = document.getElementById('chatTableBody')!;
    chatTableBody.addEventListener('change', (e) => {
        if ((e.target as HTMLElement).matches('input[type="checkbox"]')) {
            updateSelectedCount();
        }
    });

    const exportBtn = document.getElementById("export-conversations-button")! as HTMLButtonElement;
    exportBtn.addEventListener("click", exportChats)

    const exportFormatSelect = document.getElementById("export-format-select")! as HTMLSelectElement;
    const exportOutputModeSelect = document.getElementById("export-output-mode-select")! as HTMLSelectElement;
    const combinedOutputOption = exportOutputModeSelect.querySelector('option[value="combined"]')! as HTMLOptionElement;

    exportFormatSelect.addEventListener("change", () => {
        if (exportFormatSelect.value === "markdown") {
            combinedOutputOption.disabled = true;
            if (exportOutputModeSelect.value === "combined") {
                exportOutputModeSelect.value = "zip";
            }
        } else {
            combinedOutputOption.disabled = false;
        }
    });

    const deleteBtn = document.getElementById("delete-conversations-button")! as HTMLButtonElement;
    deleteBtn.addEventListener("click", deleteChats)

    const refetchButton = document.getElementById("conversation-refetch")! as HTMLButtonElement;
    refetchButton.addEventListener("click", fetchChats);

    const fileInput = document.getElementById("copilot-json-upload")! as HTMLInputElement;
    const convertBtn = document.getElementById("convert-uploaded-button")! as HTMLButtonElement;
    const convertFormatSelect = document.getElementById("convert-format-select")! as HTMLSelectElement;

    convertBtn.addEventListener("click", () => fileInput.click());

    fileInput.addEventListener("change", async () => {
        const files = fileInput.files;
        if (!files || files.length === 0) return;

        const format = convertFormatSelect.value as Exclude<ExportFormat, "json">;
        const conversations: CopilotConversation[] = [];

        for (const file of files) {
            conversations.push(...parseCopilotJsonFile(await file.text()));
        }

        if (format === "chatgpt") {
            const converted = conversations.map(mapToConversationJson);
            const blob = new Blob([JSON.stringify(converted, null, 2)], { type: "application/json" });
            downloadBlobAsFile(blob, "conversations.json");
        } else {
            for (const conversation of conversations) {
                const blob = new Blob([mapToMarkdown(conversation)], { type: "text/markdown" });
                downloadBlobAsFile(blob, `${sanitizeFilename(conversation.chatName)}.md`);
            }
        }

        fileInput.value = "";
    });

    fetchChats();
}
