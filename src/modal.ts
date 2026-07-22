import { CopilotConversationOverview, fetchCopilotChats } from "./api";
import type { CopilotConversation } from "./api";
import { downloadBlobAsFile } from "./blob";
import { mapToConversationJson } from "./converters/chatgpt";
import { mapToMarkdown } from "./converters/markdown";
import { deleteBulk, exportBulkDirect, ExportCallback, ExportFormat, OutputMode } from "./expoter";
import { APP_TAG } from "./main";
import { getAccessToken, getMsalIds } from "./token";
import { version, author, repository } from '../package.json' with { type: 'json' };
import MODAL_STYLES from './assets/styles.css?inline';

type TransportObject = {
    id: string;
    title: string;
}

type RowStatus = 'exporting' | 'exported' | 'deleting' | 'deleted' | 'error';

const STATUS_LABELS: Record<RowStatus, string> = {
    exporting: 'exporting…',
    exported: 'exported',
    deleting: 'deleting…',
    deleted: 'deleted',
    error: 'error',
};

function waitForNextPaint(): Promise<void> {
    return new Promise((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
    });
}

function setRowStatusOnCell(cell: HTMLTableCellElement, status: RowStatus, error?: string): void {
    cell.textContent = STATUS_LABELS[status];
    cell.className = 'status-cell';
    cell.dataset.status = status;
    if (status === 'error' && error) {
        cell.title = error;
    } else {
        cell.removeAttribute('title');
    }
}

export function showExportModal() {
    if (document.getElementById('copilotExportOverlay')) return;

    const overlay = document.createElement('div');
    overlay.id = 'copilotExportOverlay';

    overlay.addEventListener("click", () => {
        overlay.remove();
    })

    const style = document.createElement('style');
    style.textContent = MODAL_STYLES;
    overlay.appendChild(style);

    const modal = document.createElement('div');
    modal.id = 'copilotExportModal';

    modal.addEventListener("click", (e) => {
        e.stopPropagation();
    });

    modal.innerHTML = `
    <h2 id="copilotExportTitle">Export conversations</h2>
    <p id="copilotExportByline"><a href="${repository.url}" target="_blank">M365 Copilot Exporter</a> v${version} by <a href="${author.url}" target="_blank">${author.name}</a></p>

    <div id="chatTableContainer">
      <div id="chatTableToolbar">
        <div id="toolbarSelectRow">
          <label><input type="checkbox" id="selectAllCheckbox"> Select All</label>
          <span id="selectedCount">(0/0)</span>
        </div>
        <div id="toolbarFetchRow">
          <label for="conversation-fetch-list-max">Max conversations</label>
          <input type="number" id="conversation-fetch-list-max" name="quantity" min="0" placeholder="15">
          <button id="conversation-refetch">Refetch</button>
        </div>
      </div>
      <div id="chatTableScroll">
      <table id="chatTable">
        <colgroup>
          <col>
          <col>
          <col>
          <col>
          <col>
        </colgroup>
        <thead>
          <tr>
            <th></th>
            <th>Name</th>
            <th>Created</th>
            <th>Updated</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody id="chatTableBody">
          <tr><td colspan="5" class="placeholder">Loading…</td></tr>
        </tbody>
      </table>
      </div>
    </div>

    <div id="exportActions">
      <div id="exportActionsFormats">
        <select id="export-format-select">
          <option value="json">Copilot JSON</option>
          <option value="markdown">Markdown</option>
          <option value="chatgpt">ChatGPT JSON</option>
        </select>
        <select id="export-output-mode-select">
          <option value="individual">Individual files</option>
          <option value="combined">Combined file</option>
          <option value="zip" selected>Individual files (ZIP)</option>
        </select>
      </div>
      <div id="exportActionsButtons">
        <button id="delete-conversations-button">Delete</button>
        <button id="export-conversations-button">Export</button>
      </div>
    </div>

    <div id="convertSection">
      <p id="convertHelp">
        Re-import and convert exported Copilot JSON files to other formats.
      </p>
      <input type="file" id="copilot-json-upload" accept=".json,application/json" multiple hidden>
      <select id="convert-format-select">
        <option value="chatgpt">ChatGPT JSON</option>
        <option value="markdown">Markdown</option>
      </select>
      <button id="convert-uploaded-button">Open and convert Copilot JSON</button>
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
        setRowStatusOnCell(cell, status, error);
    }

    function clearRowStatus(conversationIds: string[]): void {
        for (const id of conversationIds) {
            const cell = findStatusCell(id);
            if (!cell) continue;
            cell.textContent = '';
            cell.className = 'status-cell';
            delete cell.dataset.status;
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
        status: RowStatus | null;
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
                status: statusCell ? statusFromCell(statusCell) : null,
                statusTitle: statusCell?.getAttribute('title') ?? null,
            });
        }

        return state;
    }

    function statusFromCell(cell: HTMLTableCellElement): RowStatus | null {
        const status = cell.dataset.status;
        if (status && status in STATUS_LABELS) {
            return status as RowStatus;
        }
        return null;
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
            cell.className = 'placeholder';
            row.appendChild(cell);
            tbody.appendChild(row);
        } else {
            for (const data of sorted) {
                const row = document.createElement('tr');
                row.setAttribute('data-conversation-id', data.conversationId);
                row.setAttribute('data-chat-name', data.chatName);
                row.setAttribute('data-create-time', String(data.createTimeUtc));
                row.setAttribute('data-update-time', String(data.updateTimeUtc));

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
                nameTd.title = data.chatName;
                nameTd.textContent = data.chatName;

                const createdTd = document.createElement('td');
                createdTd.className = 'date-cell';
                createdTd.textContent = formatPrettyDate(data.createTimeUtc);

                const updatedTd = document.createElement('td');
                updatedTd.className = 'date-cell';
                updatedTd.textContent = formatPrettyDate(data.updateTimeUtc);

                const statusTd = document.createElement('td');
                statusTd.className = 'status-cell';

                if (previous && chatDataMatches(previous, data) && previous.status) {
                    setRowStatusOnCell(statusTd, previous.status, previous.statusTitle ?? undefined);
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
        progressBarContainer.id = "chat-export-progress-bar-container";
        const progressBar = document.createElement("progress") as HTMLProgressElement;
        const label = document.createElement("label") as HTMLLabelElement;
        label.id = "chat-export-progress-label";
        const titleSpan = document.createElement("span") as HTMLSpanElement;
        const progressTextSpan = document.createElement("span") as HTMLSpanElement;
        titleSpan.id = "chat-export-progress-title";
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
        progressBarContainer.id = "chat-export-progress-bar-container";
        const progressBar = document.createElement("progress") as HTMLProgressElement;
        const label = document.createElement("label") as HTMLLabelElement;
        label.id = "chat-export-progress-label";
        const titleSpan = document.createElement("span") as HTMLSpanElement;
        const progressTextSpan = document.createElement("span") as HTMLSpanElement;
        titleSpan.id = "chat-export-progress-title";
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
        tbody.innerHTML = '<tr><td colspan="5" class="placeholder">Loading…</td></tr>';

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
            tbody.innerHTML = '<tr><td colspan="5" class="error-text">Failed to load conversations.</td></tr>';
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
            // need to wait a frame so that progress gets updated before the confirm dialog appears
            await waitForNextPaint();
            const refetchMessage =  `Deleted ${items.length} conversation${items.length === 1 ? '' : 's'}. Refetch now to update the list?`;
            if (confirm(refetchMessage)) {
                await fetchChats();
            }
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
