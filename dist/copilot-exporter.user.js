// ==UserScript==
// @name         M365 Copilot Exporter
// @namespace    ganyuke
// @version      1.2.1
// @author       ganyuke
// @description  An exporter for the Copilot Chat integrated into the M365 dashboard.
// @license      MIT
// @icon         https://upload.wikimedia.org/wikipedia/commons/0/0e/Microsoft_365_%282022%29.svg
// @source       https://github.com/ganyuke/copilot-exporter.git
// @match        https://m365.cloud.microsoft/
// @match        https://m365.cloud.microsoft/chat/*
// @match        https://m365.cloud.microsoft/chat
// @grant        GM.registerMenuCommand
// @run-at       document-end
// ==/UserScript==

(function() {
	"use strict";
	async function fetchCopilotChats(token, userOid, tenantId, maxChats, variants = "feature.EnableLastMessageForGetChats,feature.EnableMRUAgents,feature.EnableHasLoopPages") {
		const requestObj = {
			source: "officeweb",
			traceId: crypto.randomUUID(),
			threadType: "webchat",
			MaxReturnedChatsCount: maxChats
		};
		const url = `https://substrate.office.com/m365Copilot/GetChats?request=${encodeURIComponent(JSON.stringify(requestObj))}&variants=${encodeURIComponent(variants)}`;
		const headers = {
			"authorization": `Bearer ${token}`,
			"content-type": "application/json",
			"x-anchormailbox": `Oid:${userOid}@${tenantId}`,
			"x-clientrequestid": crypto.randomUUID().replace(/-/g, ""),
			"x-routingparameter-sessionkey": userOid,
			"x-scenario": "OfficeWebIncludedCopilot",
			"x-variants": variants
		};
		const res = await fetch(url, {
			method: "GET",
			headers
		});
		if (!res.ok) {
			console.debug(res);
			console.debug(res.body);
			throw new Error(`Fetch failed with status ${res.status}`);
		}
		return await res.json();
	}
	async function fetchCopilotConversation(token, userOid, tenantId, conversationId) {
		const requestObj = {
			conversationId,
			source: "officeweb",
			traceId: crypto.randomUUID().replace(/-/g, "")
		};
		const url = `https://substrate.office.com/m365Copilot/GetConversation?request=${encodeURIComponent(JSON.stringify(requestObj))}`;
		const headers = {
			"authorization": `Bearer ${token}`,
			"content-type": "application/json",
			"x-anchormailbox": `Oid:${userOid}@${tenantId}`,
			"x-clientrequestid": crypto.randomUUID().replace(/-/g, ""),
			"x-routingparameter-sessionkey": userOid,
			"x-scenario": "OfficeWebIncludedCopilot"
		};
		const response = await fetch(url, {
			method: "GET",
			headers
		});
		if (!response.ok) {
			console.debug(response);
			console.debug(response.body);
			throw new Error(`Fetch failed with status ${response.status}`);
		}
		return await response.blob();
	}
	async function deleteCopilotConversation(token, userOid, tenantId, conversationIds) {
		const requestObj = {
			conversationIdsToDelete: conversationIds,
			source: "officeweb",
			traceId: crypto.randomUUID()
		};
		const encodedRequest = JSON.stringify(requestObj);
		const url = `https://substrate.office.com/m365Copilot/DeleteConversation`;
		const headers = {
			"authorization": `Bearer ${token}`,
			"content-type": "application/json",
			"x-anchormailbox": `Oid:${userOid}@${tenantId}`,
			"x-clientrequestid": crypto.randomUUID(),
			"x-routingparameter-sessionkey": userOid,
			"x-scenario": "OfficeWebIncludedCopilot"
		};
		const response = await fetch(url, {
			method: "POST",
			headers,
			body: encodedRequest
		});
		if (!response.ok) {
			console.debug(response);
			console.debug(response.body);
			throw new Error(`Fetch failed with status ${response.status}`);
		}
	}
	function downloadBlobAsFile(blob, filename) {
		const url = URL.createObjectURL(blob);
		const a = document.createElement("a");
		a.href = url;
		a.download = filename;
		a.style.display = "none";
		document.body.appendChild(a);
		a.click();
		document.body.removeChild(a);
		URL.revokeObjectURL(url);
	}
	var getCookie = (key) => document.cookie.match(`(^|;)\\s*${key}\\s*=\\s*([^;]+)`)?.pop() || "";
	var ENCRYPTION_KEY = "msal.cache.encryption";
	var AES_GCM = "AES-GCM";
	var HKDF = "HKDF";
	var S256_HASH_ALG = "SHA-256";
	var RAW = "raw";
	var ENCRYPT = "encrypt";
	var DECRYPT = "decrypt";
	var DERIVE_KEY = "deriveKey";
	function base64DecToArr(base64String) {
		let encodedString = base64String.replace(/-/g, "+").replace(/_/g, "/");
		switch (encodedString.length % 4) {
			case 0: break;
			case 2:
				encodedString += "==";
				break;
			case 3:
				encodedString += "=";
				break;
			default: throw Error(`${APP_TAG} Error extracting base64`);
		}
		const binString = atob(encodedString);
		return Uint8Array.from(binString, (m) => m.codePointAt(0) || 0);
	}
	function toArrayBuffer(bufferLike) {
		return Uint8Array.from(bufferLike).buffer;
	}
	async function deriveKey(baseKey, nonce, context) {
		return window.crypto.subtle.deriveKey({
			name: HKDF,
			salt: toArrayBuffer(nonce),
			hash: S256_HASH_ALG,
			info: new TextEncoder().encode(context)
		}, baseKey, {
			name: AES_GCM,
			length: 256
		}, false, [ENCRYPT, DECRYPT]);
	}
	async function decrypt(baseKey, nonce, context, encryptedData) {
		const encodedData = base64DecToArr(encryptedData);
		const derivedKey = await deriveKey(baseKey, base64DecToArr(nonce), context);
		const decryptedData = await window.crypto.subtle.decrypt({
			name: AES_GCM,
			iv: new Uint8Array(12)
		}, derivedKey, toArrayBuffer(encodedData));
		return new TextDecoder().decode(decryptedData);
	}
	function generateHKDF(baseKey) {
		return window.crypto.subtle.importKey(RAW, toArrayBuffer(baseKey), HKDF, false, [DERIVE_KEY]);
	}
	async function getEncryptionCookie() {
		const cookieString = decodeURIComponent(getCookie(ENCRYPTION_KEY));
		let parsedCookie = {
			key: "",
			id: ""
		};
		if (cookieString) try {
			parsedCookie = JSON.parse(cookieString);
		} catch (e) {
			throw Error(`${APP_TAG} Failed to parse encryption cookie`);
		}
		if (parsedCookie.key && parsedCookie.id) {
			const baseKey = base64DecToArr(parsedCookie.key);
			return {
				id: parsedCookie.id,
				key: await generateHKDF(baseKey)
			};
		} else throw Error(`${APP_TAG} No encryption cookie found`);
	}
	var getMsalIds = () => {
		const clientId = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
		const msalIds = localStorage.getItem("msal.3.account.keys");
		if (!msalIds) throw Error(`${APP_TAG} No account keys found for Copilot application`);
		const accountKeys = JSON.parse(msalIds);
		if (accountKeys.length === 0) throw Error(`${APP_TAG} No account keys found for Copilot application`);
		const [homeAccountId, _1, tenantId] = accountKeys[0].split("|");
		const [localAccountId, _2] = homeAccountId.split(".");
		return {
			localAccountId,
			tenantId,
			homeAccountId,
			clientId
		};
	};
	var getAccessToken = async (msalIds) => {
		const encryptionCookie = await getEncryptionCookie();
		const tokenKeys = localStorage.getItem(`msal.3.token.keys.${msalIds.clientId}`);
		if (!tokenKeys) throw Error(`${APP_TAG} No token keys found for Copilot application`);
		const sydneyKey = JSON.parse(tokenKeys).accessToken.find((token) => token.includes("https://substrate.office.com/sydney/.default"));
		if (!sydneyKey) throw Error(`${APP_TAG} No Sydney access token found for Copilot application`);
		const sydneyTokenEntry = localStorage.getItem(sydneyKey);
		if (!sydneyTokenEntry) throw Error(`${APP_TAG} No Sydney token found for Copilot application`);
		const payload = JSON.parse(sydneyTokenEntry);
		const decryptedData = await decrypt(encryptionCookie.key, payload.nonce, msalIds.clientId, payload.data);
		return JSON.parse(decryptedData).secret;
	};
	var FETCH_DELAY = 1500;
	async function getTokenAndIds() {
		console.log(`${APP_TAG} Getting MSAL ids...`);
		const msalIds = getMsalIds();
		console.log(`${APP_TAG} Getting access token...`);
		return {
			token: await getAccessToken(msalIds),
			...msalIds
		};
	}
	async function exportBulkDirect(conversationIds, callback) {
		const { token, localAccountId, tenantId } = await getTokenAndIds();
		for (let i = 0; i < conversationIds.length; i++) {
			const conversationId = conversationIds[i];
			const blob = await fetchCopilotConversation(token, localAccountId, tenantId, conversationId);
			console.log(`${APP_TAG} Completed download for conversation ${conversationId}`);
			callback(i);
			downloadBlobAsFile(blob, `m365-copilot-${conversationId}.json`);
			await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
		}
	}
	async function deleteBulk(conversationIds, callback) {
		const { token, localAccountId, tenantId } = await getTokenAndIds();
		await deleteCopilotConversation(token, localAccountId, tenantId, conversationIds);
		callback(conversationIds.length - 1);
		console.log(`${APP_TAG} Completed deletion for conversations ${conversationIds.join()}`);
	}
	function showExportModal() {
		if (document.getElementById("copilotExportOverlay")) return;
		const overlay = document.createElement("div");
		overlay.id = "copilotExportOverlay";
		overlay.style.cssText = `
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.5);
    display: flex; align-items: center; justify-content: center;
    z-index: 9999;
  `;
		overlay.addEventListener("click", () => {
			overlay.remove();
		});
		const modal = document.createElement("div");
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
		function createProgressBar(idsToExport, initalString) {
			if (idsToExport.length < 1) return;
			if (document.querySelector("#chat-export-progress-bar-container")) return;
			const progressBarContainer = document.createElement("div");
			progressBarContainer.id = "chat-export-progress-bar-container";
			progressBarContainer.style = "display: flex;flex-direction: column;margin-top: 0.5em;";
			const progressBar = document.createElement("progress");
			const label = document.createElement("label");
			label.style = "display:flex;";
			const titleSpan = document.createElement("span");
			const progressTextSpan = document.createElement("span");
			titleSpan.style = "flex-grow:1;";
			progressBar.id = "chat-export-progress-bar";
			progressBar.max = idsToExport.length;
			progressBar.value = 0;
			label.htmlFor = "chat-export-progress-bar";
			titleSpan.textContent = initalString;
			progressTextSpan.textContent = `0/${idsToExport.length}`;
			label.append(titleSpan, progressTextSpan);
			progressBarContainer.append(label, progressBar);
			modal.append(progressBarContainer);
			const progressUpdater = (progress) => {
				titleSpan.textContent = idsToExport[progress].title;
				progressTextSpan.textContent = `${progress + 1}/${idsToExport.length}`;
				progressBar.value = progress + 1;
				if (progressBar.value === progressBar.max) setTimeout(() => {
					progressBarContainer.remove();
				}, 3e3);
			};
			return progressUpdater;
		}
		async function fetchChats() {
			const n = document.getElementById("conversation-fetch-list-max").valueAsNumber;
			const maxChats = isNaN(n) ? 15 : n;
			console.log(`${APP_TAG} Getting MSAL ids...`);
			const msalIds = getMsalIds();
			console.log(`${APP_TAG} Getting access token...`);
			const copilotChatList = await fetchCopilotChats(await getAccessToken(msalIds), msalIds.localAccountId, msalIds.tenantId, maxChats);
			const chatList = document.getElementById("chatList");
			chatList.innerText = "";
			copilotChatList.chats.forEach((data) => {
				const label = document.createElement("label");
				label.style = "column-gap:0.5em;display:flex;";
				const checkbox = document.createElement("input");
				const span = document.createElement("span");
				checkbox.type = "checkbox";
				checkbox.dataset["id"] = data.conversationId;
				checkbox.dataset["title"] = data.chatName;
				span.innerText = data.chatName;
				label.append(checkbox);
				label.append(span);
				chatList.appendChild(label);
			});
			const selectAll = document.getElementById("selectAllCheckbox");
			selectAll.addEventListener("change", () => {
				document.querySelectorAll("#chatList input[type=\"checkbox\"]").forEach((cb) => {
					cb.checked = selectAll.checked;
				});
			});
		}
		function getSelectedChats() {
			const checkboxes = document.querySelectorAll("#chatList input[type=\"checkbox\"]:checked");
			const listToExport = [];
			checkboxes.forEach((c) => {
				const uuid = c.dataset["id"];
				const title = c.dataset["title"];
				listToExport.push({
					id: uuid,
					title
				});
			});
			return listToExport;
		}
		function exportChats() {
			const idsToExport = getSelectedChats();
			const progressUpdater = createProgressBar(idsToExport, "Exporting...");
			if (!progressUpdater) return;
			exportBulkDirect(idsToExport.map((obj) => obj.id), progressUpdater);
		}
		function deleteChats() {
			const idsToDelete = getSelectedChats();
			const progressUpdater = createProgressBar(idsToDelete, "Deleting...");
			if (!progressUpdater) return;
			deleteBulk(idsToDelete.map((obj) => obj.id), progressUpdater);
		}
		document.getElementById("export-conversations-button").addEventListener("click", exportChats);
		document.getElementById("delete-conversations-button").addEventListener("click", deleteChats);
		document.getElementById("conversation-refetch").addEventListener("click", fetchChats);
		fetchChats();
	}
	var APP_TAG = "[Copilot Exporter]";
	console.log(`${APP_TAG} Userscript initalized.`);
	var EXPORT_SVG = `<svg width="100%" height="100%" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
<path d="M12 5L11.2929 4.29289L12 3.58579L12.7071 4.29289L12 5ZM13 14C13 14.5523 12.5523 15 12 15C11.4477 15 11 14.5523 11 14L13 14ZM6.29289 9.29289L11.2929 4.29289L12.7071 5.70711L7.70711 10.7071L6.29289 9.29289ZM12.7071 4.29289L17.7071 9.29289L16.2929 10.7071L11.2929 5.70711L12.7071 4.29289ZM13 5L13 14L11 14L11 5L13 5Z" fill="#33363F"/>
<path d="M5 16L5 17C5 18.1046 5.89543 19 7 19L17 19C18.1046 19 19 18.1046 19 17V16" stroke="#33363F" stroke-width="2"/>
</svg>`;
	var BUTTON_ID = "export-menu-button";
	var inject = () => {
		if (document.getElementById(BUTTON_ID)) return;
		const btn = document.createElement("button");
		const svgEl = new DOMParser().parseFromString(EXPORT_SVG, "image/svg+xml").documentElement;
		const svg = document.importNode(svgEl, true);
		btn.id = BUTTON_ID;
		btn.style.width = "3em";
		btn.style.height = "3em";
		btn.style.bottom = "16px";
		btn.style.right = "16px";
		btn.style.cursor = "pointer";
		btn.style.position = "fixed";
		btn.append(svg);
		btn.addEventListener("click", showExportModal);
		document.body.appendChild(btn);
	};
	GM.registerMenuCommand("Open export menu", showExportModal);
	if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", inject);
	else inject();
})();
