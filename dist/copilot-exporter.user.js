// ==UserScript==
// @name         M365 Copilot Exporter
// @namespace    ganyuke
// @version      2.0.1
// @author       ganyuke
// @description  View, bulk delete, and export your Microsoft 365 Copilot Chat conversations into raw JSON, readable Markdown, or ChatGPT's conversation.json format.
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
	var __create = Object.create;
	var __defProp = Object.defineProperty;
	var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
	var __getOwnPropNames = Object.getOwnPropertyNames;
	var __getProtoOf = Object.getPrototypeOf;
	var __hasOwnProp = Object.prototype.hasOwnProperty;
	var __commonJSMin = (cb, mod) => () => (mod || (cb((mod = { exports: {} }).exports, mod), cb = null), mod.exports);
	var __copyProps = (to, from, except, desc) => {
		if (from && typeof from === "object" || typeof from === "function") for (var keys = __getOwnPropNames(from), i = 0, n = keys.length, key; i < n; i++) {
			key = keys[i];
			if (!__hasOwnProp.call(to, key) && key !== except) __defProp(to, key, {
				get: ((k) => from[k]).bind(null, key),
				enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable
			});
		}
		return to;
	};
	var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", {
		value: mod,
		enumerable: true
	}) : target, mod));
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
		const data = await res.json();
		console.log(data);
		return data;
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
	function millisToSeconds(ms) {
		return typeof ms === "number" ? ms / 1e3 : null;
	}
	function isoToSeconds(iso) {
		if (!iso) return null;
		const ms = Date.parse(iso);
		return Number.isFinite(ms) ? ms / 1e3 : null;
	}
	function makeUuid() {
		return crypto.randomUUID();
	}
	function roleFromSource(author) {
		return author === "user" ? "user" : "assistant";
	}
	function makeRootNode(id) {
		return {
			id,
			message: null,
			parent: null,
			children: []
		};
	}
	function makeExportMessage(source) {
		const role = roleFromSource(source.author);
		return {
			id: source.messageId,
			author: {
				role,
				name: null,
				metadata: {}
			},
			create_time: isoToSeconds(source.createdAt ?? source.timestamp),
			update_time: null,
			content: {
				content_type: "text",
				parts: [source.text ?? ""]
			},
			status: "finished_successfully",
			end_turn: role === "assistant" ? true : null,
			weight: 1,
			metadata: {},
			recipient: "all",
			channel: null
		};
	}
	function makeExportNode(source, parentId) {
		return {
			id: source.messageId,
			message: makeExportMessage(source),
			parent: parentId,
			children: []
		};
	}
	function mapToConversationJson(source) {
		const rootId = makeUuid();
		const mapping = { [rootId]: makeRootNode(rootId) };
		let parentId = rootId;
		let currentNode = rootId;
		for (const message of source.messages) {
			const nodeId = message.messageId;
			mapping[nodeId] = makeExportNode(message, parentId);
			mapping[parentId].children.push(nodeId);
			parentId = nodeId;
			currentNode = nodeId;
		}
		return {
			title: source.chatName,
			create_time: millisToSeconds(source.createTimeUtc),
			update_time: millisToSeconds(source.updateTimeUtc),
			mapping,
			moderation_results: [],
			current_node: currentNode,
			plugin_ids: null,
			conversation_id: source.conversationId,
			conversation_template_id: null,
			gizmo_id: null,
			gizmo_type: null,
			is_archived: false,
			is_starred: null,
			safe_urls: [],
			default_model_slug: null,
			disabled_tool_ids: [],
			id: source.conversationId
		};
	}
	var SOURCE_URL = "https://m365.cloud.microsoft/chat/conversation";
	var CITATION_MARKER_RE = /【([^】]+)】/g;
	function millisToIsoOffset(ms) {
		return new Date(ms).toISOString().replace("Z", "+00:00");
	}
	function parseMessageDate(iso) {
		if (!iso) return null;
		const ms = Date.parse(iso);
		return Number.isFinite(ms) ? new Date(ms) : null;
	}
	function formatTimeElement(date) {
		return `<time datetime="${date.toISOString()}" title="${formatTimeTitle(date)}">${formatTimeShort(date)}</time>`;
	}
	function speakerLabel(author) {
		return author === "user" ? "You" : "Copilot";
	}
	function formatTimeTitle(date) {
		return new Intl.DateTimeFormat("en-US", {
			dateStyle: "short",
			timeStyle: "medium"
		}).format(date);
	}
	function formatTimeShort(date) {
		return new Intl.DateTimeFormat("en-US", {
			hour: "numeric",
			minute: "2-digit"
		}).format(date);
	}
	function parseCitationContent(content) {
		if (!content) return null;
		try {
			return JSON.parse(content);
		} catch {
			return null;
		}
	}
	function resolveMessageBody(message) {
		const rawText = message.text ?? "";
		try {
			const cardText = message.adaptiveCards?.[0]?.body?.[0]?.text;
			if (typeof cardText !== "string") return rawText;
			const references = message.references ?? {};
			const used = new Map();
			const rewritten = cardText.replace(CITATION_MARKER_RE, (_match, key) => {
				const ref = references[key];
				if (!ref?.targetLink) throw new Error(`missing reference for ${key}`);
				const parsed = parseCitationContent(ref.displayData?.content);
				const n = parsed?.label;
				const title = parsed?.Title;
				if (!n || !title) throw new Error(`incomplete citation metadata for ${key}`);
				if (!used.has(key)) used.set(key, {
					n,
					title,
					url: ref.targetLink
				});
				return `[${n}]`;
			});
			if (used.size === 0) return rewritten;
			return `${rewritten}\n\n**Sources:**\n\n${[...used.values()].sort((a, b) => Number(a.n) - Number(b.n)).map(({ n, title, url }) => `[${n}] [${title}](${url})`).join("\n")}`;
		} catch {
			return rawText;
		}
	}
	function formatMessage(message) {
		const date = parseMessageDate(message.createdAt ?? message.timestamp);
		const lines = [`## ${speakerLabel(message.author)}:`];
		if (date) lines.push(formatTimeElement(date));
		lines.push("", resolveMessageBody(message));
		return lines.join("\n");
	}
	function mapToMarkdown(source) {
		return `${[
			"---",
			`title: ${JSON.stringify(source.chatName)}`,
			`createdAt: ${JSON.stringify(millisToIsoOffset(source.createTimeUtc))}`,
			`updatedAt: ${JSON.stringify(millisToIsoOffset(source.updateTimeUtc))}`,
			`source: ${SOURCE_URL}/${source.conversationId}`,
			"---"
		].join("\n")}\n\n${[`# ${source.chatName}`, ...source.messages.map((message) => formatMessage(message))].join("\n\n")}\n`;
	}
	function createZip(files, opts, next) {
		if (typeof opts == "function") {
			next = opts;
			opts = {};
		}
		var i = 256, j, k, offset = 0, crcTable = [], cd = "", out = [], outLen = 0, CompressionStream = window.CompressionStream, now = Date.now(), push = (arr) => {
			out.push(arr);
			outLen += arr.length;
		}, dosDate = (date) => Math.max(2162688, date.getSeconds() >> 1 | date.getMinutes() << 5 | date.getHours() << 11 | date.getDate() << 16 | date.getMonth() + 1 << 21 | date.getFullYear() - 1980 << 25), le16 = (n) => String.fromCharCode(n & 255, n >>> 8 & 255), le32 = (n) => le16(n) + le16(n >>> 16), toUint = (str) => {
			for (var pos = str.length, arr = new Uint8Array(pos); pos--; arr[pos] = str.charCodeAt(pos));
			return arr;
		}, toUtf8 = (str) => unescape(encodeURIComponent(str || "")), compress = (uint, len, cb) => {
			if (opts && opts.deflate) {
				var compressed = opts.deflate(uint);
				cb(len > compressed.length ? compressed : uint);
			} else if (CompressionStream) new Response(new Blob([uint]).stream().pipeThrough(new CompressionStream("deflate"))).arrayBuffer().then((arr) => {
				cb(len > arr.byteLength - 6 ? new Uint8Array(arr).subarray(2, -4) : uint);
			});
			else cb(uint);
		}, add = (resolve) => {
			k = files[i++];
			if (!k) {
				k = files.length;
				name = toUtf8(opts && opts.comment);
				push(toUint(cd + "PK" + le32(0) + le32((k << 16) + k) + le32(cd.length) + le32(offset) + le16(name.length) + name));
				file = new Uint8Array(outLen);
				for (i = 0, offset = 0; j = out[i++]; offset += j.length) file.set(j, offset);
				return resolve(file);
			}
			var fileLen, name = toUtf8(k.name), nameLen = name.length, file = k.content, crc = -1;
			if (typeof file === "string") file = toUint(toUtf8(file));
			fileLen = file.length;
			for (j = 0; j < fileLen;) crc = crc >>> 8 ^ crcTable[(crc ^ file[j++]) & 255];
			compress(file, fileLen, (compressed, method) => {
				method = file === compressed ? "\0\0" : "\b\0";
				method = le32(134217748) + method + le32(dosDate(new Date(k.time > 0 || k.time === 0 ? k.time : now))) + le32(-1 ^ crc >>> 0) + le32(compressed.length) + le32(fileLen) + le32(nameLen);
				push(toUint("PK" + method + name));
				push(compressed);
				cd += "PK\0" + method + "\0\0" + le32(0) + le32(32) + le32(offset) + name;
				offset += 30 + compressed.length + nameLen;
				add(resolve);
			});
		};
		for (; i; crcTable[i] = k) {
			k = --i;
			for (j = 8; j--;) k = 3988292384 * (1 & k) ^ k >>> 1;
		}
		if (next) add(next.bind(next, null));
		else if (opts && opts.deflate) {
			add((file) => out = file);
			return out;
		} else return new Promise(add);
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
	var require_truncate = __commonJSMin(((exports, module) => {
		function isHighSurrogate(codePoint) {
			return codePoint >= 55296 && codePoint <= 56319;
		}
		function isLowSurrogate(codePoint) {
			return codePoint >= 56320 && codePoint <= 57343;
		}
		module.exports = function truncate(getLength, string, byteLength) {
			if (typeof string !== "string") throw new Error("Input must be string");
			var charLength = string.length;
			var curByteLength = 0;
			var codePoint;
			var segment;
			for (var i = 0; i < charLength; i += 1) {
				codePoint = string.charCodeAt(i);
				segment = string[i];
				if (isHighSurrogate(codePoint) && isLowSurrogate(string.charCodeAt(i + 1))) {
					i += 1;
					segment += string[i];
				}
				curByteLength += getLength(segment);
				if (curByteLength === byteLength) return string.slice(0, i + 1);
				else if (curByteLength > byteLength) return string.slice(0, i - segment.length + 1);
			}
			return string;
		};
	}));
	var require_browser$1 = __commonJSMin(((exports, module) => {
		function isHighSurrogate(codePoint) {
			return codePoint >= 55296 && codePoint <= 56319;
		}
		function isLowSurrogate(codePoint) {
			return codePoint >= 56320 && codePoint <= 57343;
		}
		module.exports = function getByteLength(string) {
			if (typeof string !== "string") throw new Error("Input must be string");
			var charLength = string.length;
			var byteLength = 0;
			var codePoint = null;
			var prevCodePoint = null;
			for (var i = 0; i < charLength; i++) {
				codePoint = string.charCodeAt(i);
				if (isLowSurrogate(codePoint)) if (prevCodePoint != null && isHighSurrogate(prevCodePoint)) byteLength += 1;
				else byteLength += 3;
				else if (codePoint <= 127) byteLength += 1;
				else if (codePoint >= 128 && codePoint <= 2047) byteLength += 2;
				else if (codePoint >= 2048 && codePoint <= 65535) byteLength += 3;
				prevCodePoint = codePoint;
			}
			return byteLength;
		};
	}));
	var require_browser = __commonJSMin(((exports, module) => {
		var truncate = require_truncate();
		var getLength = require_browser$1();
		module.exports = truncate.bind(null, getLength);
	}));
	var import_sanitize_filename = __toESM(__commonJSMin(((exports, module) => {
		var truncate = require_browser();
		var illegalRe = /[\/\?<>\\:\*\|"]/g;
		var controlRe = /[\x00-\x1f\x80-\x9f]/g;
		var reservedRe = /^\.+$/;
		var windowsReservedRe = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])(\..*)?$/i;
		function replaceTrailingDotsAndSpaces(str, replacement) {
			var end = str.length;
			while (end > 0 && (str[end - 1] === "." || str[end - 1] === " ")) end--;
			return end < str.length ? str.slice(0, end) + replacement : str;
		}
		function sanitize(input, replacement) {
			if (typeof input !== "string") throw new Error("Input must be string");
			var sanitized = input.replace(illegalRe, replacement).replace(controlRe, replacement).replace(reservedRe, replacement).replace(windowsReservedRe, replacement);
			sanitized = replaceTrailingDotsAndSpaces(sanitized, replacement);
			return truncate(sanitized, 255);
		}
		module.exports = function(input, options) {
			var replacement = options && options.replacement || "";
			var output = sanitize(input, replacement);
			if (replacement === "") return output;
			return sanitize(output, "");
		};
	}))(), 1);
	var FETCH_DELAY = 1500;
	function sanitizeFilename(name) {
		return (0, import_sanitize_filename.default)(name, { replacement: "_" });
	}
	function exportFilename(conversation, format) {
		const base = sanitizeFilename(conversation.chatName) || conversation.conversationId;
		switch (format) {
			case "markdown": return `m365-copilot-${base}.md`;
			case "chatgpt": return `m365-copilot-as-chatgpt-${base}.json`;
			default: return `m365-copilot-${base}.json`;
		}
	}
	function zipArchiveFilename(format) {
		switch (format) {
			case "markdown": return "m365-copilot-export-markdown.zip";
			case "chatgpt": return "m365-copilot-export-chatgpt-json.zip";
			default: return "m365-copilot-export-json.zip";
		}
	}
	function combinedArchiveFilename(format) {
		switch (format) {
			case "chatgpt": return "conversations.json";
			default: return "m365-copilot-conversations.json";
		}
	}
	function conversationToBlob(conversation, format) {
		switch (format) {
			case "markdown": return new Blob([mapToMarkdown(conversation)], { type: "text/markdown" });
			case "chatgpt": return new Blob([JSON.stringify(mapToConversationJson(conversation), null, 2)], { type: "application/json" });
			default: return new Blob([JSON.stringify(conversation, null, 2)], { type: "application/json" });
		}
	}
	function conversationToCombinedEntry(conversation, format) {
		switch (format) {
			case "chatgpt": return mapToConversationJson(conversation);
			default: return conversation;
		}
	}
	async function getTokenAndIds() {
		console.log(`${APP_TAG} Getting MSAL ids...`);
		const msalIds = getMsalIds();
		console.log(`${APP_TAG} Getting access token...`);
		return {
			token: await getAccessToken(msalIds),
			...msalIds
		};
	}
	async function fetchConversation(token, localAccountId, tenantId, conversationId) {
		const blob = await fetchCopilotConversation(token, localAccountId, tenantId, conversationId);
		return JSON.parse(await blob.text());
	}
	async function exportIndividual(conversationIds, callback, format, token, localAccountId, tenantId) {
		let successCount = 0;
		for (let i = 0; i < conversationIds.length; i++) {
			const conversationId = conversationIds[i];
			callback({
				phase: "start",
				index: i
			});
			try {
				const conversation = await fetchConversation(token, localAccountId, tenantId, conversationId);
				downloadBlobAsFile(conversationToBlob(conversation, format), exportFilename(conversation, format));
				console.log(`${APP_TAG} Completed download for conversation ${conversationId}`);
				callback({
					phase: "success",
					index: i
				});
				successCount++;
			} catch (err) {
				callback({
					phase: "error",
					index: i,
					error: err instanceof Error ? err.message : String(err)
				});
			}
			await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
		}
		return {
			successCount,
			totalCount: conversationIds.length
		};
	}
	async function exportCombined(conversationIds, callback, format, token, localAccountId, tenantId) {
		const combined = [];
		let successCount = 0;
		for (let i = 0; i < conversationIds.length; i++) {
			const conversationId = conversationIds[i];
			callback({
				phase: "start",
				index: i
			});
			try {
				const conversation = await fetchConversation(token, localAccountId, tenantId, conversationId);
				combined.push(conversationToCombinedEntry(conversation, format));
				console.log(`${APP_TAG} Completed export for conversation ${conversationId}`);
				callback({
					phase: "success",
					index: i
				});
				successCount++;
			} catch (err) {
				callback({
					phase: "error",
					index: i,
					error: err instanceof Error ? err.message : String(err)
				});
			}
			await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
		}
		if (combined.length > 0) downloadBlobAsFile(new Blob([JSON.stringify(combined, null, 2)], { type: "application/json" }), combinedArchiveFilename(format));
		return {
			successCount,
			totalCount: conversationIds.length
		};
	}
	async function exportZip(conversationIds, callback, format, token, localAccountId, tenantId) {
		let successCount = 0;
		const files = [];
		for (let i = 0; i < conversationIds.length; i++) {
			const conversationId = conversationIds[i];
			callback({
				phase: "start",
				index: i
			});
			try {
				const conversation = await fetchConversation(token, localAccountId, tenantId, conversationId);
				const blob = conversationToBlob(conversation, format);
				files.push({
					name: exportFilename(conversation, format),
					content: new TextEncoder().encode(await blob.text()),
					time: conversation.createTimeUtc
				});
				console.log(`${APP_TAG} Completed export for conversation ${conversationId}`);
				callback({
					phase: "success",
					index: i
				});
				successCount++;
			} catch (err) {
				callback({
					phase: "error",
					index: i,
					error: err instanceof Error ? err.message : String(err)
				});
			}
			await new Promise((resolve) => setTimeout(resolve, FETCH_DELAY));
		}
		if (successCount > 0) {
			const zipUint8Array = await createZip(files);
			downloadBlobAsFile(new Blob([zipUint8Array], { type: "application/zip" }), zipArchiveFilename(format));
		}
		return {
			successCount,
			totalCount: conversationIds.length
		};
	}
	async function exportBulkDirect(conversationIds, callback, format = "json", outputMode = "individual") {
		const { token, localAccountId, tenantId } = await getTokenAndIds();
		switch (outputMode) {
			case "combined": return exportCombined(conversationIds, callback, format, token, localAccountId, tenantId);
			case "zip": return exportZip(conversationIds, callback, format, token, localAccountId, tenantId);
			default: return exportIndividual(conversationIds, callback, format, token, localAccountId, tenantId);
		}
	}
	async function deleteBulk(conversationIds, callback) {
		const { token, localAccountId, tenantId } = await getTokenAndIds();
		await deleteCopilotConversation(token, localAccountId, tenantId, conversationIds);
		callback(conversationIds.length - 1);
		console.log(`${APP_TAG} Completed deletion for conversations ${conversationIds.join()}`);
	}
	var author = {
		"name": "ganyuke",
		"url": "https://github.com/ganyuke"
	};
	var version = "2.0.1";
	var repository = {
		"type": "git",
		"url": "https://github.com/ganyuke/copilot-exporter.git"
	};
	var STATUS_COLORS = {
		exporting: "#ca8a04",
		exported: "#16a34a",
		deleting: "#ca8a04",
		deleted: "#6b7280",
		error: "#dc2626"
	};
	var STATUS_LABELS = {
		exporting: "exporting…",
		exported: "exported",
		deleting: "deleting…",
		deleted: "deleted",
		error: "error"
	};
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
    width: 90vw; max-width: 800px;
    box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    font-family: sans-serif;
  `;
		modal.innerHTML = `
    <h2 style="margin:0;">Export conversations</h2>
    <p style="margin: 0.5rem 0;color: darkorchid;"><a style="color: inherit;" href="${repository.url}" target="_blank">M365 Copilot Exporter</a> v${version} by <a style="color: inherit;" href="${author.url}" target="_blank">${author.name}</a></p>

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
		function formatPrettyDate(ms) {
			return new Intl.DateTimeFormat(void 0, {
				dateStyle: "medium",
				timeStyle: "short"
			}).format(new Date(ms));
		}
		function findStatusCell(conversationId) {
			return document.querySelector(`#chatTableBody input[type="checkbox"][data-id="${CSS.escape(conversationId)}"]`)?.closest("tr")?.querySelector(".status-cell") ?? null;
		}
		function setRowStatus(conversationId, status, error) {
			const cell = findStatusCell(conversationId);
			if (!cell) return;
			cell.textContent = STATUS_LABELS[status];
			cell.style.color = STATUS_COLORS[status];
			if (status === "error" && error) cell.title = error;
			else cell.removeAttribute("title");
		}
		function clearRowStatus(conversationIds) {
			for (const id of conversationIds) {
				const cell = findStatusCell(id);
				if (!cell) continue;
				cell.textContent = "";
				cell.style.color = "";
				cell.removeAttribute("title");
			}
		}
		function updateSelectedCount() {
			const checkboxes = document.querySelectorAll("#chatTableBody input[type=\"checkbox\"]");
			const selected = document.querySelectorAll("#chatTableBody input[type=\"checkbox\"]:checked").length;
			const loaded = checkboxes.length;
			const countEl = document.getElementById("selectedCount");
			countEl.textContent = `(${selected}/${loaded})`;
			const selectAll = document.getElementById("selectAllCheckbox");
			selectAll.checked = selected > 0 && selected === loaded;
		}
		function captureTableState() {
			const state = new Map();
			const rows = document.querySelectorAll("#chatTableBody tr[data-conversation-id]");
			for (const row of rows) {
				const id = row.getAttribute("data-conversation-id");
				if (!id) continue;
				const checkbox = row.querySelector("input[type=\"checkbox\"]");
				const statusCell = row.querySelector(".status-cell");
				state.set(id, {
					checked: checkbox?.checked ?? false,
					chatName: row.getAttribute("data-chat-name") ?? "",
					createTimeUtc: Number(row.getAttribute("data-create-time")),
					updateTimeUtc: Number(row.getAttribute("data-update-time")),
					statusText: statusCell?.textContent ?? "",
					statusColor: statusCell?.style.color ?? "",
					statusTitle: statusCell?.getAttribute("title") ?? null
				});
			}
			return state;
		}
		function chatDataMatches(snapshot, data) {
			return snapshot.chatName === data.chatName && snapshot.createTimeUtc === data.createTimeUtc && snapshot.updateTimeUtc === data.updateTimeUtc;
		}
		function renderChatTable(chats, previousState) {
			const tbody = document.getElementById("chatTableBody");
			tbody.innerHTML = "";
			const sorted = [...chats].sort((a, b) => b.updateTimeUtc - a.updateTimeUtc);
			if (sorted.length === 0) {
				const row = document.createElement("tr");
				const cell = document.createElement("td");
				cell.colSpan = 5;
				cell.textContent = "No conversations found.";
				cell.style.padding = "8px";
				row.appendChild(cell);
				tbody.appendChild(row);
			} else for (const data of sorted) {
				const row = document.createElement("tr");
				row.setAttribute("data-conversation-id", data.conversationId);
				row.setAttribute("data-chat-name", data.chatName);
				row.setAttribute("data-create-time", String(data.createTimeUtc));
				row.setAttribute("data-update-time", String(data.updateTimeUtc));
				row.style.borderBottom = "1px solid #e5e7eb";
				const checkboxTd = document.createElement("td");
				const checkbox = document.createElement("input");
				checkbox.type = "checkbox";
				checkbox.dataset.id = data.conversationId;
				checkbox.dataset.title = data.chatName;
				const previous = previousState?.get(data.conversationId);
				if (previous) checkbox.checked = previous.checked;
				checkboxTd.appendChild(checkbox);
				const nameTd = document.createElement("td");
				nameTd.className = "name-cell";
				nameTd.style.cssText = "overflow: hidden; text-overflow: ellipsis; white-space: nowrap; padding: 4px 8px;";
				nameTd.title = data.chatName;
				nameTd.textContent = data.chatName;
				const createdTd = document.createElement("td");
				createdTd.style.cssText = "font-size: 0.875em; padding: 4px 8px;";
				createdTd.textContent = formatPrettyDate(data.createTimeUtc);
				const updatedTd = document.createElement("td");
				updatedTd.style.cssText = "font-size: 0.875em; padding: 4px 8px;";
				updatedTd.textContent = formatPrettyDate(data.updateTimeUtc);
				const statusTd = document.createElement("td");
				statusTd.className = "status-cell";
				statusTd.style.padding = "4px 8px";
				if (previous && chatDataMatches(previous, data)) {
					statusTd.textContent = previous.statusText;
					statusTd.style.color = previous.statusColor;
					if (previous.statusTitle) statusTd.title = previous.statusTitle;
				}
				row.append(checkboxTd, nameTd, createdTd, updatedTd, statusTd);
				tbody.appendChild(row);
			}
			updateSelectedCount();
		}
		function removeProgressBar() {
			document.querySelector("#chat-export-progress-bar-container")?.remove();
		}
		function createProgressBar(items, initialString) {
			if (items.length < 1) return;
			removeProgressBar();
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
			progressBar.max = items.length;
			progressBar.value = 0;
			label.htmlFor = "chat-export-progress-bar";
			titleSpan.textContent = initialString;
			progressTextSpan.textContent = `0/${items.length}`;
			label.append(titleSpan, progressTextSpan);
			progressBarContainer.append(label, progressBar);
			modal.append(progressBarContainer);
			const progressUpdater = (progress) => {
				titleSpan.textContent = items[progress].title;
				progressTextSpan.textContent = `${progress + 1}/${items.length}`;
				progressBar.value = progress + 1;
				if (progressBar.value === progressBar.max) setTimeout(() => {
					progressBarContainer.remove();
				}, 3e3);
			};
			return progressUpdater;
		}
		function createExportProgressHandler(items, initialString) {
			if (items.length < 1) return;
			removeProgressBar();
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
			progressBar.max = items.length;
			progressBar.value = 0;
			label.htmlFor = "chat-export-progress-bar";
			titleSpan.textContent = initialString;
			progressTextSpan.textContent = `0/${items.length}`;
			label.append(titleSpan, progressTextSpan);
			progressBarContainer.append(label, progressBar);
			modal.append(progressBarContainer);
			let completed = 0;
			const handler = (event) => {
				const item = items[event.index];
				if (event.phase === "start") {
					titleSpan.textContent = item.title;
					setRowStatus(item.id, "exporting");
				} else if (event.phase === "success") {
					setRowStatus(item.id, "exported");
					completed++;
					progressBar.value = completed;
					progressTextSpan.textContent = `${completed}/${items.length}`;
				} else {
					setRowStatus(item.id, "error", event.error);
					completed++;
					progressBar.value = completed;
					progressTextSpan.textContent = `${completed}/${items.length}`;
				}
				if (completed === items.length) setTimeout(() => {
					progressBarContainer.remove();
				}, 3e3);
			};
			return handler;
		}
		async function fetchChats() {
			const previousState = captureTableState();
			const tbody = document.getElementById("chatTableBody");
			tbody.innerHTML = "<tr><td colspan=\"5\" style=\"color: #666; padding: 8px;\">Loading…</td></tr>";
			try {
				const n = document.getElementById("conversation-fetch-list-max").valueAsNumber;
				const maxChats = isNaN(n) ? 15 : n;
				console.log(`${APP_TAG} Getting MSAL ids...`);
				const msalIds = getMsalIds();
				console.log(`${APP_TAG} Getting access token...`);
				renderChatTable((await fetchCopilotChats(await getAccessToken(msalIds), msalIds.localAccountId, msalIds.tenantId, maxChats)).chats, previousState);
				updateSelectedCount();
			} catch {
				tbody.innerHTML = "<tr><td colspan=\"5\" style=\"color: #dc2626; padding: 8px;\">Failed to load conversations.</td></tr>";
				const selectAll = document.getElementById("selectAllCheckbox");
				selectAll.checked = false;
				document.getElementById("selectedCount").textContent = "(0/0)";
			}
		}
		function getSelectedChats() {
			const checkboxes = document.querySelectorAll("#chatTableBody input[type=\"checkbox\"]:checked");
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
		function getExportFormat() {
			return document.getElementById("export-format-select").value;
		}
		function getOutputMode() {
			return document.getElementById("export-output-mode-select").value;
		}
		function parseCopilotJsonFile(text) {
			const parsed = JSON.parse(text);
			if (Array.isArray(parsed)) return parsed;
			return [parsed];
		}
		function showExportResultAlert(successCount, totalCount) {
			if (successCount === totalCount) alert(`Successfully exported ${successCount} of ${totalCount} conversations.`);
			else alert(`Exported ${successCount} of ${totalCount} conversations. Hover over red statuses for error details.`);
		}
		function sanitizeFilename(name) {
			return name.replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").trim() || "conversation";
		}
		async function exportChats() {
			const items = getSelectedChats();
			if (items.length === 0) return;
			clearRowStatus(items.map((i) => i.id));
			const handler = createExportProgressHandler(items, "Exporting...");
			if (!handler) return;
			const result = await exportBulkDirect(items.map((i) => i.id), handler, getExportFormat(), getOutputMode());
			showExportResultAlert(result.successCount, result.totalCount);
		}
		async function deleteChats() {
			const items = getSelectedChats();
			if (items.length === 0) return;
			const message = items.length === 1 ? `Permanently delete "${items[0].title}"? This cannot be undone.` : `Permanently delete ${items.length} conversations? This cannot be undone.`;
			if (!confirm(message)) return;
			clearRowStatus(items.map((i) => i.id));
			items.forEach((i) => setRowStatus(i.id, "deleting"));
			const progressUpdater = createProgressBar(items, "Deleting...");
			try {
				await deleteBulk(items.map((i) => i.id), progressUpdater ?? (() => {}));
				items.forEach((i) => setRowStatus(i.id, "deleted"));
			} catch (err) {
				const msg = err instanceof Error ? err.message : String(err);
				items.forEach((i) => setRowStatus(i.id, "error", msg));
			}
		}
		const selectAllCheckbox = document.getElementById("selectAllCheckbox");
		selectAllCheckbox.addEventListener("change", () => {
			document.querySelectorAll("#chatTableBody input[type=\"checkbox\"]").forEach((cb) => {
				cb.checked = selectAllCheckbox.checked;
			});
			updateSelectedCount();
		});
		document.getElementById("chatTableBody").addEventListener("change", (e) => {
			if (e.target.matches("input[type=\"checkbox\"]")) updateSelectedCount();
		});
		document.getElementById("export-conversations-button").addEventListener("click", exportChats);
		const exportFormatSelect = document.getElementById("export-format-select");
		const exportOutputModeSelect = document.getElementById("export-output-mode-select");
		const combinedOutputOption = exportOutputModeSelect.querySelector("option[value=\"combined\"]");
		exportFormatSelect.addEventListener("change", () => {
			if (exportFormatSelect.value === "markdown") {
				combinedOutputOption.disabled = true;
				if (exportOutputModeSelect.value === "combined") exportOutputModeSelect.value = "zip";
			} else combinedOutputOption.disabled = false;
		});
		document.getElementById("delete-conversations-button").addEventListener("click", deleteChats);
		document.getElementById("conversation-refetch").addEventListener("click", fetchChats);
		const fileInput = document.getElementById("copilot-json-upload");
		const convertBtn = document.getElementById("convert-uploaded-button");
		const convertFormatSelect = document.getElementById("convert-format-select");
		convertBtn.addEventListener("click", () => fileInput.click());
		fileInput.addEventListener("change", async () => {
			const files = fileInput.files;
			if (!files || files.length === 0) return;
			const format = convertFormatSelect.value;
			const conversations = [];
			for (const file of files) conversations.push(...parseCopilotJsonFile(await file.text()));
			if (format === "chatgpt") {
				const converted = conversations.map(mapToConversationJson);
				downloadBlobAsFile(new Blob([JSON.stringify(converted, null, 2)], { type: "application/json" }), "conversations.json");
			} else for (const conversation of conversations) downloadBlobAsFile(new Blob([mapToMarkdown(conversation)], { type: "text/markdown" }), `${sanitizeFilename(conversation.chatName)}.md`);
			fileInput.value = "";
		});
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
