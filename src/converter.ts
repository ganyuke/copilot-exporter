/**
 * For M365 Copilot conversation to ChatGPT JSON format conversion.
 */
import type { CopilotConversation, CopilotMessage } from "./api";

type ExportRole = "system" | "user" | "assistant";

interface ExportMessage {
    id: string;
    author: {
        role: ExportRole;
        name: string | null;
        metadata: Record<string, never>;
    };
    create_time: number | null;
    update_time: number | null;
    content: {
        content_type: "text";
        parts: string[];
    };
    status: "finished_successfully";
    end_turn: boolean | null;
    weight: number;
    metadata: Record<string, unknown>;
    recipient: "all";
    channel: string | null;
}

interface ExportNode {
    id: string;
    message: ExportMessage | null;
    parent: string | null;
    children: string[];
}

interface ConversationJson {
    title: string;
    create_time: number | null;
    update_time: number | null;
    mapping: Record<string, ExportNode>;
    moderation_results: unknown[];
    current_node: string;
    plugin_ids: null;
    conversation_id: string;
    conversation_template_id: null;
    gizmo_id: null;
    gizmo_type: null;
    is_archived: boolean;
    is_starred: boolean | null;
    safe_urls: string[];
    default_model_slug: string | null;
    disabled_tool_ids: string[];
    id: string;
}

function millisToSeconds(ms: number | undefined | null): number | null {
    return typeof ms === "number" ? ms / 1000 : null;
}

function isoToSeconds(iso: string | undefined | null): number | null {
    if (!iso) return null;

    const ms = Date.parse(iso);
    return Number.isFinite(ms) ? ms / 1000 : null;
}

function makeUuid(): string {
    return crypto.randomUUID();
}

function roleFromSource(author: string): "user" | "assistant" {
    return author === "user" ? "user" : "assistant";
}

function makeRootNode(id: string): ExportNode {
    return {
        id,
        message: null,
        parent: null,
        children: [],
    };
}

function makeExportMessage(source: CopilotMessage): ExportMessage {
    const role = roleFromSource(source.author);

    return {
        id: source.messageId,
        author: {
            role,
            name: null,
            metadata: {},
        },
        create_time: isoToSeconds(source.createdAt ?? source.timestamp),
        update_time: null,
        content: {
            content_type: "text",
            parts: [source.text ?? ""],
        },
        status: "finished_successfully",
        end_turn: role === "assistant" ? true : null,
        weight: 1,
        metadata: {},
        recipient: "all",
        channel: null,
    };
}

function makeExportNode(source: CopilotMessage, parentId: string): ExportNode {
    return {
        id: source.messageId,
        message: makeExportMessage(source),
        parent: parentId,
        children: [],
    };
}

export function mapToConversationJson(source: CopilotConversation): ConversationJson {
    const rootId = makeUuid();

    const mapping: Record<string, ExportNode> = {
        [rootId]: makeRootNode(rootId),
    };

    let parentId = rootId;
    let currentNode = rootId;

    for (const message of source.messages) {
        const nodeId = message.messageId;

        const node = makeExportNode(message, parentId);

        mapping[nodeId] = node;
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
        id: source.conversationId,
    };
}