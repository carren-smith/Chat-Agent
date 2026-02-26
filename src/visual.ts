"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { FormattingSettings } from "./settings";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import DataViewObjects = powerbi.DataViewObjects;

// ─────────────────────────────────────────────────────────────────────────────
// Interfaces
// ─────────────────────────────────────────────────────────────────────────────

interface ChatMessage {
    type: "user" | "assistant" | "loading" | "error";
    content: string;
    timestamp: Date;
    chartData?: any;
}

interface ChatSession {
    id: string;
    title: string;
    messages: ChatMessage[];
    lastUpdated: Date;
}

interface DataQuery {
    intent: "data_query";
    filters?: Array<{
        column: string;
        operator: ">" | "<" | ">=" | "<=" | "==" | "!=" | "contains";
        value: string | number;
    }>;
    groupBy?: string[];
    aggregations?: Array<{ column: string; op: "sum" | "avg" | "count" | "max" | "min" | "first" }>;
    sort?: { column: string; direction: "asc" | "desc" };
    limit?: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// TmdlManager
// ─────────────────────────────────────────────────────────────────────────────

class TmdlManager {
    private static readonly MAX_CHUNK_SIZE = 25000;
    private static readonly MAX_CHUNKS = 50;

    public static cleanTmdl(code: string): string {
        return (code || "")
            .split("\n")
            .filter(l => !l.trim().startsWith("//") && !l.includes("annotation"))
            .join("\n")
            .replace(/\bmodifiedTime\s*=.*$/gm, "")
            .trim();
    }

    public static saveTmdl(host: IVisualHost, code: string): void {
        const clean = this.cleanTmdl(code);
        const chunks: string[] = [];
        for (let i = 0; i < clean.length && chunks.length < this.MAX_CHUNKS; i += this.MAX_CHUNK_SIZE) {
            chunks.push(clean.slice(i, i + this.MAX_CHUNK_SIZE));
        }
        const props: any = { chunkCount: String(chunks.length), tmdlCode: clean };
        for (let i = 0; i < this.MAX_CHUNKS; i++) props[`chunk${i}`] = chunks[i] || "";
        host.persistProperties({ merge: [{ objectName: "tmdlSettings", properties: props, selector: null }] });
    }

    public static loadTmdl(objects?: DataViewObjects): string {
        const s = objects && (objects as any).tmdlSettings;
        if (!s) return "";
        const count = Number(s.chunkCount || 0);
        if (count > 0) {
            return Array.from({ length: count }, (_, i) => String(s[`chunk${i}`] || "")).join("");
        }
        return String(s.tmdlCode || "");
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// ChatHistoryManager
// ─────────────────────────────────────────────────────────────────────────────

class ChatHistoryManager {
    private static readonly MAX_CHUNK_SIZE = 25000;
    private static readonly MAX_CHUNKS = 50;

    public static saveHistory(host: IVisualHost, sessions: ChatSession[]): void {
        const payload = JSON.stringify(sessions);
        const chunks: string[] = [];
        for (let i = 0; i < payload.length && chunks.length < this.MAX_CHUNKS; i += this.MAX_CHUNK_SIZE) {
            chunks.push(payload.slice(i, i + this.MAX_CHUNK_SIZE));
        }
        const props: any = { chunkCount: String(chunks.length) };
        for (let i = 0; i < this.MAX_CHUNKS; i++) props[`chunk${i}`] = chunks[i] || "";
        host.persistProperties({ merge: [{ objectName: "historySettings", properties: props, selector: null }] });
    }

    public static loadHistory(objects?: DataViewObjects): ChatSession[] {
        const s = objects && (objects as any).historySettings;
        if (!s) return [];
        const count = Number(s.chunkCount || 0);
        if (count <= 0) return [];
        let json = "";
        for (let i = 0; i < count; i++) json += String(s[`chunk${i}`] || "");
        try {
            const parsed = JSON.parse(json) as ChatSession[];
            return parsed.map(sess => ({
                ...sess,
                lastUpdated: new Date(sess.lastUpdated),
                messages: (sess.messages || []).map(m => ({ ...m, timestamp: new Date(m.timestamp) }))
            }));
        } catch { return []; }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Visual
// ─────────────────────────────────────────────────────────────────────────────

export class Visual implements IVisual {
    // Core
    private target: HTMLElement;
    private formattingSettings: FormattingSettings;
    private formattingSettingsService: FormattingSettingsService;
    private dataView: DataView;
    private host: IVisualHost;

    // State
    private chatMessages: ChatMessage[] = [];
    private currentSessionId: string;
    private charts: Map<string, any> = new Map();

    // DOM
    private messagesContainer: HTMLElement;
    private inputElement: HTMLTextAreaElement;
    private sendButton: HTMLButtonElement;
    private newChatButton: HTMLButtonElement;
    private historyButton: HTMLButtonElement;
    private settingsButton: HTMLButtonElement;

    // Streaming
    private currentStreamingMessageIndex = -1;
    private streamingMessageElement: HTMLElement | null = null;
    private abortController: AbortController | null = null;
    private isGenerating = false;
    private isComposing = false;

    // ── Constructor ───────────────────────────────────────────────────────────

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new FormattingSettings();
        this.currentSessionId = this.generateSessionId();
        this.initializeDOM();
        this.setupEventListeners();
    }

    // ── DOM Init ──────────────────────────────────────────────────────────────

    private initializeDOM(): void {
        this.target.innerHTML = "";

        const style = document.createElement("style");
        style.textContent = `
/* Layout */
.abi-root{height:100%;display:flex;flex-direction:column;font-family:'Segoe UI',Tahoma,sans-serif;background:#f5f6f7;overflow:hidden}
.abi-header{display:flex;justify-content:space-between;align-items:center;padding:10px 14px;background:#fff;border-bottom:1px solid #e4e4e4;flex-shrink:0;box-shadow:0 1px 3px rgba(0,0,0,.04)}
.abi-title{font-weight:700;font-size:15px;color:#1a1a2e;letter-spacing:.3px}
.abi-actions{display:flex;gap:6px}
.abi-actions button{background:none;border:1px solid #dce0e5;border-radius:6px;padding:5px 9px;cursor:pointer;font-size:13px;color:#555;line-height:1;transition:all .12s}
.abi-actions button:hover{background:#f0f2f5;border-color:#bbb}

/* Messages area */
.abi-messages{flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px}

/* Message bubbles */
.msg{display:flex;flex-direction:column;max-width:86%}
.msg.user{align-self:flex-end;align-items:flex-end}
.msg.assistant,.msg.error,.msg.loading{align-self:flex-start;align-items:flex-start}
.msg-bubble{padding:10px 14px;border-radius:14px;font-size:13px;line-height:1.65;word-break:break-word;max-width:100%}
.msg.user .msg-bubble{background:#0078d4;color:#fff;border-bottom-right-radius:4px}
.msg.assistant .msg-bubble{background:#fff;color:#222;border:1px solid #e5e8ec;border-bottom-left-radius:4px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.msg.error .msg-bubble{background:#fff4e5;color:#a94400;border:1px solid #ffb347;border-radius:8px;font-size:12px}
.msg.loading .msg-bubble{background:#f0f0f0;color:#888;font-style:italic;border-radius:8px}

/* Assistant content */
.msg-content{min-width:0}
.msg-content h3{font-size:14px;margin:10px 0 5px;color:#0f2044}
.msg-content h4{font-size:13px;margin:8px 0 4px;color:#333}
.msg-content p{margin:4px 0}
.msg-content ul,.msg-content ol{padding-left:18px;margin:4px 0}
.msg-content li{margin:2px 0}
.msg-content b,.msg-content strong{color:#0f2044}
.msg-content .report-section{font-size:12px;font-weight:600;color:#0078d4;margin-bottom:6px}
.msg-content .report-emphasis{font-weight:600;color:#0078d4}

/* Tables */
.msg-content table.report-table{border-collapse:collapse;width:100%;margin:8px 0;font-size:12px;border-radius:6px;overflow:hidden}
.msg-content .report-table th{background:#e8f0fb;padding:7px 10px;border:1px solid #c8d8f0;font-weight:600;text-align:left;white-space:nowrap;color:#1a3a6e}
.msg-content .report-table td{padding:6px 10px;border:1px solid #dde4ee;vertical-align:top}
.msg-content .report-table tbody tr:nth-child(even){background:#f8fafd}
.msg-content .report-table tbody tr:hover{background:#eef4ff}

/* Copy button */
.copy-btn{margin-top:7px;padding:3px 10px;background:none;border:1px solid #dce0e5;border-radius:4px;cursor:pointer;font-size:11px;color:#666;display:inline-block;transition:all .12s}
.copy-btn:hover{background:#f0f2f5;color:#333}

/* Suggestion chips */
.abi-suggestions{display:flex;gap:8px;flex-wrap:wrap;padding:2px 0;align-self:flex-start;max-width:100%}
.suggestion-btn{padding:6px 14px;background:#fff;border:1px solid #0078d4;border-radius:16px;color:#0078d4;font-size:12px;cursor:pointer;white-space:nowrap;transition:all .15s;flex-shrink:0}
.suggestion-btn:hover{background:#0078d4;color:#fff}

/* Input area */
.abi-input-wrap{display:flex;gap:8px;padding:10px 14px;background:#fff;border-top:1px solid #e4e4e4;align-items:flex-end;flex-shrink:0}
.abi-input{flex:1;resize:none;border:1px solid #dce0e5;border-radius:10px;padding:9px 13px;font-size:13px;font-family:inherit;outline:none;line-height:1.45;overflow-y:hidden;min-height:38px;max-height:150px;transition:border-color .15s}
.abi-input:focus{border-color:#0078d4;box-shadow:0 0 0 2px rgba(0,120,212,.12)}
.abi-send{padding:9px 18px;background:#0078d4;color:#fff;border:none;border-radius:10px;cursor:pointer;font-size:13px;white-space:nowrap;font-family:inherit;font-weight:500;transition:background .15s;flex-shrink:0}
.abi-send:hover{background:#006cbf}
.abi-send:disabled{background:#c5c9d0;cursor:not-allowed}
.abi-send.generating{background:#d32f2f}
.abi-send.generating:hover{background:#b71c1c}

/* Modal */
.abi-modal{position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:9999}
.modal-box{background:#fff;border-radius:10px;width:620px;max-width:94vw;max-height:88vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,.18)}
.modal-hdr{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;border-bottom:1px solid #eee;flex-shrink:0}
.modal-hdr h3{margin:0;font-size:15px;color:#1a1a2e}
.modal-x{background:none;border:none;font-size:22px;cursor:pointer;color:#aaa;line-height:1;padding:0 2px}
.modal-x:hover{color:#333}
.modal-body{padding:18px;overflow-y:auto;flex:1}
.modal-field{margin-bottom:16px}
.modal-field label{display:block;font-size:12px;color:#555;margin-bottom:5px;font-weight:600}
.modal-field input,.modal-field textarea{width:100%;box-sizing:border-box;border:1px solid #dce0e5;border-radius:7px;padding:8px 11px;font-size:13px;font-family:inherit;outline:none;color:#222}
.modal-field input:focus,.modal-field textarea:focus{border-color:#0078d4;box-shadow:0 0 0 2px rgba(0,120,212,.1)}
.modal-field textarea{resize:vertical;min-height:120px}
.modal-ftr{padding:12px 18px;border-top:1px solid #eee;display:flex;justify-content:flex-end;gap:8px;flex-shrink:0}
.btn-primary{padding:8px 20px;background:#0078d4;color:#fff;border:none;border-radius:7px;cursor:pointer;font-size:13px;font-weight:500}
.btn-primary:hover{background:#006cbf}
.btn-secondary{padding:8px 20px;background:#f5f6f7;color:#333;border:1px solid #dce0e5;border-radius:7px;cursor:pointer;font-size:13px}
.btn-secondary:hover{background:#ebedf0}

/* History items */
.hist-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border:1px solid #e4e4e4;border-radius:8px;margin-bottom:8px;cursor:pointer;transition:background .12s}
.hist-item:hover{background:#f5f7fb}
.hist-title{flex:1;font-size:13px;color:#222;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.hist-date{font-size:11px;color:#999;white-space:nowrap}
.hist-del{background:none;border:none;color:#bbb;cursor:pointer;font-size:16px;padding:0 4px;line-height:1}
.hist-del:hover{color:#d32f2f}
`;
        this.target.appendChild(style);

        // Root
        const root = document.createElement("div");
        root.className = "abi-root";

        // Header
        const header = document.createElement("div");
        header.className = "abi-header";
        header.innerHTML = `<div class="abi-title">ABI Chat</div><div class="abi-actions"></div>`;
        const actions = header.querySelector(".abi-actions") as HTMLElement;

        this.newChatButton = this.makeBtn("+", "新建会话");
        this.historyButton = this.makeBtn("🕒", "历史记录");
        this.settingsButton = this.makeBtn("⚙", "设置");
        this.settingsButton.style.display = "none";
        actions.append(this.newChatButton, this.historyButton, this.settingsButton);

        // Messages
        this.messagesContainer = document.createElement("div");
        this.messagesContainer.className = "abi-messages";

        // Input
        const inputWrap = document.createElement("div");
        inputWrap.className = "abi-input-wrap";

        this.inputElement = document.createElement("textarea");
        this.inputElement.className = "abi-input";
        this.inputElement.rows = 1;
        this.inputElement.placeholder = "请输入您的问题...";

        this.sendButton = document.createElement("button");
        this.sendButton.className = "abi-send";
        this.sendButton.textContent = "发送";
        this.sendButton.disabled = true;

        inputWrap.append(this.inputElement, this.sendButton);
        root.append(header, this.messagesContainer, inputWrap);
        this.target.appendChild(root);
    }

    private makeBtn(text: string, title: string): HTMLButtonElement {
        const b = document.createElement("button");
        b.textContent = text;
        b.title = title;
        return b;
    }

    private setupEventListeners(): void {
        this.sendButton.addEventListener("click", () => {
            if (this.isGenerating) this.stopGeneration();
            else this.sendMessage();
        });
        this.inputElement.addEventListener("compositionstart", () => (this.isComposing = true));
        this.inputElement.addEventListener("compositionend", () => (this.isComposing = false));
        this.inputElement.addEventListener("keydown", e => {
            if (e.key === "Enter" && !e.shiftKey && !this.isComposing) {
                e.preventDefault();
                if (!this.isGenerating) this.sendMessage();
            }
        });
        this.inputElement.addEventListener("input", () => {
            this.adjustTextareaHeight();
            this.updateSendButtonState();
        });
        this.newChatButton.addEventListener("click", () => this.startNewChat());
        this.historyButton.addEventListener("click", () => this.showHistoryModal());
        this.settingsButton.addEventListener("click", () => this.showSettingsModal());
    }

    // ── PowerBI Lifecycle ─────────────────────────────────────────────────────

    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
            FormattingSettings, options.dataViews?.[0]
        );
        this.formattingSettings.aboutCardSettings.revertToDefault();
        this.dataView = options.dataViews?.[0];

        const viewMode = (options as any).viewMode;
        this.settingsButton.style.display = viewMode === 1 ? "inline-block" : "none";

        // Update welcome message text if it's the only message shown
        if (this.chatMessages.length === 1 && this.chatMessages[0].type === "assistant") {
            this.chatMessages[0].content = this.getWelcomeMessage();
            this.renderChatMessages();
        }

        // Initialise chat on first load
        if (this.chatMessages.length === 0) {
            const objects = this.dataView?.metadata?.objects;
            const sessions = ChatHistoryManager.loadHistory(objects);
            const lastId = (objects as any)?.historySettings?.lastActiveSessionId;
            const lastSess = sessions.find(s => s.id === lastId);
            if (lastSess) {
                this.loadSession(lastSess);
            } else if (sessions.length > 0) {
                sessions.sort((a, b) => b.lastUpdated.getTime() - a.lastUpdated.getTime());
                this.loadSession(sessions[0]);
            } else {
                this.addWelcomeMessage();
            }
        }
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ── Message Flow ──────────────────────────────────────────────────────────

    private async sendMessage(): Promise<void> {
        const text = this.inputElement.value.trim();
        if (!text || this.isGenerating) return;

        this.addMessage("user", text);
        this.inputElement.value = "";
        this.adjustTextareaHeight();
        this.updateSendButtonState();

        this.isGenerating = true;
        this.abortController = new AbortController();
        this.updateSendButtonState();

        try {
            this.startStreamingMessage();
            const full = await this.callAIAPIWithStreaming(text, chunk => this.updateStreamingMessage(chunk));
            this.finalizeStreamingMessage(full);
            this.saveCurrentChatToHistory();
        } catch (err: any) {
            if (err?.name === "AbortError") {
                this.finalizeStreamingMessage(
                    this.currentStreamingMessageIndex >= 0
                        ? this.chatMessages[this.currentStreamingMessageIndex].content
                        : ""
                );
                this.saveCurrentChatToHistory();
            } else {
                this.removeStreamingMessage();
                this.showErrorMessage(String(err?.message || err || "请求失败，请检查 API 配置"));
            }
        } finally {
            this.isGenerating = false;
            this.abortController = null;
            this.updateSendButtonState();
        }
    }

    private stopGeneration(): void {
        this.abortController?.abort();
        this.isGenerating = false;
        this.abortController = null;
        this.updateSendButtonState();
    }

    private addMessage(type: ChatMessage["type"], content: string, chartData?: any): void {
        this.chatMessages.push({ type, content, timestamp: new Date(), chartData });
        this.renderChatMessages();
    }

    private startStreamingMessage(): void {
        this.chatMessages.push({ type: "assistant", content: "", timestamp: new Date() });
        this.currentStreamingMessageIndex = this.chatMessages.length - 1;
        this.renderChatMessages();
        // Get direct DOM reference to the new message's content element for efficient updates
        const contentEls = this.messagesContainer.querySelectorAll<HTMLElement>(".msg.assistant .msg-content");
        this.streamingMessageElement = contentEls[contentEls.length - 1] || null;
    }

    private updateStreamingMessage(chunk: string): void {
        if (this.currentStreamingMessageIndex < 0) return;
        this.chatMessages[this.currentStreamingMessageIndex].content += chunk;
        if (this.streamingMessageElement) {
            const raw = this.chatMessages[this.currentStreamingMessageIndex].content;
            this.streamingMessageElement.innerHTML = this.formatToReportStyle(raw);
            this.messagesContainer.scrollTop = this.messagesContainer.scrollHeight;
        }
    }

    private finalizeStreamingMessage(fullText: string): void {
        if (this.currentStreamingMessageIndex < 0) return;
        const parsed = this.parseAIResponse(fullText);
        this.chatMessages[this.currentStreamingMessageIndex].content = parsed.text;
        this.chatMessages[this.currentStreamingMessageIndex].chartData = parsed.chartData;
        this.currentStreamingMessageIndex = -1;
        this.streamingMessageElement = null;
        this.renderChatMessages();
    }

    private removeStreamingMessage(): void {
        if (this.currentStreamingMessageIndex >= 0) {
            this.chatMessages.splice(this.currentStreamingMessageIndex, 1);
        }
        this.currentStreamingMessageIndex = -1;
        this.streamingMessageElement = null;
        this.renderChatMessages();
    }

    private showErrorMessage(text: string): void {
        this.addMessage("error", text);
    }

    // ── Rendering ─────────────────────────────────────────────────────────────

    private renderChatMessages(): void {
        this.messagesContainer.innerHTML = "";
        this.chatMessages.forEach(msg => {
            const row = document.createElement("div");
            row.className = `msg ${msg.type}`;

            const bubble = document.createElement("div");
            bubble.className = "msg-bubble";

            if (msg.type === "user") {
                bubble.textContent = msg.content;
            } else if (msg.type === "assistant") {
                const content = document.createElement("div");
                content.className = "msg-content";
                content.innerHTML = this.formatToReportStyle(msg.content);
                bubble.appendChild(content);

                const copyBtn = document.createElement("button");
                copyBtn.className = "copy-btn";
                copyBtn.textContent = "复制";
                copyBtn.onclick = () => this.copyToClipboard(msg.content);
                bubble.appendChild(copyBtn);
            } else {
                bubble.textContent = msg.content;
            }

            row.appendChild(bubble);
            this.messagesContainer.appendChild(row);
        });
        this.messagesContainer.scrollTop = this.messagesContainer.scrollHeight;
    }

    private addWelcomeMessage(): void {
        this.chatMessages = [{ type: "assistant", content: this.getWelcomeMessage(), timestamp: new Date() }];
        this.renderChatMessages();
        this.addSuggestionChips();
    }

    private getWelcomeMessage(): string {
        const obj = this.dataView?.metadata?.objects as any;
        return obj?.welcomeSettings?.welcomeMessage
            || this.formattingSettings.welcomeSettingsCard.welcomeMessage.value
            || "我是 ABI Chat，帮助您从数据中挖掘价值，请告诉我您想了解什么内容？";
    }

    private addSuggestionChips(): void {
        setTimeout(() => {
            const card = this.formattingSettings.suggestionSettingsCard;
            const obj = this.dataView?.metadata?.objects as any;
            const qs = [
                obj?.suggestionSettings?.question1 || card.question1.value,
                obj?.suggestionSettings?.question2 || card.question2.value,
                obj?.suggestionSettings?.question3 || card.question3.value,
            ].filter(Boolean);

            if (qs.length === 0) return;

            const wrap = document.createElement("div");
            wrap.className = "abi-suggestions";
            qs.forEach(q => {
                const btn = document.createElement("button");
                btn.className = "suggestion-btn";
                btn.textContent = q;
                btn.onclick = () => {
                    this.inputElement.value = q;
                    this.adjustTextareaHeight();
                    this.sendMessage();
                };
                wrap.appendChild(btn);
            });
            this.messagesContainer.appendChild(wrap);
            this.messagesContainer.scrollTop = this.messagesContainer.scrollHeight;
        }, 50);
    }

    private adjustTextareaHeight(): void {
        this.inputElement.style.height = "auto";
        this.inputElement.style.height = `${Math.min(this.inputElement.scrollHeight, 150)}px`;
    }

    private updateSendButtonState(): void {
        const hasText = this.inputElement.value.trim().length > 0;
        if (this.isGenerating) {
            this.sendButton.disabled = false;
            this.sendButton.classList.add("generating");
            this.sendButton.textContent = "停止";
        } else {
            this.sendButton.disabled = !hasText;
            this.sendButton.classList.remove("generating");
            this.sendButton.textContent = "发送";
        }
    }

    // ── Modals ────────────────────────────────────────────────────────────────

    private showHistoryModal(): void {
        const sessions = ChatHistoryManager.loadHistory(this.dataView?.metadata?.objects);
        const modal = document.createElement("div");
        modal.className = "abi-modal";

        const box = document.createElement("div");
        box.className = "modal-box";
        box.innerHTML = `
            <div class="modal-hdr">
                <h3>历史记录</h3>
                <button class="modal-x" title="关闭">×</button>
            </div>
            <div class="modal-body" id="histBody"></div>
        `;

        const body = box.querySelector("#histBody") as HTMLElement;

        if (sessions.length === 0) {
            body.innerHTML = `<div style="color:#999;font-size:13px;text-align:center;padding:20px">暂无历史记录</div>`;
        } else {
            sessions
                .sort((a, b) => b.lastUpdated.getTime() - a.lastUpdated.getTime())
                .forEach(sess => {
                    const item = document.createElement("div");
                    item.className = "hist-item";
                    item.innerHTML = `
                        <div class="hist-title">${this.sanitizeHTML(sess.title)}</div>
                        <div class="hist-date">${new Date(sess.lastUpdated).toLocaleDateString()}</div>
                        <button class="hist-del" title="删除">🗑</button>
                    `;
                    item.querySelector(".hist-del")!.addEventListener("click", e => {
                        e.stopPropagation();
                        const all = ChatHistoryManager.loadHistory(this.dataView?.metadata?.objects);
                        ChatHistoryManager.saveHistory(this.host, all.filter(s => s.id !== sess.id));
                        item.remove();
                    });
                    item.addEventListener("click", e => {
                        if ((e.target as HTMLElement).classList.contains("hist-del")) return;
                        this.loadSession(sess);
                        modal.remove();
                    });
                    body.appendChild(item);
                });
        }

        box.querySelector(".modal-x")!.addEventListener("click", () => modal.remove());
        modal.addEventListener("click", e => { if (e.target === modal) modal.remove(); });
        modal.appendChild(box);
        document.body.appendChild(modal);
    }

    private showSettingsModal(): void {
        const objects = this.dataView?.metadata?.objects as any;
        const apiUrl = String(objects?.aiSettings?.apiUrl || this.formattingSettings.aiSettingsCard.apiUrl.value || "");
        const apiKey = this.unmaskString(String(objects?.aiSettings?.apiKey || this.formattingSettings.aiSettingsCard.apiKey.value || ""));
        const model  = String(objects?.aiSettings?.model  || this.formattingSettings.aiSettingsCard.model.value  || "");
        const tmdl   = TmdlManager.loadTmdl(this.dataView?.metadata?.objects);

        const modal = document.createElement("div");
        modal.className = "abi-modal";
        modal.innerHTML = `
            <div class="modal-box">
                <div class="modal-hdr">
                    <h3>设置</h3>
                    <button class="modal-x">×</button>
                </div>
                <div class="modal-body">
                    <div class="modal-field">
                        <label>Base URL</label>
                        <input id="s-url" type="text" value="${this.sanitizeHTML(apiUrl)}" placeholder="https://api.openai.com/v1/chat/completions"/>
                    </div>
                    <div class="modal-field">
                        <label>API Key</label>
                        <input id="s-key" type="password" value="${this.sanitizeHTML(apiKey)}" placeholder="sk-..."/>
                    </div>
                    <div class="modal-field">
                        <label>模型名称</label>
                        <input id="s-model" type="text" value="${this.sanitizeHTML(model)}" placeholder="gpt-4o"/>
                    </div>
                    <div class="modal-field">
                        <label>数据模型上下文 (TMDL / JSON / TXT)</label>
                        <textarea id="s-tmdl" style="height:160px">${this.sanitizeHTML(tmdl)}</textarea>
                    </div>
                </div>
                <div class="modal-ftr">
                    <button class="btn-secondary" id="s-cancel">取消</button>
                    <button class="btn-primary" id="s-save">保存</button>
                </div>
            </div>
        `;

        const close = () => modal.remove();
        modal.querySelector(".modal-x")!.addEventListener("click", close);
        modal.querySelector("#s-cancel")!.addEventListener("click", close);
        modal.addEventListener("click", e => { if (e.target === modal) close(); });

        modal.querySelector("#s-save")!.addEventListener("click", () => {
            const url   = (modal.querySelector("#s-url") as HTMLInputElement).value.trim();
            const key   = (modal.querySelector("#s-key") as HTMLInputElement).value.trim();
            const mdl   = (modal.querySelector("#s-model") as HTMLInputElement).value.trim();
            const tmdlv = (modal.querySelector("#s-tmdl") as HTMLTextAreaElement).value;

            this.host.persistProperties({
                merge: [{ objectName: "aiSettings", selector: null, properties: { apiUrl: url, apiKey: this.maskString(key), model: mdl } }]
            });
            TmdlManager.saveTmdl(this.host, tmdlv);
            close();
        });

        document.body.appendChild(modal);
    }

    // ── API Calls ─────────────────────────────────────────────────────────────

    private buildEnrichedMessage(userMessage: string): string {
        const ctx = this.prepareDataContext();
        if (!ctx) return userMessage;
        return (
            `数据上下文：\n${ctx}\n\n` +
            `用户问题：${userMessage}\n\n` +
            `请基于提供的数据回答用户问题。请务必使用HTML格式返回结果：` +
            `1. 使用<h3>、<h4>作为标题 ` +
            `2. 使用<table class="report-table">显示表格 ` +
            `3. 使用<ul>、<ol>显示列表 ` +
            `4. 换行请使用<br>或<p> ` +
            `5. 重点内容使用<span class="report-emphasis">标注</span> ` +
            `6. 不要使用Markdown格式，直接返回HTML代码`
        );
    }

    private getSystemPrompt(): string {
        return "你是一个专业的数据分析助手，帮助用户分析 PowerBI 报表中的数据，给出清晰、准确、有洞察力的分析结论。";
    }

    private async callAIAPI(userMessage: string): Promise<string> {
        const objects = this.dataView?.metadata?.objects as any;
        const apiUrl  = String(objects?.aiSettings?.apiUrl || this.formattingSettings.aiSettingsCard.apiUrl.value || "");
        const rawKey  = this.unmaskString(String(objects?.aiSettings?.apiKey || this.formattingSettings.aiSettingsCard.apiKey.value || ""));
        const model   = String(objects?.aiSettings?.model  || this.formattingSettings.aiSettingsCard.model.value  || "gpt-3.5-turbo");
        const isGemini = apiUrl.includes("googleapis.com") || model.toLowerCase().startsWith("gemini");

        const enriched = this.buildEnrichedMessage(userMessage);
        const history  = this.getConversationHistoryMessages();
        const sysPmt   = this.getSystemPrompt();

        if (isGemini) {
            const histText = history.map(m => `${m.role === "user" ? "User" : "Assistant"}: ${m.content}`).join("\n");
            const resp = await fetch(apiUrl, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: `${sysPmt}\n\n${histText}\nUser: ${enriched}` }] }],
                    generationConfig: { temperature: 0.7 }
                }),
                signal: this.abortController?.signal
            });
            const d = await resp.json();
            return d?.candidates?.[0]?.content?.parts?.[0]?.text || "";
        }

        const resp = await fetch(apiUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${rawKey}` },
            body: JSON.stringify({
                model,
                stream: false,
                temperature: 0.7,
                messages: [{ role: "system", content: sysPmt }, ...history, { role: "user", content: enriched }]
            }),
            signal: this.abortController?.signal
        });
        const d = await resp.json();
        return d?.choices?.[0]?.message?.content || "";
    }

    private async callAIAPIWithStreaming(userMessage: string, onChunk: (c: string) => void): Promise<string> {
        const objects  = this.dataView?.metadata?.objects as any;
        const apiUrl   = String(objects?.aiSettings?.apiUrl || this.formattingSettings.aiSettingsCard.apiUrl.value || "");
        const rawKey   = this.unmaskString(String(objects?.aiSettings?.apiKey || this.formattingSettings.aiSettingsCard.apiKey.value || ""));
        const model    = String(objects?.aiSettings?.model  || this.formattingSettings.aiSettingsCard.model.value  || "gpt-3.5-turbo");
        const isGemini = apiUrl.includes("googleapis.com") || model.toLowerCase().startsWith("gemini");

        // Gemini: no native SSE — call once then replay
        if (isGemini) {
            const full = await this.callAIAPI(userMessage);
            for (const c of full) {
                onChunk(c);
                await new Promise(r => setTimeout(r, 10));
            }
            return full;
        }

        const enriched = this.buildEnrichedMessage(userMessage);
        const history  = this.getConversationHistoryMessages();
        const sysPmt   = this.getSystemPrompt();

        const resp = await fetch(apiUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${rawKey}` },
            body: JSON.stringify({
                model,
                stream: true,
                temperature: 0.7,
                messages: [{ role: "system", content: sysPmt }, ...history, { role: "user", content: enriched }]
            }),
            signal: this.abortController?.signal
        });

        const reader  = resp.body!.getReader();
        const decoder = new TextDecoder();
        let accumulated = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            for (const line of decoder.decode(value, { stream: true }).split("\n")) {
                if (!line.startsWith("data: ")) continue;
                const payload = line.slice(6).trim();
                if (payload === "[DONE]") continue;
                try {
                    const t = JSON.parse(payload)?.choices?.[0]?.delta?.content;
                    if (t) { accumulated += t; onChunk(t); }
                } catch { /* skip malformed SSE */ }
            }
        }
        return accumulated;
    }

    private getConversationHistoryMessages(): Array<{ role: string; content: string }> {
        return this.chatMessages
            .filter(m => m.type === "user" || m.type === "assistant")
            .slice(0, -1)   // exclude the message currently being sent
            .slice(-10)
            .map(m => ({ role: m.type === "user" ? "user" : "assistant", content: m.content }));
    }

    // ── Data Processing ───────────────────────────────────────────────────────

    private prepareDataContext(): string {
        const parts: string[] = [];

        // TMDL Schema — prepend if present (static, filter-independent metadata)
        const tmdl = TmdlManager.loadTmdl(this.dataView?.metadata?.objects);
        if (tmdl) parts.push(`【数据模型上下文】\n${tmdl}`);

        const table = this.dataView?.table;
        if (!table) return parts.join("\n\n");

        const columns = table.columns;
        const rows    = table.rows;

        // Layer 1: Overview — describes what the user actually sees after slicers/filters
        const colNames = columns.map(c => c.displayName).join("、");
        parts.push(`数据概览：\n- 列数：${columns.length}\n- 行数：${rows.length}\n- 列名：${colNames}`);

        // Layer 2: Pre-computed statistics from the filtered row set
        const statLines: string[] = [];
        columns.forEach((col, ci) => {
            const nums = rows
                .map(r => r[ci])
                .filter(v => typeof v === "number" && !isNaN(v as number)) as number[];
            if (nums.length === 0) return;
            const name = col.displayName;
            const isRatio = /[%率占比比率]/.test(name);
            if (isRatio) {
                statLines.push(`- ${name}：[比率指标] (不可直接汇总)`);
            } else {
                statLines.push(`- ${name}：总计=${nums.reduce((a, b) => a + b, 0).toLocaleString()}`);
            }
        });
        if (statLines.length > 0) {
            parts.push(`【关键统计指标（已复核，请直接引用）】：\n${statLines.join("\n")}`);
        }

        // Layer 3: All filtered rows serialised line by line
        const rowLines = rows.map((row, ri) => {
            const cells = columns.map((col, ci) => `${col.displayName}: ${this.formatCellValue(row[ci])}`).join("、");
            return `${ri + 1}. ${cells}`;
        }).join("\n");
        if (rowLines) parts.push(rowLines);

        return parts.join("\n\n");
    }

    private executeDataQuery(queryObj: DataQuery): string {
        const table = this.dataView?.table;
        if (!table) return `<div class="report-section">暂无数据</div>`;

        const columns = table.columns;

        // Step 1: row → plain object
        let rows: Record<string, any>[] = table.rows.map(row => {
            const obj: Record<string, any> = {};
            columns.forEach((col, i) => { obj[col.displayName] = row[i]; });
            return obj;
        });

        // Step 2: filters
        if (queryObj.filters?.length) {
            rows = rows.filter(row => queryObj.filters!.every(f => {
                const val = row[f.column], fval = f.value;
                switch (f.operator) {
                    case ">":        return Number(val) > Number(fval);
                    case "<":        return Number(val) < Number(fval);
                    case ">=":       return Number(val) >= Number(fval);
                    case "<=":       return Number(val) <= Number(fval);
                    case "==":       return String(val) === String(fval);
                    case "!=":       return String(val) !== String(fval);
                    case "contains": return String(val).includes(String(fval));
                    default:         return true;
                }
            }));
        }

        // Step 3: groupBy + aggregations
        if (queryObj.groupBy?.length) {
            const groups = new Map<string, Record<string, any>[]>();
            rows.forEach(row => {
                const key = queryObj.groupBy!.map(g => String(row[g])).join("|");
                if (!groups.has(key)) groups.set(key, []);
                groups.get(key)!.push(row);
            });
            rows = Array.from(groups.values()).map(grp => {
                const result: Record<string, any> = {};
                queryObj.groupBy!.forEach(g => { result[g] = grp[0][g]; });
                (queryObj.aggregations || []).forEach(agg => {
                    const vals = grp.map(r => r[agg.column]).filter(v => v !== null && v !== undefined);
                    const nums = vals.map(Number).filter(v => !isNaN(v));
                    const k = `${agg.op}(${agg.column})`;
                    switch (agg.op) {
                        case "sum":   result[k] = nums.reduce((a, b) => a + b, 0); break;
                        case "avg":   result[k] = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0; break;
                        case "count": result[k] = vals.length; break;
                        case "max":   result[k] = nums.length ? Math.max(...nums) : ""; break;
                        case "min":   result[k] = nums.length ? Math.min(...nums) : ""; break;
                        case "first": result[k] = vals[0] ?? ""; break;
                    }
                });
                return result;
            });
        }

        // Step 4: sort
        if (queryObj.sort) {
            const { column, direction } = queryObj.sort;
            rows.sort((a, b) => {
                const an = Number(a[column]), bn = Number(b[column]);
                const cmp = !isNaN(an) && !isNaN(bn) ? an - bn : String(a[column]).localeCompare(String(b[column]));
                return direction === "desc" ? -cmp : cmp;
            });
        }

        // Step 5: limit
        if (queryObj.limit && queryObj.limit > 0) rows = rows.slice(0, queryObj.limit);

        if (rows.length === 0) return `<div class="report-section">查询结果 (共 0 条)</div>`;

        // Step 6: HTML table
        const headers = Object.keys(rows[0]);
        const thead = `<thead><tr>${headers.map(h => `<th>${this.formatCellValue(h)}</th>`).join("")}</tr></thead>`;
        const tbody = `<tbody>${rows.map(r =>
            `<tr>${headers.map(h => `<td>${this.formatCellValue(r[h])}</td>`).join("")}</tr>`
        ).join("")}</tbody>`;

        return `<div class="report-section">查询结果 (共 ${rows.length} 条)</div><table class="report-table">${thead}${tbody}</table>`;
    }

    private formatCellValue(value: any): string {
        if (value === null || value === undefined) return "";
        if (typeof value === "number") return value.toLocaleString();
        return String(value)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    private parseAIResponse(text: string): { text: string; chartData?: any } {
        const formattedText = this.formatToReportStyle(text);
        const chartData     = this.shouldGenerateChart(text) ? this.generateChartFromData(text) : undefined;
        return { text: formattedText, chartData };
    }

    private shouldGenerateChart(text: string): boolean {
        const kw = ["生成图表","创建图表","画图表","制作图表","显示图表","绘制图表",
                    "画个图表","做个图表","来个图表","要个图表",
                    "画柱状图","画折线图","画饼图","做柱状图","做折线图","做饼图",
                    "生成柱状图","生成折线图","生成饼图","创建柱状图","创建折线图","创建饼图",
                    "用图表显示","用图表展示","图表展示","图表呈现"];
        return kw.some(k => text.includes(k));
    }

    private generateChartFromData(text: string): any {
        const table = this.dataView?.table;
        if (!table || table.columns.length < 2 || table.rows.length === 0) return null;
        const labels = table.rows.map(r => String(r[0] ?? ""));
        const data   = table.rows.map(r => Number(r[1]) || 0);
        let type: "bar"|"line"|"pie"|"doughnut" = "bar";
        if (text.includes("折线") || text.includes("趋势")) type = "line";
        else if (text.includes("饼图")) type = "pie";
        else if (text.includes("环形")) type = "doughnut";
        const colors = ["#FF6384","#36A2EB","#FFCE56","#4BC0C0","#9966FF","#FF9F40","#FF6384","#C9CBCF"];
        return {
            type,
            data: { labels, datasets: [{ label: table.columns[1].displayName, data, backgroundColor: colors, borderColor: "#36A2EB", borderWidth: 1 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: true } } }
        };
    }

    // ── HTML Sanitisation & Formatting ────────────────────────────────────────

    private sanitizeHTML(html: string): string {
        const TAGS  = new Set(["h3","h4","p","br","div","span","table","thead","tbody","tr","th","td","ul","ol","li","b","strong","i","em","u"]);
        const ATTRS = new Set(["class","style","rowspan","colspan","width","height","align"]);
        try {
            const doc = new DOMParser().parseFromString(html, "text/html");
            const clean = (node: Node): Node | null => {
                if (node.nodeType === Node.TEXT_NODE) return node.cloneNode(true);
                if (node.nodeType !== Node.ELEMENT_NODE) return null;
                const el = node as Element;
                const tag = el.tagName.toLowerCase();
                if (!TAGS.has(tag)) {
                    const frag = document.createDocumentFragment();
                    el.childNodes.forEach(c => { const r = clean(c); if (r) frag.appendChild(r); });
                    return frag;
                }
                const ne = document.createElement(tag);
                Array.from(el.attributes).forEach(a => { if (ATTRS.has(a.name.toLowerCase())) ne.setAttribute(a.name, a.value); });
                el.childNodes.forEach(c => { const r = clean(c); if (r) ne.appendChild(r); });
                return ne;
            };
            const frag = document.createDocumentFragment();
            doc.body.childNodes.forEach(c => { const r = clean(c); if (r) frag.appendChild(r); });
            const w = document.createElement("div");
            w.appendChild(frag);
            return w.innerHTML;
        } catch {
            return (html || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        }
    }

    private formatContentForDisplay(text: string): string {
        const stripped = (text || "")
            .replace(/^```html\s*/i, "")
            .replace(/^```\s*/i, "")
            .replace(/\s*```$/, "")
            .trim();
        return this.sanitizeHTML(stripped);
    }

    private formatToReportStyle(text: string): string {
        const jsonBlockRe = /```json\s*(\{[\s\S]*?"intent"\s*:\s*"data_query"[\s\S]*?\})\s*```/;
        const m = text.match(jsonBlockRe);
        if (m) {
            try {
                const q: DataQuery = JSON.parse(m[1]);
                const before = text.slice(0, m.index);
                const after  = text.slice((m.index ?? 0) + m[0].length);
                return this.formatContentForDisplay(before) + this.executeDataQuery(q) + this.formatContentForDisplay(after);
            } catch { /* fall through */ }
        }
        return this.formatContentForDisplay(text);
    }

    // ── Session Management ────────────────────────────────────────────────────

    private startNewChat(): void {
        if (this.chatMessages.length > 0) this.saveCurrentChatToHistory();
        this.currentSessionId = this.generateSessionId();
        this.chatMessages = [];
        this.addWelcomeMessage();
        this.saveLastActiveSessionId(this.currentSessionId);
    }

    private saveCurrentChatToHistory(): void {
        if (!this.chatMessages.some(m => m.type === "user")) return;
        const objects  = this.dataView?.metadata?.objects;
        const sessions = ChatHistoryManager.loadHistory(objects);
        const title    = this.chatMessages.find(m => m.type === "user")?.content?.slice(0, 20) || "新会话";
        const payload: ChatSession = { id: this.currentSessionId, title, messages: this.chatMessages, lastUpdated: new Date() };
        const idx = sessions.findIndex(s => s.id === this.currentSessionId);
        if (idx >= 0) sessions[idx] = payload; else sessions.unshift(payload);
        ChatHistoryManager.saveHistory(this.host, sessions);
        this.saveLastActiveSessionId(this.currentSessionId);
    }

    private loadSession(session: ChatSession): void {
        this.currentSessionId = session.id;
        this.chatMessages = (session.messages || []).map(m => ({ ...m, timestamp: new Date(m.timestamp) }));
        this.renderChatMessages();
        this.saveLastActiveSessionId(session.id);
    }

    private saveLastActiveSessionId(id: string): void {
        this.host.persistProperties({
            merge: [{ objectName: "historySettings", selector: null, properties: { lastActiveSessionId: id } }]
        });
    }

    // ── Utilities ─────────────────────────────────────────────────────────────

    private generateSessionId(): string {
        return `${Date.now().toString(36)}${this.generateRandomString(8)}`;
    }

    private generateRandomString(n: number): string {
        const chars = "abcdefghijklmnopqrstuvwxyz0123456789";
        if (window.crypto?.getRandomValues) {
            const arr = new Uint8Array(n);
            window.crypto.getRandomValues(arr);
            return Array.from(arr, v => chars[v % chars.length]).join("");
        }
        return Array.from({ length: n }, () => chars[Math.floor(Math.random() * chars.length)]).join("");
    }

    private maskString(s: string): string {
        return `ENC_${btoa(encodeURIComponent(s || ""))}`;
    }

    private unmaskString(s: string): string {
        if (!s || !s.startsWith("ENC_")) return s || "";
        try { return decodeURIComponent(atob(s.slice(4))); } catch { return s; }
    }

    private copyToClipboard(text: string): void {
        navigator.clipboard?.writeText(text).catch(() => {
            const el = document.createElement("textarea");
            el.value = text;
            document.body.appendChild(el);
            el.select();
            document.execCommand("copy");
            el.remove();
        });
    }
}
