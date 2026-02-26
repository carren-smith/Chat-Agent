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

interface TableRow {
    [columnName: string]: string | number | null;
}

interface FilterInfo {
    table: string;
    column: string;
    values: string[];
    filterType: string;
}

interface DataQuery {
    intent: "data_query";
    filters?: Array<{ column: string; operator: ">" | "<" | ">=" | "<=" | "==" | "!=" | "contains"; value: string | number }>;
    groupBy?: string[];
    aggregations?: Array<{ column: string; op: "sum" | "avg" | "count" | "max" | "min" | "first" }>;
    sort?: { column: string; direction: "asc" | "desc" };
    limit?: number;
}

class TmdlManager {
    private static readonly MAX_CHUNK_SIZE: number = 25000;
    private static readonly MAX_CHUNKS: number = 50;

    public static cleanTmdl(code: string): string {
        return (code || "")
            .split("\n")
            .filter((line) => !line.trim().startsWith("//") && !line.includes("annotation"))
            .join("\n")
            .replace(/\bmodifiedTime\s*=.*$/gm, "")
            .trim();
    }

    public static saveTmdl(host: IVisualHost, code: string): void {
        const cleanCode = this.cleanTmdl(code);
        const chunks = [];
        for (let i = 0; i < cleanCode.length && chunks.length < this.MAX_CHUNKS; i += this.MAX_CHUNK_SIZE) {
            chunks.push(cleanCode.slice(i, i + this.MAX_CHUNK_SIZE));
        }

        const properties: any = { chunkCount: String(chunks.length), tmdlCode: cleanCode };
        for (let i = 0; i < this.MAX_CHUNKS; i++) {
            properties[`chunk${i}`] = chunks[i] || "";
        }

        host.persistProperties({
            merge: [{ objectName: "tmdlSettings", properties, selector: null }]
        });
    }

    public static loadTmdl(objects?: DataViewObjects): string {
        const settings = objects && (objects as any).tmdlSettings;
        if (!settings) return "";

        const count = Number(settings.chunkCount || 0);
        if (count > 0) {
            const parts: string[] = [];
            for (let i = 0; i < count; i++) {
                parts.push(String(settings[`chunk${i}`] || ""));
            }
            return parts.join("");
        }

        return String(settings.tmdlCode || "");
    }
}

class ChatHistoryManager {
    private static readonly MAX_CHUNK_SIZE: number = 25000;
    private static readonly MAX_CHUNKS: number = 50;

    public static saveHistory(host: IVisualHost, sessions: ChatSession[]): void {
        const payload = JSON.stringify(sessions);
        const chunks = [];

        for (let i = 0; i < payload.length && chunks.length < this.MAX_CHUNKS; i += this.MAX_CHUNK_SIZE) {
            chunks.push(payload.slice(i, i + this.MAX_CHUNK_SIZE));
        }

        const properties: any = { chunkCount: String(chunks.length) };
        for (let i = 0; i < this.MAX_CHUNKS; i++) {
            properties[`chunk${i}`] = chunks[i] || "";
        }

        host.persistProperties({ merge: [{ objectName: "historySettings", properties, selector: null }] });
    }

    public static loadHistory(objects?: DataViewObjects): ChatSession[] {
        const settings = objects && (objects as any).historySettings;
        if (!settings) return [];

        const count = Number(settings.chunkCount || 0);
        if (count <= 0) return [];

        let json = "";
        for (let i = 0; i < count; i++) {
            json += String(settings[`chunk${i}`] || "");
        }

        try {
            const parsed = JSON.parse(json) as ChatSession[];
            return parsed.map((session) => ({
                ...session,
                lastUpdated: new Date(session.lastUpdated),
                messages: (session.messages || []).map((m) => ({ ...m, timestamp: new Date(m.timestamp) }))
            }));
        } catch {
            return [];
        }
    }
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: FormattingSettings;
    private formattingSettingsService: FormattingSettingsService;
    private dataView: DataView;
    private host: IVisualHost;
    private storageService: any;

    private chatMessages: ChatMessage[];
    private currentSessionId: string;
    private charts: Map<string, any>;

    private container: HTMLElement;
    private messagesContainer: HTMLElement;
    private chatHeader: HTMLElement;
    private inputElement: HTMLTextAreaElement;
    private sendButton: HTMLButtonElement;

    private settingsButton: HTMLButtonElement;
    private newChatButton: HTMLButtonElement;
    private historyButton: HTMLButtonElement;

    private currentStreamingMessageIndex: number = -1;
    private streamingMessageElement: HTMLElement | null = null;

    private abortController: AbortController | null = null;
    private isGenerating: boolean = false;
    private isComposing: boolean = false;

    private licenseKey: string = "";
    private activeSystemSecret: string = "";
    private currentDeviceId: string = "";
    private systemPrompt: string = "";

    private isDesktopEnv: boolean;
    private lastValidationError: string = "";
    private validationDebounceTimer: number = 0;
    private isValidationInProgress: boolean = false;
    private validationPromise: Promise<boolean> | null = null;

    private reportContext: { filters: FilterInfo[]; tableData: TableRow[] } = { filters: [], tableData: [] };

    private get isLicenseValid(): boolean {
        return !!this.activeSystemSecret && this.activeSystemSecret.length > 5;
    }

    constructor(t: VisualConstructorOptions) {
        console.log("Visual constructor", t);
        this.target = t.element;
        this.host = t.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.isDesktopEnv = this.checkIsDesktop();
        this.storageService = (this.host as any).storageService;
        if (!this.storageService) {
            console.warn("storageService unavailable");
        }
        this.formattingSettings = new FormattingSettings();
        this.chatMessages = [];
        this.currentSessionId = this.generateSessionId();
        this.charts = new Map();
        this.initializeDOM();
        this.setupEventListeners();
    }

    private checkIsDesktop(): boolean {
        const host = window.location.hostname.toLowerCase();
        if (host === "localhost" || host === "127.0.0.1") return true;
        return !(host.includes("powerbi.com") || host.includes("powerbi.cn") || host.includes("analysis.windows.net"));
    }

    private initializeDOM(): void {
        this.target.innerHTML = "";
        const style = document.createElement("style");
        style.textContent = `.ai-chat-container{height:100%;display:flex}.chat-container{display:flex;flex-direction:column;width:100%;height:100%;font-family:Segoe UI}.chat-header{display:flex;justify-content:space-between;padding:8px;border-bottom:1px solid #eee}.chat-messages{flex:1;overflow:auto;padding:12px}.message{margin:8px 0}.message.user{text-align:right}.chat-input-wrapper{display:flex;gap:8px;padding:8px;border-top:1px solid #eee}.chat-input{flex:1;resize:none}.send-button.generating{opacity:.7}.suggestions-container{display:flex;gap:8px;flex-wrap:wrap}.tmdl-modal{position:fixed;inset:0;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center}`;
        this.target.appendChild(style);

        this.container = document.createElement("div");
        this.container.className = "ai-chat-container";
        this.container.innerHTML = `<div class="chat-container"><div class="chat-header"><div class="header-title">ABI Chat</div><div class="header-actions"></div></div><div class="chat-messages"></div><div class="chat-input-wrapper"></div></div>`;

        this.chatHeader = this.container.querySelector(".chat-header") as HTMLElement;
        this.messagesContainer = this.container.querySelector(".chat-messages") as HTMLElement;
        const actions = this.container.querySelector(".header-actions") as HTMLElement;
        const inputWrapper = this.container.querySelector(".chat-input-wrapper") as HTMLElement;

        this.newChatButton = document.createElement("button");
        this.newChatButton.id = "newChatBtn";
        this.newChatButton.title = "新建任务";
        this.newChatButton.textContent = "+";

        this.historyButton = document.createElement("button");
        this.historyButton.id = "historyBtn";
        this.historyButton.title = "历史任务";
        this.historyButton.textContent = "🕒";

        this.settingsButton = document.createElement("button");
        this.settingsButton.id = "settingsBtn";
        this.settingsButton.title = "设置";
        this.settingsButton.style.display = "none";
        this.settingsButton.textContent = "⚙";

        actions.append(this.newChatButton, this.historyButton, this.settingsButton);

        this.inputElement = document.createElement("textarea");
        this.inputElement.className = "chat-input";
        this.inputElement.rows = 1;
        this.inputElement.placeholder = "请输入您的问题...";
        this.sendButton = document.createElement("button");
        this.sendButton.className = "send-button";
        this.sendButton.textContent = "发送";
        inputWrapper.append(this.inputElement, this.sendButton);
        this.target.appendChild(this.container);
    }

    private setupEventListeners(): void {
        this.sendButton.addEventListener("click", () => this.sendMessage());
        this.inputElement.addEventListener("compositionstart", () => (this.isComposing = true));
        this.inputElement.addEventListener("compositionend", () => (this.isComposing = false));
        this.inputElement.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey && !this.isComposing) {
                e.preventDefault();
                this.sendMessage();
            }
        });
        this.newChatButton.addEventListener("click", () => this.startNewChat());
        this.historyButton.addEventListener("click", () => this.showHistoryModal());
        this.settingsButton.addEventListener("click", () => this.showTmdlModal());
        this.inputElement.addEventListener("input", () => this.adjustTextareaHeight());
    }

    public update(options: VisualUpdateOptions): void {
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(FormattingSettings, options.dataViews?.[0]);
        this.formattingSettings.aboutCardSettings.revertToDefault();
        this.dataView = options.dataViews?.[0];

        const objects = this.dataView?.metadata?.objects as any;
        const rawLicenseKey = objects?.licenseSettings?.licenseKey || "";
        const licenseKey = this.unmaskString(String(rawLicenseKey));
        const viewMode = (options as any).viewMode;

        if (licenseKey !== this.licenseKey || !this.isLicenseValid) {
            this.licenseKey = licenseKey;
            this.validateLicense(licenseKey, viewMode);
        }

        this.settingsButton.style.display = viewMode === 1 && this.isDesktopEnv ? "inline-block" : "none";

        if (this.chatMessages.length === 1 && this.chatMessages[0].type === "assistant") {
            this.chatMessages[0].content = this.getWelcomeMessage();
            this.renderChatMessages();
        }

        if (this.chatMessages.length === 0) {
            const sessions = ChatHistoryManager.loadHistory(objects);
            const lastId = objects?.historySettings?.lastActiveSessionId;
            const fromLast = sessions.find((s) => s.id === lastId);
            if (fromLast) {
                this.loadSession(fromLast);
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

    private async sendMessage(): Promise<void> {
        const text = this.inputElement.value.trim();
        if (!text) return;

        const valid = await this.ensureLicenseValidated();
        if (!valid) return;

        this.addMessage("user", text);
        this.inputElement.value = "";
        this.adjustTextareaHeight();

        this.isGenerating = true;
        this.abortController = new AbortController();
        this.updateSendButtonState();

        try {
            this.startStreamingMessage();
            const fullResponse = await this.callAIAPIWithStreaming(text, (chunk) => this.updateStreamingMessage(chunk));
            this.finalizeStreamingMessage(fullResponse);
            this.saveCurrentChatToHistory();
        } catch (err: any) {
            if (err?.name === "AbortError") {
                this.saveCurrentChatToHistory();
            } else {
                this.removeStreamingMessage();
                this.showErrorMessage(String(err?.message || err || "请求失败"));
            }
        } finally {
            this.isGenerating = false;
            this.abortController = null;
            this.updateSendButtonState();
        }
    }

    private addMessage(type: ChatMessage["type"], content: string): void {
        this.chatMessages.push({ type, content, timestamp: new Date() });
        this.renderChatMessages();
    }

    private startStreamingMessage(): void {
        this.chatMessages.push({ type: "assistant", content: "", timestamp: new Date() });
        this.currentStreamingMessageIndex = this.chatMessages.length - 1;
        this.renderChatMessages();
    }

    private updateStreamingMessage(chunk: string): void {
        if (this.currentStreamingMessageIndex < 0) return;
        const msg = this.chatMessages[this.currentStreamingMessageIndex];
        msg.content += chunk;
        this.renderChatMessages();
    }

    private finalizeStreamingMessage(fullText: string): void {
        if (this.currentStreamingMessageIndex < 0) return;
        const parsed = this.parseAIResponse(fullText);
        this.chatMessages[this.currentStreamingMessageIndex].content = parsed.text;
        this.chatMessages[this.currentStreamingMessageIndex].chartData = parsed.chartData;
        this.currentStreamingMessageIndex = -1;
        this.streamingMessageElement = null;
        this.renderChatMessages();
        this.addSuggestionChips();
    }

    private removeStreamingMessage(): void {
        if (this.currentStreamingMessageIndex >= 0) {
            this.chatMessages.splice(this.currentStreamingMessageIndex, 1);
            this.currentStreamingMessageIndex = -1;
            this.streamingMessageElement = null;
            this.renderChatMessages();
        }
    }

    private renderChatMessages(): void {
        this.messagesContainer.innerHTML = "";
        this.chatMessages.forEach((message) => {
            const row = document.createElement("div");
            row.className = `message ${message.type}`;
            row.innerHTML = `<div class="message-content">${this.formatToReportStyle(message.content)}</div>`;
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
        const objects = this.dataView?.metadata?.objects as any;
        return objects?.welcomeSettings?.welcomeMessage || this.formattingSettings.welcomeSettingsCard.welcomeMessage.value || "我是PowerBI星球打造的ABI Chat，欢迎使用。";
    }

    private addSuggestionChips(): void {
        setTimeout(() => {
            const wrap = document.createElement("div");
            wrap.className = "suggestions-container";
            const card = this.formattingSettings.suggestionSettingsCard;
            [card.question1.value, card.question2.value, card.question3.value].forEach((q) => {
                if (!q) return;
                const btn = document.createElement("button");
                btn.className = "suggestion-btn";
                btn.textContent = q;
                btn.onclick = () => {
                    this.inputElement.value = q;
                    this.sendMessage();
                };
                wrap.appendChild(btn);
            });
            this.messagesContainer.appendChild(wrap);
        }, 50);
    }

    private showHistoryModal(): void {
        const modal = document.createElement("div");
        modal.className = "tmdl-modal history-modal";
        const sessions = ChatHistoryManager.loadHistory(this.dataView?.metadata?.objects);
        modal.innerHTML = `<div style="background:#fff;padding:16px;max-width:640px;width:90%"><h3>历史任务</h3><div>${sessions
            .map((s) => `<div data-id="${s.id}">${this.sanitizeHTML(s.title)} <button data-load="${s.id}">加载</button></div>`)
            .join("")}</div><button id="closeHistory">关闭</button></div>`;
        modal.addEventListener("click", (e: any) => {
            const id = e.target?.getAttribute("data-load");
            if (id) {
                const session = sessions.find((s) => s.id === id);
                if (session) this.loadSession(session);
                modal.remove();
            }
            if (e.target?.id === "closeHistory" || e.target === modal) modal.remove();
        });
        document.body.appendChild(modal);
    }

    private showTmdlModal(): void {
        const modal = document.createElement("div");
        modal.className = "tmdl-modal";
        const objects = this.dataView?.metadata?.objects as any;
        const currentTmdl = TmdlManager.loadTmdl(objects);
        modal.innerHTML = `<div style="background:#fff;padding:16px;max-width:760px;width:95%"><h3>设置</h3><textarea id="licenseInput" style="width:100%" placeholder="License">${this.sanitizeHTML(this.licenseKey)}</textarea><input id="apiUrlInput" value="${this.sanitizeHTML(objects?.aiSettings?.apiUrl || this.formattingSettings.aiSettingsCard.apiUrl.value)}"/><input id="apiKeyInput" value="${this.sanitizeHTML(objects?.aiSettings?.apiKey || this.formattingSettings.aiSettingsCard.apiKey.value)}"/><input id="modelInput" value="${this.sanitizeHTML(objects?.aiSettings?.model || this.formattingSettings.aiSettingsCard.model.value)}"/><textarea id="tmdlInput" style="width:100%;height:180px">${this.sanitizeHTML(currentTmdl)}</textarea><button id="saveBtn">保存</button><button id="closeBtn">关闭</button></div>`;
        modal.addEventListener("click", async (e: any) => {
            if (e.target?.id === "closeBtn" || e.target === modal) modal.remove();
            if (e.target?.id === "saveBtn") {
                const license = (modal.querySelector("#licenseInput") as HTMLTextAreaElement).value;
                const apiUrl = (modal.querySelector("#apiUrlInput") as HTMLInputElement).value;
                const apiKey = (modal.querySelector("#apiKeyInput") as HTMLInputElement).value;
                const model = (modal.querySelector("#modelInput") as HTMLInputElement).value;
                const tmdl = (modal.querySelector("#tmdlInput") as HTMLTextAreaElement).value;

                this.host.persistProperties({
                    merge: [
                        { objectName: "aiSettings", selector: null, properties: { apiUrl, apiKey, model } },
                        { objectName: "licenseSettings", selector: null, properties: { licenseKey: this.maskString(license), boundFingerprint: this.getOrCreateDeviceId() } }
                    ]
                });
                TmdlManager.saveTmdl(this.host, tmdl);
                await this.validateLicense(license, 1);
                modal.remove();
            }
        });
        document.body.appendChild(modal);
    }

    private adjustTextareaHeight(): void {
        this.inputElement.style.height = "auto";
        this.inputElement.style.height = `${Math.min(this.inputElement.scrollHeight, 180)}px`;
        this.updateSendButtonState();
    }

    private updateSendButtonState(): void {
        this.sendButton.disabled = this.isGenerating || !this.inputElement.value.trim();
        this.sendButton.classList.toggle("generating", this.isGenerating);
    }

    private buildEnrichedMessage(userMessage: string): string {
        const dataContext = this.prepareDataContext();
        if (!dataContext) return userMessage;
        return (
            `数据上下文：\n${dataContext}\n\n` +
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

    private async callAIAPI(userMessage: string): Promise<string> {
        const objects = this.dataView?.metadata?.objects as any;
        const apiUrl = String(objects?.aiSettings?.apiUrl || this.formattingSettings.aiSettingsCard.apiUrl.value || "");
        const rawKey = this.unmaskString(String(objects?.aiSettings?.apiKey || this.formattingSettings.aiSettingsCard.apiKey.value || ""));
        const model  = String(objects?.aiSettings?.model  || this.formattingSettings.aiSettingsCard.model.value  || "gpt-3.5-turbo");
        const isGemini = apiUrl.includes("googleapis.com") || model.toLowerCase().startsWith("gemini");

        const systemPrompt = this.getPowerBIStarSystemPrompt();
        const history      = this.getConversationHistoryMessages();
        const enriched     = this.buildEnrichedMessage(userMessage);

        if (isGemini) {
            const historyText = history.map(m => `${m.role === "user" ? "User" : "Assistant"}: ${m.content}`).join("\n");
            const resp = await fetch(apiUrl, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: `${systemPrompt}\n\n${historyText}\nUser: ${enriched}` }] }],
                    generationConfig: { temperature: 0.7 }
                }),
                signal: this.abortController?.signal
            });
            const data = await resp.json();
            return data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
        }

        const resp = await fetch(apiUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${rawKey}` },
            body: JSON.stringify({
                model,
                stream: false,
                temperature: 0.7,
                messages: [
                    { role: "system", content: systemPrompt },
                    ...history,
                    { role: "user", content: enriched }
                ]
            }),
            signal: this.abortController?.signal
        });
        const data = await resp.json();
        return data?.choices?.[0]?.message?.content || "";
    }

    private async callAIAPIWithStreaming(userMessage: string, onChunk: (chunk: string) => void): Promise<string> {
        const objects = this.dataView?.metadata?.objects as any;
        const apiUrl = String(objects?.aiSettings?.apiUrl || this.formattingSettings.aiSettingsCard.apiUrl.value || "");
        const rawKey = this.unmaskString(String(objects?.aiSettings?.apiKey || this.formattingSettings.aiSettingsCard.apiKey.value || ""));
        const model  = String(objects?.aiSettings?.model  || this.formattingSettings.aiSettingsCard.model.value  || "gpt-3.5-turbo");
        const isGemini = apiUrl.includes("googleapis.com") || model.toLowerCase().startsWith("gemini");

        // Gemini: no native streaming — call non-streaming then replay char-by-char
        if (isGemini) {
            const full = await this.callAIAPI(userMessage);
            for (const c of full) {
                onChunk(c);
                await new Promise(r => setTimeout(r, 10));
            }
            return full;
        }

        // OpenAI-compatible: SSE streaming
        const systemPrompt = this.getPowerBIStarSystemPrompt();
        const history      = this.getConversationHistoryMessages();
        const enriched     = this.buildEnrichedMessage(userMessage);

        const resp = await fetch(apiUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json", Authorization: `Bearer ${rawKey}` },
            body: JSON.stringify({
                model,
                stream: true,
                temperature: 0.7,
                messages: [
                    { role: "system", content: systemPrompt },
                    ...history,
                    { role: "user", content: enriched }
                ]
            }),
            signal: this.abortController?.signal
        });

        const reader  = resp.body!.getReader();
        const decoder = new TextDecoder();
        let accumulated = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            const lines = decoder.decode(value, { stream: true }).split("\n");
            for (const line of lines) {
                if (!line.startsWith("data: ")) continue;
                const payload = line.slice(6).trim();
                if (payload === "[DONE]") continue;
                try {
                    const parsed = JSON.parse(payload);
                    const text = parsed?.choices?.[0]?.delta?.content;
                    if (text) { accumulated += text; onChunk(text); }
                } catch { /* skip malformed SSE line */ }
            }
        }

        return accumulated;
    }

    private getConversationHistoryMessages(): Array<{ role: string; content: string }> {
        return this.chatMessages
            .filter((m) => m.type === "user" || m.type === "assistant")
            .slice(-10)
            .map((m) => ({ role: m.type === "user" ? "user" : "assistant", content: m.content }));
    }

    private getPowerBIStarSystemPrompt(): string {
        return this.getExpertPrompt();
    }

    private getDeviceFingerprint(): string {
        const raw = `${navigator.userAgent}|${navigator.language}|${navigator.platform}|${navigator.hardwareConcurrency}|${Intl.DateTimeFormat().resolvedOptions().timeZone}`;
        let hash = 5381;
        for (let i = 0; i < raw.length; i++) hash = (hash * 33) ^ raw.charCodeAt(i);
        return Math.abs(hash).toString(36);
    }

    private getOrCreateDeviceId(): string {
        if (this.currentDeviceId) return this.currentDeviceId;
        const key = "abi_visual_fingerprint";
        const existing = localStorage.getItem(key);
        if (existing) {
            this.currentDeviceId = existing;
            return existing;
        }
        const created = this.getDeviceFingerprint();
        localStorage.setItem(key, created);
        this.currentDeviceId = created;
        return created;
    }

    private async ensureLicenseValidated(): Promise<boolean> {
        if (this.isLicenseValid) return true;
        if (!this.licenseKey) {
            this.handleInvalidLicense(1, "缺少许可证");
            return false;
        }
        if (this.validationPromise) return this.validationPromise;
        this.validationPromise = this.validateLicense(this.licenseKey, 1).finally(() => (this.validationPromise = null));
        return this.validationPromise;
    }

    private async validateLicense(key: string, viewMode: number): Promise<boolean> {
        if (!key || key.length < 8) {
            this.handleInvalidLicense(viewMode, "许可证格式错误");
            return false;
        }
        if (this.validationDebounceTimer) window.clearTimeout(this.validationDebounceTimer);
        return new Promise((resolve) => {
            this.validationDebounceTimer = window.setTimeout(async () => {
                resolve(this.performLicenseValidation(key, viewMode));
            }, 500);
        });
    }

    private async performLicenseValidation(key: string, viewMode: number): Promise<boolean> {
        try {
            this.isValidationInProgress = true;
            const endpoint = atob("aHR0cHM6Ly9hYmktY2hhdC12ZXJpZnkteXhvcmFuYmhlYS5jbi1oYW5nemhvdS5mY2FwcC5ydW4=");
            const response = await fetch(endpoint, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ licenseKey: key, fingerprint: this.getOrCreateDeviceId() })
            });
            if (!response.ok) throw new Error("校验失败");
            const data = await response.json();
            this.activeSystemSecret = String(data.secret || "");
            this.systemPrompt = this.decryptPrompt(String(data.prompt || ""), key);
            this.lastValidationError = "";
            return true;
        } catch (err: any) {
            this.handleInvalidLicense(viewMode, String(err?.message || err));
            return false;
        } finally {
            this.isValidationInProgress = false;
        }
    }

    private handleInvalidLicense(viewMode: number, err?: string): void {
        this.activeSystemSecret = "";
        this.systemPrompt = "";
        this.lastValidationError = err || "许可证无效";
        if (viewMode === 1) {
            this.showErrorMessage(`许可证无效，请联系微信 powerai001。${this.lastValidationError}`);
        }
    }

    private decryptPrompt(data: string, key: string): string {
        try {
            const bytes = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
            const keyBytes = new TextEncoder().encode(key);
            const out = bytes.map((b, i) => b ^ keyBytes[i % keyBytes.length]);
            return new TextDecoder().decode(out);
        } catch {
            return "";
        }
    }

    private getExpertPrompt(): string {
        return this.isLicenseValid ? this.systemPrompt : "这是演示模式回答。";
    }

    private maskString(s: string): string {
        return `ENC_${btoa(encodeURIComponent(s || ""))}`;
    }

    private unmaskString(s: string): string {
        if (!s) return "";
        if (!s.startsWith("ENC_")) return s;
        try {
            return decodeURIComponent(atob(s.slice(4)));
        } catch {
            return s;
        }
    }

    private sanitizeHTML(html: string): string {
        const ALLOWED_TAGS = new Set(["h3","h4","p","br","div","span","table","thead","tbody","tr","th","td","ul","ol","li","b","strong","i","em","u"]);
        const ALLOWED_ATTRS = new Set(["class","style","rowspan","colspan","width","height","align"]);
        try {
            const doc = new DOMParser().parseFromString(html, "text/html");
            const clean = (node: Node): Node | null => {
                if (node.nodeType === Node.TEXT_NODE) return node.cloneNode(true);
                if (node.nodeType !== Node.ELEMENT_NODE) return null;
                const el = node as Element;
                const tag = el.tagName.toLowerCase();
                if (!ALLOWED_TAGS.has(tag)) {
                    const frag = document.createDocumentFragment();
                    el.childNodes.forEach(c => { const r = clean(c); if (r) frag.appendChild(r); });
                    return frag;
                }
                const newEl = document.createElement(tag);
                Array.from(el.attributes).forEach(attr => {
                    if (ALLOWED_ATTRS.has(attr.name.toLowerCase())) newEl.setAttribute(attr.name, attr.value);
                });
                el.childNodes.forEach(c => { const r = clean(c); if (r) newEl.appendChild(r); });
                return newEl;
            };
            const frag = document.createDocumentFragment();
            doc.body.childNodes.forEach(c => { const r = clean(c); if (r) frag.appendChild(r); });
            const wrapper = document.createElement("div");
            wrapper.appendChild(frag);
            return wrapper.innerHTML;
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
        // Detect embedded JSON data_query block and execute it client-side
        const jsonBlockRegex = /```json\s*(\{[\s\S]*?"intent"\s*:\s*"data_query"[\s\S]*?\})\s*```/;
        const match = text.match(jsonBlockRegex);
        if (match) {
            try {
                const queryObj: DataQuery = JSON.parse(match[1]);
                const before = text.slice(0, match.index);
                const after = text.slice((match.index ?? 0) + match[0].length);
                const queryResult = this.executeDataQuery(queryObj);
                return this.formatContentForDisplay(before) + queryResult + this.formatContentForDisplay(after);
            } catch {
                // Fall through to default rendering
            }
        }
        return this.formatContentForDisplay(text);
    }

    private showErrorMessage(message: string): void {
        this.addMessage("error", message);
    }

    private parseAIResponse(text: string): { text: string; chartData?: any } {
        const formattedText = this.formatToReportStyle(text);
        const chartData = this.shouldGenerateChart(text) ? this.generateChartFromData(text) : undefined;
        return { text: formattedText, chartData };
    }

    private shouldGenerateChart(text: string): boolean {
        const keywords = [
            "生成图表","创建图表","画图表","制作图表","显示图表","绘制图表",
            "画个图表","做个图表","来个图表","要个图表",
            "画柱状图","画折线图","画饼图","做柱状图","做折线图","做饼图",
            "生成柱状图","生成折线图","生成饼图","创建柱状图","创建折线图","创建饼图",
            "用图表显示","用图表展示","图表展示","图表呈现"
        ];
        return keywords.some(k => text.includes(k));
    }

    private generateChartFromData(text: string): any {
        const table = this.dataView?.table;
        if (!table || table.columns.length < 2 || table.rows.length === 0) return null;

        const labels = table.rows.map(r => String(r[0] ?? ""));
        const data = table.rows.map(r => Number(r[1]) || 0);

        let type: "bar" | "line" | "pie" | "doughnut" = "bar";
        if (text.includes("折线") || text.includes("趋势")) type = "line";
        else if (text.includes("饼图")) type = "pie";
        else if (text.includes("环形")) type = "doughnut";

        const colors = ["#FF6384","#36A2EB","#FFCE56","#4BC0C0","#9966FF","#FF9F40","#FF6384","#C9CBCF"];

        return {
            type,
            data: {
                labels,
                datasets: [{
                    label: table.columns[1].displayName,
                    data,
                    backgroundColor: colors,
                    borderColor: "#36A2EB",
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: true } }
            }
        };
    }

    private startNewChat(): void {
        if (this.chatMessages.length > 0) this.saveCurrentChatToHistory();
        this.currentSessionId = this.generateSessionId();
        this.chatMessages = [];
        this.addWelcomeMessage();
        this.saveLastActiveSessionId(this.currentSessionId);
    }

    private saveCurrentChatToHistory(): void {
        const objects = this.dataView?.metadata?.objects;
        const sessions = ChatHistoryManager.loadHistory(objects);
        const title = this.chatMessages.find((m) => m.type === "user")?.content?.slice(0, 20) || "新会话";
        const existing = sessions.find((s) => s.id === this.currentSessionId);
        const payload: ChatSession = { id: this.currentSessionId, title, messages: this.chatMessages, lastUpdated: new Date() };

        if (existing) {
            Object.assign(existing, payload);
        } else {
            sessions.push(payload);
        }

        ChatHistoryManager.saveHistory(this.host, sessions);
        this.saveLastActiveSessionId(this.currentSessionId);
    }

    private loadSession(session: ChatSession): void {
        this.currentSessionId = session.id;
        this.chatMessages = (session.messages || []).map((m) => ({ ...m, timestamp: new Date(m.timestamp) }));
        this.renderChatMessages();
        this.saveLastActiveSessionId(session.id);
        this.saveCurrentChatToHistory();
    }

    private saveLastActiveSessionId(id: string): void {
        this.host.persistProperties({ merge: [{ objectName: "historySettings", selector: null, properties: { lastActiveSessionId: id } }] });
    }

    private generateSessionId(): string {
        return `${Date.now().toString(36)}${this.generateRandomString(8)}`;
    }

    private generateRandomString(n: number): string {
        const chars = "abcdefghijklmnopqrstuvwxyz0123456789";
        if (window.crypto?.getRandomValues) {
            const arr = new Uint8Array(n);
            window.crypto.getRandomValues(arr);
            return Array.from(arr).map((v) => chars[v % chars.length]).join("");
        }
        return Array.from({ length: n }).map(() => chars[Math.floor(Math.random() * chars.length)]).join("");
    }

    private prepareDataContext(): string {
        const parts: string[] = [];

        // TMDL Schema — prepend if present (static, filter-independent metadata)
        const tmdl = TmdlManager.loadTmdl(this.dataView?.metadata?.objects);
        if (tmdl) parts.push(`【数据模型上下文】\n${tmdl}`);

        const table = this.dataView?.table;
        // No table bound to this visual — return whatever we have (TMDL only or empty)
        if (!table) return parts.join("\n\n");

        const columns = table.columns;
        const rows = table.rows;

        // ── Layer 1: Data Overview ──────────────────────────────────────
        // dataView.table already reflects the current slicer / filter context;
        // row count and column list here describe what the user actually sees.
        const colNames = columns.map(c => c.displayName).join("、");
        parts.push(`数据概览：\n- 列数：${columns.length}\n- 行数：${rows.length}\n- 列名：${colNames}`);

        // ── Layer 2: Pre-computed Key Statistics ────────────────────────
        // Computed from filtered rows so values match what the user sees on screen.
        // Percentage / ratio columns are flagged to prevent incorrect aggregation.
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
                const total = nums.reduce((a, b) => a + b, 0);
                statLines.push(`- ${name}：总计=${total.toLocaleString()}`);
            }
        });
        if (statLines.length > 0) {
            parts.push(`【关键统计指标（已复核，请直接引用）】：\n${statLines.join("\n")}`);
        }

        // ── Layer 3: All Filtered Rows ──────────────────────────────────
        // Serialised line by line; row set is already narrowed by PowerBI filters/slicers.
        const rowLines = rows.map((row, ri) => {
            const cells = columns.map((col, ci) => `${col.displayName}: ${this.formatCellValue(row[ci])}`).join("、");
            return `${ri + 1}. ${cells}`;
        }).join("\n");
        if (rowLines) parts.push(rowLines);

        return parts.join("\n\n");
    }

    private executeDataQuery(queryObj: DataQuery): string {
        const table = this.dataView?.table;
        if (!table) return "<div class=\"report-section\">暂无数据</div>";

        const columns = table.columns;

        // ── Step 1: Convert to plain objects ───────────────────────────
        let rows: Record<string, any>[] = table.rows.map(row => {
            const obj: Record<string, any> = {};
            columns.forEach((col, i) => { obj[col.displayName] = row[i]; });
            return obj;
        });

        // ── Step 2: Apply filters ──────────────────────────────────────
        if (queryObj.filters?.length) {
            rows = rows.filter(row =>
                queryObj.filters!.every(f => {
                    const val = row[f.column];
                    const fval = f.value;
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
                })
            );
        }

        // ── Step 3: GroupBy + Aggregations ────────────────────────────
        if (queryObj.groupBy?.length) {
            const groups = new Map<string, Record<string, any>[]>();
            rows.forEach(row => {
                const key = queryObj.groupBy!.map(g => String(row[g])).join("|");
                if (!groups.has(key)) groups.set(key, []);
                groups.get(key)!.push(row);
            });

            rows = Array.from(groups.values()).map(groupRows => {
                const result: Record<string, any> = {};
                queryObj.groupBy!.forEach(g => { result[g] = groupRows[0][g]; });
                (queryObj.aggregations || []).forEach(agg => {
                    const vals = groupRows.map(r => r[agg.column]).filter(v => v !== null && v !== undefined);
                    const nums = vals.map(Number).filter(v => !isNaN(v));
                    const outKey = `${agg.op}(${agg.column})`;
                    switch (agg.op) {
                        case "sum":   result[outKey] = nums.reduce((a, b) => a + b, 0); break;
                        case "avg":   result[outKey] = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0; break;
                        case "count": result[outKey] = vals.length; break;
                        case "max":   result[outKey] = nums.length ? Math.max(...nums) : ""; break;
                        case "min":   result[outKey] = nums.length ? Math.min(...nums) : ""; break;
                        case "first": result[outKey] = vals[0] ?? ""; break;
                    }
                });
                return result;
            });
        }

        // ── Step 4: Sort ───────────────────────────────────────────────
        if (queryObj.sort) {
            const { column, direction } = queryObj.sort;
            rows.sort((a, b) => {
                const av = a[column], bv = b[column];
                const an = Number(av), bn = Number(bv);
                const cmp = !isNaN(an) && !isNaN(bn) ? an - bn : String(av).localeCompare(String(bv));
                return direction === "desc" ? -cmp : cmp;
            });
        }

        // ── Step 5: Limit ──────────────────────────────────────────────
        if (queryObj.limit && queryObj.limit > 0) {
            rows = rows.slice(0, queryObj.limit);
        }

        if (rows.length === 0) {
            return `<div class="report-section">查询结果 (共 0 条)</div>`;
        }

        // ── Step 6: Render HTML table ──────────────────────────────────
        const headers = Object.keys(rows[0]);
        const thead = `<thead><tr>${headers.map(h => `<th>${this.formatCellValue(h)}</th>`).join("")}</tr></thead>`;
        const tbody = `<tbody>${rows.map(row =>
            `<tr>${headers.map(h => `<td>${this.formatCellValue(row[h])}</td>`).join("")}</tr>`
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

    private copyToClipboard(text: string): void {
        navigator.clipboard?.writeText(text);
    }

    private isCapabilityBoundaryIssue(error: any): boolean {
        const text = String(error?.message || error || "").toLowerCase();
        return text.includes("privilege") || text.includes("webaccess") || text.includes("cors");
    }

    private saveChatHistory(): void { }
    private loadChatHistory(): void { }
    private clearChatHistory(): void { }

    private extractFilters(dataView: DataView): void {
        const filters = (dataView?.metadata as any)?.filters || [];
        this.reportContext.filters = Array.isArray(filters)
            ? filters.map((f: any) => ({ table: f?.target?.table || "", column: f?.target?.column || "", values: (f?.values || []).map(String), filterType: f?.$schema || "" }))
            : [];
    }

    private extractJsonFilters(jsonFilters: any[]): void {
        if (!Array.isArray(jsonFilters)) return;
        this.reportContext.filters.push(...jsonFilters.map((f: any) => ({ table: f?.target?.table || "", column: f?.target?.column || "", values: (f?.values || []).map(String), filterType: f?.filterType || "json" })));
    }

    private extractTableData(dataView: DataView): void {
        const table = dataView?.table;
        if (!table) return;
        this.reportContext.tableData = table.rows.map((row) => {
            const item: TableRow = {};
            table.columns.forEach((c, i) => {
                item[c.displayName] = row[i] as any;
            });
            return item;
        });
    }

    private extractMeasures(_dataView: DataView): void { }
    private extractDateRange(_dataView: DataView): void { }
}
