/* eslint-disable powerbi-visuals/no-inner-outer-html, powerbi-visuals/insecure-random */
"use strict";
import powerbi from "powerbi-visuals-api";
import { Chart, registerables } from "chart.js";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import { VisualFormattingSettingsModel } from "./settings";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

Chart.register(...registerables);

interface Message { text: string; isUser: boolean; timestamp: Date; loading?: boolean; error?: boolean; }
interface ChatSession { id: string; title: string; messages: Message[]; updatedAt: string; }
interface DataQuery {
    intent: "data_query";
    filters?: Array<{ column: string; operator: ">" | "<" | ">=" | "<=" | "==" | "!=" | "contains"; value: string | number; }>;
    groupBy?: string[];
    aggregations?: Array<{ column: string; op: "sum" | "avg" | "count" | "max" | "min" | "first"; }>;
    sort?: { column: string; direction: "asc" | "desc"; };
    limit?: number;
}

class ChatHistoryStorage {
    private static readonly MAX_CHUNKS = 50;
    private static readonly MAX_KB = 500;
    static saveHistory(host: IVisualHost, histories: ChatSession[]): void {
        const payload = JSON.stringify(histories);
        const chunks = this.chunk(payload, this.MAX_CHUNKS);
        const merge: any[] = [{ objectName: "historySettings", selector: null, properties: { chunkCount: chunks.length } }];
        for (let i = 0; i < this.MAX_CHUNKS; i++) {
            merge.push({ objectName: "historySettings", selector: null, properties: { ["chunk" + i]: chunks[i] || "" } });
        }
        host.persistProperties({ merge });
    }
    static loadHistory(metadata: any): ChatSession[] {
        try {
            const obj = metadata?.objects?.historySettings;
            const count = Number(obj?.chunkCount || 0);
            if (!count) return [];
            let text = "";
            for (let i = 0; i < count; i++) text += String(obj["chunk" + i] || "");
            return JSON.parse(text);
        } catch { return []; }
    }
    static loadTmdl(metadata: any): string {
        const obj = metadata?.objects?.tmdlSettings;
        const count = Number(obj?.chunkCount || 0);
        if (!count) return "";
        let text = "";
        for (let i = 0; i < count; i++) text += String(obj["chunk" + i] || "");
        return text;
    }
    static saveTmdl(host: IVisualHost, code: string): void {
        const chunks = this.chunk(code, 10);
        const merge: any[] = [{ objectName: "tmdlSettings", selector: null, properties: { chunkCount: chunks.length, tmdlCode: "已迁移到分块存储" } }];
        for (let i = 0; i < 10; i++) merge.push({ objectName: "tmdlSettings", selector: null, properties: { ["chunk" + i]: chunks[i] || "" } });
        host.persistProperties({ merge });
    }
    private static chunk(text: string, maxChunks: number): string[] {
        const maxLen = Math.floor((this.MAX_KB * 1024) / maxChunks);
        const chunks: string[] = [];
        for (let i = 0; i < text.length && chunks.length < maxChunks; i += maxLen) chunks.push(text.slice(i, i + maxLen));
        return chunks;
    }
}

export class Visual implements IVisual {
    private host: IVisualHost;
    private target: HTMLElement;
    private formattingSettingsService: FormattingSettingsService;
    private formattingSettings: VisualFormattingSettingsModel;
    private dataView: DataView | null = null;
    private metadata: any = {};

    private container: HTMLElement;
    private messagesContainer: HTMLElement;
    private suggestionsArea: HTMLElement;
    private inputField: HTMLTextAreaElement;
    private sendButton: HTMLButtonElement;
    private settingsModal: HTMLElement;

    private messages: Message[] = [];
    private histories: ChatSession[] = [];
    private currentSessionId: string = "";
    private tmdlCode: string = "";
    private charts: Map<string, Chart> = new Map();

    private abortController: AbortController | null = null;
    private isGenerating = false;
    private isComposing = false;
    private currentStreamingMessageIndex = -1;
    private streamingMessageElement: HTMLElement | null = null;

    private licenseKey = "";
    private activeSystemSecret = "";
    private currentDeviceId = "";
    private systemPrompt = "";
    private isDesktopEnv = false;
    private lastValidationError = "";
    private validationDebounceTimer: number | null = null;
    private isValidationInProgress = false;
    private validationPromise: Promise<void> | null = null;
    private get isLicenseValid(): boolean { return !!this.activeSystemSecret && this.activeSystemSecret.length > 5; }

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.target = options.element;
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();
        this.createUI();
        this.addStyles();
        this.currentSessionId = this.generateSessionId();
        this.loadChatHistory();
        if (this.messages.length === 0) this.addWelcomeMessage();
        this.addSuggestionChips();
    }

    public update(options: VisualUpdateOptions): void {
        this.dataView = options.dataViews?.[0] || null;
        this.metadata = this.dataView?.metadata || {};
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, this.dataView);
        this.isDesktopEnv = this.checkIsDesktop();
        this.tmdlCode = ChatHistoryStorage.loadTmdl(this.metadata);
        this.histories = ChatHistoryStorage.loadHistory(this.metadata);
        this.loadChatHistory();
        this.licenseKey = this.readObject("licenseSettings", "licenseKey", "");
        this.ensureLicenseValidated();
        this.addSuggestionChips();
    }

    private createUI(): void {
        this.container = document.createElement("div");
        this.container.className = "chat-container";
        const header = document.createElement("div");
        header.className = "chat-header";
        header.innerHTML = `<div class='chat-title'>ABI Chat</div><div class='chat-icons'><span class='icon-history'>📋</span><span class='icon-settings'>⚙</span><span class='icon-add'>+</span></div>`;
        header.querySelector(".icon-settings")!.addEventListener("click", () => this.openSettings());
        header.querySelector(".icon-add")!.addEventListener("click", () => this.startNewSession());
        header.querySelector(".icon-history")!.addEventListener("click", () => this.showHistoryModal());
        this.suggestionsArea = document.createElement("div");
        this.suggestionsArea.className = "suggestions-area";
        this.messagesContainer = document.createElement("div");
        this.messagesContainer.className = "messages-container";
        const inputContainer = document.createElement("div");
        inputContainer.className = "input-container";
        this.inputField = document.createElement("textarea");
        this.inputField.className = "input-field";
        this.inputField.placeholder = "请输入问题...";
        this.inputField.addEventListener("compositionstart", () => this.isComposing = true);
        this.inputField.addEventListener("compositionend", () => this.isComposing = false);
        this.inputField.addEventListener("input", () => this.adjustTextareaHeight());
        this.inputField.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey && !this.isComposing) { e.preventDefault(); this.sendMessage(); }
        });
        this.sendButton = document.createElement("button");
        this.sendButton.className = "send-button";
        this.sendButton.innerHTML = "<span class='icon-send'>➤</span>";
        this.sendButton.addEventListener("click", () => this.isGenerating ? this.stopGeneration() : this.sendMessage());
        inputContainer.append(this.inputField, this.sendButton);
        this.settingsModal = document.createElement("div");
        this.settingsModal.className = "tmdl-modal";
        this.settingsModal.style.display = "none";
        this.container.append(header, this.suggestionsArea, this.messagesContainer, inputContainer, this.settingsModal);
        this.target.innerHTML = "";
        this.target.appendChild(this.container);
        this.createSettingsModal();
    }

    private readObject(objectName: string, prop: string, fallback: string): string {
        return String((this.metadata?.objects as any)?.[objectName]?.[prop] ?? fallback);
    }

    private checkIsDesktop(): boolean {
        const host = (window.location.hostname || "").toLowerCase();
        return host.includes("localhost") || host.includes("desktop") || host.includes("127.0.0.1");
    }
    private getDeviceFingerprint(): string {
        const nav = window.navigator;
        const raw = [nav.userAgent, nav.language, screen.width, screen.height, Intl.DateTimeFormat().resolvedOptions().timeZone].join("|");
        let hash = 0;
        for (let i = 0; i < raw.length; i++) hash = (hash << 5) - hash + raw.charCodeAt(i);
        return "fp_" + Math.abs(hash).toString(36);
    }
    private async getOrCreateDeviceId(): Promise<string> {
        const key = "abi_device_id";
        const existing = localStorage.getItem(key);
        if (existing) return existing;
        const v = this.getDeviceFingerprint();
        localStorage.setItem(key, v);
        return v;
    }
    private async ensureLicenseValidated(): Promise<boolean> {
        if (!this.licenseKey) return false;
        if (this.validationPromise) { await this.validationPromise; return this.isLicenseValid; }
        this.validationPromise = this.validateLicense(this.licenseKey, this.isDesktopEnv ? 1 : 2).finally(() => this.validationPromise = null);
        await this.validationPromise;
        return this.isLicenseValid;
    }
    private async validateLicense(key: string, viewMode: number): Promise<void> {
        if (this.validationDebounceTimer) window.clearTimeout(this.validationDebounceTimer);
        this.validationPromise = new Promise<void>(resolve => {
            this.validationDebounceTimer = window.setTimeout(async () => { await this.performLicenseValidation(key, viewMode); resolve(); }, 120);
        });
        await this.validationPromise;
    }
    private async performLicenseValidation(key: string, _viewMode: number): Promise<void> {
        this.isValidationInProgress = true;
        try {
            const endpoint = atob("aHR0cHM6Ly9wb3dlcmJpc3Rhci5jb20vYXBpL2xpY2Vuc2UvdmFsaWRhdGU=");
            this.currentDeviceId = await this.getOrCreateDeviceId();
            const res = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ licenseKey: key, deviceId: this.currentDeviceId, environment: this.isDesktopEnv ? "desktop" : "service" }) });
            if (!res.ok) throw new Error("license request failed");
            const data = await res.json();
            this.activeSystemSecret = String(data?.system_secret || "");
            this.systemPrompt = this.decryptPrompt(String(data?.system_data || ""), this.activeSystemSecret || key);
            if (!this.activeSystemSecret) throw new Error("invalid");
        } catch (e) {
            this.handleInvalidLicense(0, String(e));
        } finally {
            this.isValidationInProgress = false;
        }
    }
    private handleInvalidLicense(_viewMode: number, error?: string): void { this.activeSystemSecret = ""; this.lastValidationError = error || "License 无效"; }
    private maskString(str: string): string { return btoa(Array.from(str).map(c => String.fromCharCode(c.charCodeAt(0) + 3)).join("")); }
    private unmaskString(str: string): string { try { return Array.from(atob(str)).map(c => String.fromCharCode(c.charCodeAt(0) - 3)).join(""); } catch { return str; } }
    private decryptPrompt(encryptedData: string, key: string): string {
        try {
            const bin = atob(encryptedData);
            let out = "";
            for (let i = 0; i < bin.length; i++) out += String.fromCharCode(bin.charCodeAt(i) ^ key.charCodeAt(i % key.length));
            return out;
        } catch { return ""; }
    }
    private getPowerBIStarSystemPrompt(): string { return this.isLicenseValid ? this.systemPrompt : "演示模式：未激活 License，仅提供基础分析。"; }

    private generateSessionId(): string { return `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`; }
    private saveLastActiveSessionId(id: string): void { this.host.persistProperties({ merge: [{ objectName: "historySettings", selector: null, properties: { lastActiveSessionId: id } }] }); }
    private async startNewSession(): Promise<void> { await this.saveCurrentChatToHistory(); this.currentSessionId = this.generateSessionId(); this.messages = []; this.messagesContainer.innerHTML = ""; this.addWelcomeMessage(); this.saveLastActiveSessionId(this.currentSessionId); }
    private async saveCurrentChatToHistory(): Promise<void> {
        const firstUser = this.messages.find(m => m.isUser)?.text || "新会话";
        const session: ChatSession = { id: this.currentSessionId, title: firstUser.slice(0, 15), messages: [...this.messages], updatedAt: new Date().toISOString() };
        this.histories = [session, ...this.histories.filter(h => h.id !== session.id)].slice(0, 50);
        await this.saveChatHistory();
    }
    private async saveChatHistory(): Promise<void> {
        try { ChatHistoryStorage.saveHistory(this.host, this.histories); localStorage.setItem("abi_histories", JSON.stringify(this.histories)); }
        catch { try { sessionStorage.setItem("abi_histories", JSON.stringify(this.histories)); } catch { } }
    }
    private async loadChatHistory(): Promise<void> {
        const pbiHistories = ChatHistoryStorage.loadHistory(this.metadata);
        if (pbiHistories.length) this.histories = pbiHistories;
        if (!this.histories.length) {
            try { this.histories = JSON.parse(localStorage.getItem("abi_histories") || "[]"); }
            catch { this.histories = JSON.parse(sessionStorage.getItem("abi_histories") || "[]"); }
        }
        const lastId = this.readObject("historySettings", "lastActiveSessionId", "");
        const active = this.histories.find(h => h.id === lastId) || this.histories[0];
        if (active) { this.currentSessionId = active.id; this.messages = active.messages.map(m => ({ ...m, timestamp: new Date(m.timestamp) })); this.renderAllMessages(); }
    }
    private async clearChatHistory(): Promise<void> { this.histories = []; await this.saveChatHistory(); localStorage.removeItem("abi_histories"); sessionStorage.removeItem("abi_histories"); }

    private showHistoryModal(): void {
        const modal = document.createElement("div");
        modal.className = "history-modal tmdl-modal";
        modal.innerHTML = `<div class='tmdl-modal-content'><div class='tmdl-modal-header'><h3>会话历史</h3><button class='close-btn'>×</button></div><div class='tmdl-modal-body'></div></div>`;
        const body = modal.querySelector(".tmdl-modal-body") as HTMLElement;
        this.histories.forEach(h => {
            const item = document.createElement("div");
            item.className = "history-item";
            item.innerHTML = `<span>${this.sanitizeHTML(h.title)} (${new Date(h.updatedAt).toLocaleString()})</span><div><button class='rename'>重命名</button><button class='load'>加载</button><button class='delete'>删除</button></div>`;
            item.querySelector(".load")!.addEventListener("click", () => { this.messages = h.messages.map(m => ({ ...m, timestamp: new Date(m.timestamp) })); this.currentSessionId = h.id; this.messagesContainer.innerHTML = ""; this.renderAllMessages(); this.saveLastActiveSessionId(h.id); modal.remove(); });
            item.querySelector(".delete")!.addEventListener("click", async () => { this.histories = this.histories.filter(x => x.id !== h.id); await this.saveChatHistory(); item.remove(); });
            item.querySelector(".rename")!.addEventListener("click", async () => { const v = prompt("输入新标题", h.title); if (v) { h.title = v; await this.saveChatHistory(); item.querySelector("span")!.textContent = `${v} (${new Date(h.updatedAt).toLocaleString()})`; } });
            body.appendChild(item);
        });
        modal.querySelector(".close-btn")!.addEventListener("click", () => modal.remove());
        this.target.appendChild(modal);
    }

    private adjustTextareaHeight(): void { this.inputField.style.height = "auto"; this.inputField.style.height = Math.min(this.inputField.scrollHeight, 200) + "px"; }
    private stopGeneration(): void { this.abortController?.abort(); this.isGenerating = false; this.updateSendButtonState(); }
    private updateSendButtonState(): void {
        this.sendButton.classList.toggle("generating", this.isGenerating);
        this.sendButton.innerHTML = this.isGenerating ? "<span class='icon-stop'>■</span>" : "<span class='icon-send'>➤</span>";
    }

    private async copyToClipboard(html: string): Promise<void> {
        try {
            await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }), "text/plain": new Blob([html.replace(/<[^>]+>/g, "")], { type: "text/plain" }) })]);
        } catch { this.fallbackCopyToClipboard(html); }
    }
    private fallbackCopyToClipboard(html: string): void {
        const div = document.createElement("div");
        div.innerHTML = html;
        document.body.appendChild(div);
        const r = document.createRange(); r.selectNodeContents(div);
        const s = window.getSelection(); s?.removeAllRanges(); s?.addRange(r);
        document.execCommand("copy"); s?.removeAllRanges(); div.remove();
    }

    private startStreamingMessage(): void {
        const msg: Message = { text: "", isUser: false, timestamp: new Date(), loading: true };
        this.messages.push(msg);
        this.currentStreamingMessageIndex = this.messages.length - 1;
        this.renderMessage(msg);
        this.streamingMessageElement = this.messagesContainer.lastElementChild?.querySelector(".message-bubble") as HTMLElement;
    }
    private updateStreamingMessage(token: string): void {
        if (this.currentStreamingMessageIndex < 0 || !this.streamingMessageElement) return;
        const m = this.messages[this.currentStreamingMessageIndex];
        m.text += token;
        const thinking = /```json[\s\S]*$/m.test(m.text) && !/```json[\s\S]*```/m.test(m.text) ? "<div class='thinking-dots'>⏳ 正在思考并查询数据...</div>" : "";
        this.streamingMessageElement.innerHTML = this.formatContentForDisplay(this.removeJsonCodeBlocks(m.text)) + thinking;
        this.messagesContainer.scrollTop = this.messagesContainer.scrollHeight;
    }
    private finalizeStreamingMessage(fullText: string): void {
        if (this.currentStreamingMessageIndex < 0) return;
        const m = this.messages[this.currentStreamingMessageIndex];
        m.loading = false;
        m.text = fullText;
        if (this.streamingMessageElement) this.streamingMessageElement.innerHTML = this.formatToReportStyle(fullText);
        this.currentStreamingMessageIndex = -1;
        this.streamingMessageElement = null;
    }
    private removeStreamingMessage(): void {
        if (this.currentStreamingMessageIndex < 0) return;
        this.messages.splice(this.currentStreamingMessageIndex, 1);
        this.messagesContainer.lastElementChild?.remove();
        this.currentStreamingMessageIndex = -1;
        this.streamingMessageElement = null;
    }

    private addWelcomeMessage(): void {
        const text = this.formattingSettings.welcomeSettingsCard.welcomeMessage.value || "你好！我是 PowerBI星球 助手。";
        this.messages.push({ text, isUser: false, timestamp: new Date() });
        this.renderMessage(this.messages[this.messages.length - 1]);
    }
    private addSuggestionChips(): void {
        this.suggestionsArea.innerHTML = "<div class='suggestions-container'></div>";
        const box = this.suggestionsArea.querySelector(".suggestions-container") as HTMLElement;
        const qs = [this.formattingSettings.suggestionSettingsCard.question1.value, this.formattingSettings.suggestionSettingsCard.question2.value, this.formattingSettings.suggestionSettingsCard.question3.value].filter(Boolean);
        qs.forEach(q => {
            const btn = document.createElement("button"); btn.className = "suggestion-btn"; btn.textContent = q; btn.addEventListener("click", () => { this.inputField.value = q; this.sendMessage(); }); box.appendChild(btn);
        });
    }

    private async sendMessage(): Promise<void> {
        const text = this.inputField.value.trim();
        if (!text) return;
        this.messages.push({ text, isUser: true, timestamp: new Date() });
        this.renderMessage(this.messages[this.messages.length - 1]);
        this.inputField.value = ""; this.adjustTextareaHeight();
        this.abortController = new AbortController();
        this.isGenerating = true; this.updateSendButtonState();
        this.startStreamingMessage();
        try {
            const full = await this.callAIAPIWithStreaming(text, (t) => this.updateStreamingMessage(t));
            this.finalizeStreamingMessage(full);
            await this.saveCurrentChatToHistory();
        } catch (e) {
            this.removeStreamingMessage();
            this.messages.push({ text: "请求失败: " + String(e), isUser: false, timestamp: new Date(), error: true });
            this.renderMessage(this.messages[this.messages.length - 1]);
        } finally {
            this.isGenerating = false; this.updateSendButtonState();
        }
    }

    private prepareDataContext(): string {
        const rows = this.dataView?.table?.rows || [];
        const columns = this.dataView?.table?.columns || [];
        const summary: string[] = [`数据概览: ${columns.length}列 ${rows.length}行`];
        summary.push("列名: " + columns.map(c => c.displayName).join(", "));
        columns.forEach((c, idx) => {
            const nums = rows.map(r => r[idx]).filter(v => typeof v === "number") as number[];
            if (!nums.length) return;
            const isRatio = (c.format || "").includes("%");
            summary.push(`${c.displayName}: ${isRatio ? "比率字段不可汇总" : "总计=" + nums.reduce((a, b) => a + b, 0).toLocaleString()}`);
        });
        summary.push("完整数据:");
        rows.forEach((r, i) => summary.push(`${i + 1}. ${columns.map((c, idx) => `${c.displayName}: ${this.formatCellValue(r[idx])}`).join(", ")}`));
        if (this.tmdlCode) summary.unshift("TMDL:\n" + this.tmdlCode);
        return summary.join("\n");
    }

    private executeDataQuery(query: DataQuery): { rows: any[]; columns: string[] } {
        try {
            const columns = this.dataView?.table?.columns || [];
            let rows = (this.dataView?.table?.rows || []).map(r => r.slice());
            const colIndex = (n: string) => columns.findIndex(c => c.displayName === n || c.queryName === n);
            if (query.filters?.length) rows = rows.filter(r => query.filters!.every(f => {
                const v = r[colIndex(f.column)]; const n = Number(v); const fn = Number(f.value);
                switch (f.operator) { case ">": return n > fn; case "<": return n < fn; case ">=": return n >= fn; case "<=": return n <= fn; case "==": return String(v) === String(f.value); case "!=": return String(v) !== String(f.value); default: return String(v).includes(String(f.value)); }
            }));
            if (query.groupBy?.length && query.aggregations?.length) {
                const map = new Map<string, any[]>();
                rows.forEach(r => { const k = query.groupBy!.map(g => r[colIndex(g)]).join("|"); if (!map.has(k)) map.set(k, []); map.get(k)!.push(r); });
                const out: any[] = [];
                map.forEach((rs, k) => {
                    const o: any = {}; query.groupBy!.forEach((g, i) => o[g] = k.split("|")[i]);
                    query.aggregations!.forEach(a => {
                        const arr = rs.map(x => Number(x[colIndex(a.column)])).filter(x => !Number.isNaN(x));
                        let v = 0;
                        if (a.op === "sum") v = arr.reduce((p, c) => p + c, 0);
                        if (a.op === "avg") v = arr.reduce((p, c) => p + c, 0) / (arr.length || 1);
                        if (a.op === "count") v = arr.length;
                        if (a.op === "max") v = Math.max(...arr);
                        if (a.op === "min") v = Math.min(...arr);
                        if (a.op === "first") v = arr[0];
                        o[`${a.column} (${a.op})`] = Math.round(100 * v) / 100;
                    });
                    out.push(o);
                });
                return { rows: out, columns: Object.keys(out[0] || {}) };
            }
            const out = rows.map(r => Object.fromEntries(columns.map((c, i) => [c.displayName, r[i]])));
            return { rows: out, columns: columns.map(c => c.displayName) };
        } catch { return { rows: [], columns: [] }; }
    }

    private formatToReportStyle(text: string): string {
        const regex = /```json\s*([\s\S]*?)\s*```/g;
        return this.sanitizeHTML(text).replace(regex, (_m, json) => {
            try {
                const q: DataQuery = JSON.parse(json);
                if (q.intent !== "data_query") return "";
                const result = this.executeDataQuery(q);
                return `<div class='report-table-container'><table class='report-table'><thead><tr>${result.columns.map(c => `<th>${this.sanitizeHTML(c)}</th>`).join("")}</tr></thead><tbody>${result.rows.map(r => `<tr>${result.columns.map(c => `<td>${this.formatCellValue((r as any)[c])}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
            } catch { return ""; }
        });
    }
    private formatContentForDisplay(text: string): string { return this.sanitizeHTML(text.replace(/```html|```/g, "")).replace(/\n/g, "<br>"); }
    private sanitizeHTML(html: string): string { const d = document.createElement("div"); d.textContent = html; return d.innerHTML; }
    private formatCellValue(value: any): string { return typeof value === "number" ? value.toLocaleString("zh-CN") : this.sanitizeHTML(String(value ?? "")); }
    private isCapabilityBoundaryIssue(text: string): boolean { return ["不清楚", "不确定", "无法确定", "不知道", "超出", "范围", "复杂", "具体", "详细", "个性化", "定制", "特殊需求"].some(k => text.includes(k)); }

    private async callAIAPIWithStreaming(userMessage: string, onChunk: (token: string) => void): Promise<string> {
        const apiUrl = this.formattingSettings.aiSettingsCard.apiUrl.value || "https://api.openai.com/v1/chat/completions";
        const apiKey = this.unmaskString(this.formattingSettings.aiSettingsCard.apiKey.value || "");
        const model = this.formattingSettings.aiSettingsCard.model.value || "gpt-4o-mini";
        const isGemini = apiUrl.includes("googleapis.com") || model.toLowerCase().startsWith("gemini");
        if (isGemini) {
            const text = await this.callAIAPI(userMessage);
            for (const c of text) { onChunk(c); await new Promise(r => setTimeout(r, 10)); }
            return text;
        }
        const body = {
            model,
            stream: true,
            messages: [
                { role: "system", content: this.getPowerBIStarSystemPrompt() },
                { role: "system", content: this.prepareDataContext() },
                ...this.getConversationHistoryMessages(),
                { role: "user", content: userMessage }
            ]
        };
        const res = await fetch(apiUrl, { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` }, body: JSON.stringify(body), signal: this.abortController?.signal });
        if (!res.body) throw new Error("no stream");
        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let full = "";
        while (true) {
            const { value, done } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split("\n").filter(l => l.startsWith("data:"));
            for (const line of lines) {
                const txt = line.replace(/^data:\s*/, "").trim();
                if (!txt || txt === "[DONE]") continue;
                try {
                    const json = JSON.parse(txt);
                    const token = json.choices?.[0]?.delta?.content || "";
                    if (token) { full += token; onChunk(token); }
                } catch { }
            }
        }
        return full;
    }
    private async callAIAPI(userMessage: string): Promise<string> {
        const apiUrl = this.formattingSettings.aiSettingsCard.apiUrl.value;
        const apiKey = this.unmaskString(this.formattingSettings.aiSettingsCard.apiKey.value);
        const model = this.formattingSettings.aiSettingsCard.model.value;
        const payload = { model, messages: [{ role: "system", content: this.getPowerBIStarSystemPrompt() }, { role: "system", content: this.prepareDataContext() }, ...this.getConversationHistoryMessages(), { role: "user", content: userMessage }] };
        const res = await fetch(apiUrl, { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` }, body: JSON.stringify(payload), signal: this.abortController?.signal });
        const json = await res.json();
        return json?.choices?.[0]?.message?.content || json?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    }
    private getConversationHistoryMessages(): Array<{ role: string; content: string; }> {
        const core = this.messages.filter(m => !m.loading && !m.error).slice(0, -1).slice(-10);
        return core.map(m => ({ role: m.isUser ? "user" : "assistant", content: m.text }));
    }

    private shouldGenerateChart(text: string): boolean { return ["图表", "柱状", "折线", "饼图", "line", "bar", "pie"].some(k => text.toLowerCase().includes(k)); }
    private generateChartFromData(text: string): any | null {
        const rows = this.dataView?.table?.rows || [];
        const cols = this.dataView?.table?.columns || [];
        if (cols.length < 2 || rows.length === 0) return null;
        const labels = rows.slice(0, 20).map(r => String(r[0]));
        const data = rows.slice(0, 20).map(r => Number(r[1]) || 0);
        const type = text.includes("饼") ? "pie" : (text.includes("折") ? "line" : "bar");
        return { type, labels, datasets: [{ label: cols[1].displayName, data }] };
    }
    private renderChart(canvas: HTMLCanvasElement, chartData: any): void {
        const id = canvas.dataset.chartId || this.generateSessionId();
        canvas.dataset.chartId = id;
        this.charts.get(id)?.destroy();
        const chart = new Chart(canvas, { type: chartData.type, data: { labels: chartData.labels, datasets: chartData.datasets } as any, options: { responsive: true, maintainAspectRatio: false } });
        this.charts.set(id, chart);
    }
    private parseAIResponse(text: string): { text: string; chartData?: any; } {
        if (!this.shouldGenerateChart(text)) return { text };
        const chartData = this.generateChartFromData(text);
        return chartData ? { text, chartData } : { text };
    }

    private renderMessage(message: Message): void {
        const row = document.createElement("div"); row.className = `message ${message.isUser ? "user" : "assistant"}`;
        const bubble = document.createElement("div"); bubble.className = "message-bubble";
        const parsed = this.parseAIResponse(message.text);
        bubble.innerHTML = message.isUser ? this.formatContentForDisplay(parsed.text) : this.formatToReportStyle(parsed.text);
        if (!message.isUser) {
            const copyBtn = document.createElement("button"); copyBtn.className = "copy-button"; copyBtn.textContent = "复制"; copyBtn.addEventListener("click", () => this.copyToClipboard(bubble.innerHTML));
            row.append(bubble, copyBtn);
            if (parsed.chartData) {
                const wrap = document.createElement("div"); wrap.className = "chart-container";
                const c = document.createElement("canvas"); c.style.height = "220px"; wrap.appendChild(c); row.appendChild(wrap); setTimeout(() => this.renderChart(c, parsed.chartData), 0);
            }
        } else row.appendChild(bubble);
        this.messagesContainer.appendChild(row);
        this.messagesContainer.scrollTop = this.messagesContainer.scrollHeight;
    }
    private renderAllMessages(): void { this.messagesContainer.innerHTML = ""; this.messages.forEach(m => this.renderMessage(m)); }
    private removeJsonCodeBlocks(text: string): string { return text.replace(/```json[\s\S]*?```/g, "").trim(); }

    private createSettingsModal(): void {
        this.settingsModal.innerHTML = `<div class='tmdl-modal-content'><div class='tmdl-modal-header'><h3>设置</h3><button class='close-btn'>×</button></div><div class='tmdl-modal-body'><div class='tabs'><button class='tab-btn active' data-tab='ai'>AI 连接配置</button><button class='tab-btn' data-tab='license'>License 配置</button><button class='tab-btn' data-tab='tmdl'>TMDL 管理</button></div><div class='tab-content' data-tab='ai'><input class='form-control' id='apiUrl' placeholder='API URL'/><input class='form-control' id='apiKey' placeholder='API Key'/><input class='form-control' id='model' placeholder='Model'/></div><div class='tab-content' data-tab='license' style='display:none'><input class='form-control' id='licenseKey' placeholder='License Key'/><button id='validateLicense'>验证</button><div id='licenseResult'></div></div><div class='tab-content' data-tab='tmdl' style='display:none'><div class='tmdl-upload-section'><input type='file' id='tmdlFile'/></div><textarea class='form-control' id='tmdlEditor' rows='10'></textarea><button id='saveTmdl'>保存 TMDL</button></div></div></div>`;
        this.settingsModal.querySelectorAll(".tab-btn").forEach(btn => btn.addEventListener("click", () => {
            const tab = (btn as HTMLElement).dataset.tab!;
            this.settingsModal.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active")); btn.classList.add("active");
            this.settingsModal.querySelectorAll(".tab-content").forEach(c => (c as HTMLElement).style.display = (c as HTMLElement).dataset.tab === tab ? "block" : "none");
        }));
        this.settingsModal.querySelector(".close-btn")!.addEventListener("click", () => this.settingsModal.style.display = "none");
        (this.settingsModal.querySelector("#saveTmdl") as HTMLButtonElement).addEventListener("click", () => {
            const text = (this.settingsModal.querySelector("#tmdlEditor") as HTMLTextAreaElement).value;
            ChatHistoryStorage.saveTmdl(this.host, text); this.tmdlCode = text;
        });
        (this.settingsModal.querySelector("#tmdlFile") as HTMLInputElement).addEventListener("change", (e) => {
            const f = (e.target as HTMLInputElement).files?.[0]; if (!f) return;
            f.text().then(t => (this.settingsModal.querySelector("#tmdlEditor") as HTMLTextAreaElement).value = t);
        });
        (this.settingsModal.querySelector("#validateLicense") as HTMLButtonElement).addEventListener("click", async () => {
            this.licenseKey = (this.settingsModal.querySelector("#licenseKey") as HTMLInputElement).value;
            await this.validateLicense(this.licenseKey, 0);
            (this.settingsModal.querySelector("#licenseResult") as HTMLElement).textContent = this.isLicenseValid ? "验证成功" : `验证失败: ${this.lastValidationError}`;
        });
    }
    private openSettings(): void {
        this.settingsModal.style.display = "flex";
        (this.settingsModal.querySelector("#apiUrl") as HTMLInputElement).value = this.formattingSettings.aiSettingsCard.apiUrl.value;
        (this.settingsModal.querySelector("#apiKey") as HTMLInputElement).value = this.unmaskString(this.formattingSettings.aiSettingsCard.apiKey.value);
        (this.settingsModal.querySelector("#model") as HTMLInputElement).value = this.formattingSettings.aiSettingsCard.model.value;
        (this.settingsModal.querySelector("#licenseKey") as HTMLInputElement).value = this.licenseKey;
        (this.settingsModal.querySelector("#tmdlEditor") as HTMLTextAreaElement).value = this.tmdlCode;
    }

    private addStyles(): void {
        const style = document.createElement("style");
        style.textContent = `.chat-container{height:100%;display:flex;flex-direction:column;background:#f5f7fa}.chat-header{display:flex;justify-content:space-between;padding:10px 12px;background:#5b6ef5;color:#fff}.chat-icons span{cursor:pointer;margin-left:8px}.messages-container{flex:1;overflow:auto;padding:12px}.message{margin-bottom:8px}.message.user{text-align:right}.message-bubble{display:inline-block;background:#fff;padding:8px 10px;border-radius:8px;max-width:90%}.message.user .message-bubble{background:#dbe7ff}.input-container{display:flex;gap:8px;padding:10px}.input-field{flex:1;resize:none;max-height:200px}.send-button{width:44px}.generating{background:#f59e0b}.suggestions-container{display:flex;gap:8px;padding:8px;flex-wrap:wrap}.suggestion-btn{border:1px solid #ddd;border-radius:14px;padding:4px 10px}.copy-button{margin-left:8px}.chart-container{height:240px;margin-top:8px}.report-table{width:100%;border-collapse:collapse}.report-table th,.report-table td{border:1px solid #ddd;padding:4px}.tmdl-modal{position:absolute;inset:0;background:#0006;display:flex;align-items:center;justify-content:center}.tmdl-modal-content{background:#fff;width:720px;max-width:92%;border-radius:8px}.tmdl-modal-header{display:flex;justify-content:space-between;padding:10px;border-bottom:1px solid #eee}.tmdl-modal-body{padding:12px}.tab-btn{padding:6px 10px}.tab-btn.active{background:#5b6ef5;color:#fff}.form-control{width:100%;margin:6px 0;padding:7px}.thinking-dots{font-size:12px;color:#666}.history-item{display:flex;justify-content:space-between;border-bottom:1px solid #eee;padding:8px 0}`;
        document.head.appendChild(style);
    }

    public destroy(): void { this.charts.forEach(c => c.destroy()); }
}
