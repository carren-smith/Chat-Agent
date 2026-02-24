"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

class AiSettingsCard extends FormattingSettingsCard {
    apiKey = new formattingSettings.TextInput({ name: "apiKey", displayName: "API Key", value: "", placeholder: "sk-..." });
    apiUrl = new formattingSettings.TextInput({ name: "apiUrl", displayName: "API URL", value: "https://api.openai.com/v1/chat/completions", placeholder: "" });
    model = new formattingSettings.TextInput({ name: "model", displayName: "Model", value: "gpt-4o-mini", placeholder: "" });
    name = "aiSettings";
    displayName = "AI 连接配置";
    slices: Array<FormattingSettingsSlice> = [this.apiKey, this.apiUrl, this.model];
}

class WelcomeSettingsCard extends FormattingSettingsCard {
    welcomeMessage = new formattingSettings.TextArea({ name: "welcomeMessage", displayName: "欢迎语", value: "你好！我是 PowerBI 星球助手。", placeholder: "" });
    name = "welcomeSettings";
    displayName = "欢迎语";
    slices: Array<FormattingSettingsSlice> = [this.welcomeMessage];
}

class SuggestionSettingsCard extends FormattingSettingsCard {
    question1 = new formattingSettings.TextInput({ name: "question1", displayName: "建议问题 1", value: "请总结当前报表关键结论", placeholder: "" });
    question2 = new formattingSettings.TextInput({ name: "question2", displayName: "建议问题 2", value: "按维度做排名分析", placeholder: "" });
    question3 = new formattingSettings.TextInput({ name: "question3", displayName: "建议问题 3", value: "有没有异常值或异常趋势", placeholder: "" });
    name = "suggestionSettings";
    displayName = "建议问题";
    slices: Array<FormattingSettingsSlice> = [this.question1, this.question2, this.question3];
}

class AboutCardSettings extends FormattingSettingsCard {
    author = new formattingSettings.TextInput({ name: "author", displayName: "作者", value: "powerai001", placeholder: "" });
    version = new formattingSettings.TextInput({ name: "version", displayName: "版本", value: "2.0.0", placeholder: "" });
    declaration = new formattingSettings.TextArea({ name: "declaration", displayName: "声明", value: "本视觉仅用于分析辅助，结果请人工复核。", placeholder: "" });
    name = "aboutSettings";
    displayName = "关于";
    slices: Array<FormattingSettingsSlice> = [this.author, this.version, this.declaration];
}

class TmdlSettingsCard extends FormattingSettingsCard {
    tmdlCode = new formattingSettings.TextArea({ name: "tmdlCode", displayName: "TMDL（已废弃）", value: "", placeholder: "" });
    lastActiveSessionId = new formattingSettings.TextInput({ name: "lastActiveSessionId", displayName: "lastActiveSessionId", value: "", visible: false, placeholder: "" });
    chunkCount = new formattingSettings.NumUpDown({ name: "chunkCount", displayName: "chunkCount", value: 0, visible: false });
    chunk0 = new formattingSettings.TextInput({ name: "chunk0", displayName: "chunk0", value: "", visible: false, placeholder: "" });
    chunk1 = new formattingSettings.TextInput({ name: "chunk1", displayName: "chunk1", value: "", visible: false, placeholder: "" });
    chunk2 = new formattingSettings.TextInput({ name: "chunk2", displayName: "chunk2", value: "", visible: false, placeholder: "" });
    chunk3 = new formattingSettings.TextInput({ name: "chunk3", displayName: "chunk3", value: "", visible: false, placeholder: "" });
    chunk4 = new formattingSettings.TextInput({ name: "chunk4", displayName: "chunk4", value: "", visible: false, placeholder: "" });
    chunk5 = new formattingSettings.TextInput({ name: "chunk5", displayName: "chunk5", value: "", visible: false, placeholder: "" });
    chunk6 = new formattingSettings.TextInput({ name: "chunk6", displayName: "chunk6", value: "", visible: false, placeholder: "" });
    chunk7 = new formattingSettings.TextInput({ name: "chunk7", displayName: "chunk7", value: "", visible: false, placeholder: "" });
    chunk8 = new formattingSettings.TextInput({ name: "chunk8", displayName: "chunk8", value: "", visible: false, placeholder: "" });
    chunk9 = new formattingSettings.TextInput({ name: "chunk9", displayName: "chunk9", value: "", visible: false, placeholder: "" });
    name = "tmdlSettings";
    displayName = "TMDL";
    slices: Array<FormattingSettingsSlice> = [
        this.tmdlCode, this.lastActiveSessionId, this.chunkCount,
        this.chunk0, this.chunk1, this.chunk2, this.chunk3, this.chunk4,
        this.chunk5, this.chunk6, this.chunk7, this.chunk8, this.chunk9
    ];
}

class LicenseSettingsCard extends FormattingSettingsCard {
    licenseKey = new formattingSettings.TextInput({ name: "licenseKey", displayName: "License Key", value: "", placeholder: "" });
    boundFingerprint = new formattingSettings.TextInput({ name: "boundFingerprint", displayName: "Bound Fingerprint", value: "", visible: false, placeholder: "" });
    name = "licenseSettings";
    displayName = "License";
    slices: Array<FormattingSettingsSlice> = [this.licenseKey, this.boundFingerprint];
}

class HistorySettingsCard extends FormattingSettingsCard {
    lastActiveSessionId = new formattingSettings.TextInput({ name: "lastActiveSessionId", displayName: "lastActiveSessionId", value: "", visible: false, placeholder: "" });
    chunkCount = new formattingSettings.NumUpDown({ name: "chunkCount", displayName: "chunkCount", value: 0, visible: false });
    chunk0 = new formattingSettings.TextInput({ name: "chunk0", displayName: "chunk0", value: "", visible: false, placeholder: "" });
    name = "historySettings";
    displayName = "History";
    slices: Array<FormattingSettingsSlice> = [this.lastActiveSessionId, this.chunkCount, this.chunk0];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    aiSettingsCard = new AiSettingsCard();
    welcomeSettingsCard = new WelcomeSettingsCard();
    suggestionSettingsCard = new SuggestionSettingsCard();
    aboutSettingsCard = new AboutCardSettings();
    tmdlSettingsCard = new TmdlSettingsCard();
    licenseSettingsCard = new LicenseSettingsCard();
    historySettingsCard = new HistorySettingsCard();

    cards = [
        this.aiSettingsCard,
        this.welcomeSettingsCard,
        this.suggestionSettingsCard,
        this.aboutSettingsCard,
        this.tmdlSettingsCard,
        this.licenseSettingsCard,
        this.historySettingsCard
    ];
}
