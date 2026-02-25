"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * AI Settings Card — objectName: "aiSettings"
 */
class AiSettingsCard extends FormattingSettingsCard {
    apiKey = new formattingSettings.TextInput({
        name: "apiKey",
        displayName: "API Key",
        placeholder: "请输入 API Key",
        value: ""
    });

    apiUrl = new formattingSettings.TextInput({
        name: "apiUrl",
        displayName: "API URL",
        placeholder: "https://api.openai.com/v1/chat/completions",
        value: ""
    });

    model = new formattingSettings.TextInput({
        name: "model",
        displayName: "模型名称",
        placeholder: "gpt-4o",
        value: ""
    });

    name: string = "aiSettings";
    displayName: string = "AI 连接配置";
    slices: Array<FormattingSettingsSlice> = [this.apiKey, this.apiUrl, this.model];
}

/**
 * Welcome Settings Card — objectName: "welcomeSettings"
 */
class WelcomeSettingsCard extends FormattingSettingsCard {
    welcomeMessage = new formattingSettings.TextInput({
        name: "welcomeMessage",
        displayName: "欢迎语",
        placeholder: "你好！我是 Chat Pro...",
        value: ""
    });

    name: string = "welcomeSettings";
    displayName: string = "欢迎语设置";
    slices: Array<FormattingSettingsSlice> = [this.welcomeMessage];
}

/**
 * Suggestion Settings Card — objectName: "suggestionSettings"
 */
class SuggestionSettingsCard extends FormattingSettingsCard {
    question1 = new formattingSettings.TextInput({
        name: "question1",
        displayName: "建议问题 1",
        placeholder: "当前页面数据概览",
        value: ""
    });

    question2 = new formattingSettings.TextInput({
        name: "question2",
        displayName: "建议问题 2",
        placeholder: "帮我分析当前数据",
        value: ""
    });

    question3 = new formattingSettings.TextInput({
        name: "question3",
        displayName: "建议问题 3",
        placeholder: "有哪些异常数据？",
        value: ""
    });

    name: string = "suggestionSettings";
    displayName: string = "建议问题";
    slices: Array<FormattingSettingsSlice> = [this.question1, this.question2, this.question3];
}

/**
 * About Card Settings — objectName: "aboutSettings"
 */
class AboutCardSettings extends FormattingSettingsCard {
    author = new formattingSettings.TextInput({
        name: "author",
        displayName: "作者",
        placeholder: "",
        value: "PowerBI星球"
    });

    version = new formattingSettings.TextInput({
        name: "version",
        displayName: "版本",
        placeholder: "",
        value: "2.0.0"
    });

    declaration = new formattingSettings.TextInput({
        name: "declaration",
        displayName: "声明",
        placeholder: "",
        value: ""
    });

    name: string = "aboutSettings";
    displayName: string = "关于";
    slices: Array<FormattingSettingsSlice> = [this.author, this.version, this.declaration];
}

/**
 * TMDL Settings Card — objectName: "tmdlSettings"
 * Uses chunk storage for large TMDL code
 */
class TmdlSettingsCard extends FormattingSettingsCard {
    tmdlCode = new formattingSettings.TextInput({
        name: "tmdlCode",
        displayName: "TMDL 代码（已废弃，使用分块存储）",
        placeholder: "",
        value: ""
    });

    lastActiveSessionId = new formattingSettings.TextInput({
        name: "lastActiveSessionId",
        displayName: "Last Active Session ID",
        placeholder: "",
        value: ""
    });

    chunkCount = new formattingSettings.NumUpDown({
        name: "chunkCount",
        displayName: "Chunk Count",
        value: 0
    });

    chunk0 = new formattingSettings.TextInput({ name: "chunk0", displayName: "chunk0", placeholder: "", value: "" });
    chunk1 = new formattingSettings.TextInput({ name: "chunk1", displayName: "chunk1", placeholder: "", value: "" });
    chunk2 = new formattingSettings.TextInput({ name: "chunk2", displayName: "chunk2", placeholder: "", value: "" });
    chunk3 = new formattingSettings.TextInput({ name: "chunk3", displayName: "chunk3", placeholder: "", value: "" });
    chunk4 = new formattingSettings.TextInput({ name: "chunk4", displayName: "chunk4", placeholder: "", value: "" });
    chunk5 = new formattingSettings.TextInput({ name: "chunk5", displayName: "chunk5", placeholder: "", value: "" });
    chunk6 = new formattingSettings.TextInput({ name: "chunk6", displayName: "chunk6", placeholder: "", value: "" });
    chunk7 = new formattingSettings.TextInput({ name: "chunk7", displayName: "chunk7", placeholder: "", value: "" });
    chunk8 = new formattingSettings.TextInput({ name: "chunk8", displayName: "chunk8", placeholder: "", value: "" });
    chunk9 = new formattingSettings.TextInput({ name: "chunk9", displayName: "chunk9", placeholder: "", value: "" });

    name: string = "tmdlSettings";
    displayName: string = "TMDL 管理";
    slices: Array<FormattingSettingsSlice> = [
        this.tmdlCode, this.lastActiveSessionId, this.chunkCount,
        this.chunk0, this.chunk1, this.chunk2, this.chunk3, this.chunk4,
        this.chunk5, this.chunk6, this.chunk7, this.chunk8, this.chunk9
    ];
}

/**
 * License Settings Card — objectName: "licenseSettings"
 * Written via persistProperties
 */
class LicenseSettingsCard extends FormattingSettingsCard {
    licenseKey = new formattingSettings.TextInput({
        name: "licenseKey",
        displayName: "License Key",
        placeholder: "",
        value: ""
    });

    boundFingerprint = new formattingSettings.TextInput({
        name: "boundFingerprint",
        displayName: "Bound Fingerprint",
        placeholder: "",
        value: ""
    });

    name: string = "licenseSettings";
    displayName: string = "License";
    slices: Array<FormattingSettingsSlice> = [this.licenseKey, this.boundFingerprint];
}

/**
 * History Settings Card — objectName: "historySettings"
 * Uses chunk storage for chat history via persistProperties
 */
class HistorySettingsCard extends FormattingSettingsCard {
    lastActiveSessionId = new formattingSettings.TextInput({
        name: "lastActiveSessionId",
        displayName: "Last Active Session ID",
        placeholder: "",
        value: ""
    });

    chunkCount = new formattingSettings.NumUpDown({
        name: "chunkCount",
        displayName: "Chunk Count",
        value: 0
    });

    chunk0 = new formattingSettings.TextInput({ name: "chunk0", displayName: "chunk0", placeholder: "", value: "" });
    chunk1 = new formattingSettings.TextInput({ name: "chunk1", displayName: "chunk1", placeholder: "", value: "" });
    chunk2 = new formattingSettings.TextInput({ name: "chunk2", displayName: "chunk2", placeholder: "", value: "" });
    chunk3 = new formattingSettings.TextInput({ name: "chunk3", displayName: "chunk3", placeholder: "", value: "" });
    chunk4 = new formattingSettings.TextInput({ name: "chunk4", displayName: "chunk4", placeholder: "", value: "" });
    chunk5 = new formattingSettings.TextInput({ name: "chunk5", displayName: "chunk5", placeholder: "", value: "" });
    chunk6 = new formattingSettings.TextInput({ name: "chunk6", displayName: "chunk6", placeholder: "", value: "" });
    chunk7 = new formattingSettings.TextInput({ name: "chunk7", displayName: "chunk7", placeholder: "", value: "" });
    chunk8 = new formattingSettings.TextInput({ name: "chunk8", displayName: "chunk8", placeholder: "", value: "" });
    chunk9 = new formattingSettings.TextInput({ name: "chunk9", displayName: "chunk9", placeholder: "", value: "" });
    chunk10 = new formattingSettings.TextInput({ name: "chunk10", displayName: "chunk10", placeholder: "", value: "" });
    chunk11 = new formattingSettings.TextInput({ name: "chunk11", displayName: "chunk11", placeholder: "", value: "" });
    chunk12 = new formattingSettings.TextInput({ name: "chunk12", displayName: "chunk12", placeholder: "", value: "" });
    chunk13 = new formattingSettings.TextInput({ name: "chunk13", displayName: "chunk13", placeholder: "", value: "" });
    chunk14 = new formattingSettings.TextInput({ name: "chunk14", displayName: "chunk14", placeholder: "", value: "" });
    chunk15 = new formattingSettings.TextInput({ name: "chunk15", displayName: "chunk15", placeholder: "", value: "" });
    chunk16 = new formattingSettings.TextInput({ name: "chunk16", displayName: "chunk16", placeholder: "", value: "" });
    chunk17 = new formattingSettings.TextInput({ name: "chunk17", displayName: "chunk17", placeholder: "", value: "" });
    chunk18 = new formattingSettings.TextInput({ name: "chunk18", displayName: "chunk18", placeholder: "", value: "" });
    chunk19 = new formattingSettings.TextInput({ name: "chunk19", displayName: "chunk19", placeholder: "", value: "" });
    chunk20 = new formattingSettings.TextInput({ name: "chunk20", displayName: "chunk20", placeholder: "", value: "" });
    chunk21 = new formattingSettings.TextInput({ name: "chunk21", displayName: "chunk21", placeholder: "", value: "" });
    chunk22 = new formattingSettings.TextInput({ name: "chunk22", displayName: "chunk22", placeholder: "", value: "" });
    chunk23 = new formattingSettings.TextInput({ name: "chunk23", displayName: "chunk23", placeholder: "", value: "" });
    chunk24 = new formattingSettings.TextInput({ name: "chunk24", displayName: "chunk24", placeholder: "", value: "" });
    chunk25 = new formattingSettings.TextInput({ name: "chunk25", displayName: "chunk25", placeholder: "", value: "" });
    chunk26 = new formattingSettings.TextInput({ name: "chunk26", displayName: "chunk26", placeholder: "", value: "" });
    chunk27 = new formattingSettings.TextInput({ name: "chunk27", displayName: "chunk27", placeholder: "", value: "" });
    chunk28 = new formattingSettings.TextInput({ name: "chunk28", displayName: "chunk28", placeholder: "", value: "" });
    chunk29 = new formattingSettings.TextInput({ name: "chunk29", displayName: "chunk29", placeholder: "", value: "" });
    chunk30 = new formattingSettings.TextInput({ name: "chunk30", displayName: "chunk30", placeholder: "", value: "" });
    chunk31 = new formattingSettings.TextInput({ name: "chunk31", displayName: "chunk31", placeholder: "", value: "" });
    chunk32 = new formattingSettings.TextInput({ name: "chunk32", displayName: "chunk32", placeholder: "", value: "" });
    chunk33 = new formattingSettings.TextInput({ name: "chunk33", displayName: "chunk33", placeholder: "", value: "" });
    chunk34 = new formattingSettings.TextInput({ name: "chunk34", displayName: "chunk34", placeholder: "", value: "" });
    chunk35 = new formattingSettings.TextInput({ name: "chunk35", displayName: "chunk35", placeholder: "", value: "" });
    chunk36 = new formattingSettings.TextInput({ name: "chunk36", displayName: "chunk36", placeholder: "", value: "" });
    chunk37 = new formattingSettings.TextInput({ name: "chunk37", displayName: "chunk37", placeholder: "", value: "" });
    chunk38 = new formattingSettings.TextInput({ name: "chunk38", displayName: "chunk38", placeholder: "", value: "" });
    chunk39 = new formattingSettings.TextInput({ name: "chunk39", displayName: "chunk39", placeholder: "", value: "" });
    chunk40 = new formattingSettings.TextInput({ name: "chunk40", displayName: "chunk40", placeholder: "", value: "" });
    chunk41 = new formattingSettings.TextInput({ name: "chunk41", displayName: "chunk41", placeholder: "", value: "" });
    chunk42 = new formattingSettings.TextInput({ name: "chunk42", displayName: "chunk42", placeholder: "", value: "" });
    chunk43 = new formattingSettings.TextInput({ name: "chunk43", displayName: "chunk43", placeholder: "", value: "" });
    chunk44 = new formattingSettings.TextInput({ name: "chunk44", displayName: "chunk44", placeholder: "", value: "" });
    chunk45 = new formattingSettings.TextInput({ name: "chunk45", displayName: "chunk45", placeholder: "", value: "" });
    chunk46 = new formattingSettings.TextInput({ name: "chunk46", displayName: "chunk46", placeholder: "", value: "" });
    chunk47 = new formattingSettings.TextInput({ name: "chunk47", displayName: "chunk47", placeholder: "", value: "" });
    chunk48 = new formattingSettings.TextInput({ name: "chunk48", displayName: "chunk48", placeholder: "", value: "" });
    chunk49 = new formattingSettings.TextInput({ name: "chunk49", displayName: "chunk49", placeholder: "", value: "" });

    name: string = "historySettings";
    displayName: string = "历史记录";
    slices: Array<FormattingSettingsSlice> = [
        this.lastActiveSessionId, this.chunkCount,
        this.chunk0, this.chunk1, this.chunk2, this.chunk3, this.chunk4,
        this.chunk5, this.chunk6, this.chunk7, this.chunk8, this.chunk9,
        this.chunk10, this.chunk11, this.chunk12, this.chunk13, this.chunk14,
        this.chunk15, this.chunk16, this.chunk17, this.chunk18, this.chunk19,
        this.chunk20, this.chunk21, this.chunk22, this.chunk23, this.chunk24,
        this.chunk25, this.chunk26, this.chunk27, this.chunk28, this.chunk29,
        this.chunk30, this.chunk31, this.chunk32, this.chunk33, this.chunk34,
        this.chunk35, this.chunk36, this.chunk37, this.chunk38, this.chunk39,
        this.chunk40, this.chunk41, this.chunk42, this.chunk43, this.chunk44,
        this.chunk45, this.chunk46, this.chunk47, this.chunk48, this.chunk49
    ];
}

/**
 * Visual formatting settings model class
 */
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
