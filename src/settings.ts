"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsSlice = formattingSettings.Slice;

class AISettingsCard extends FormattingSettingsCard {
    name: string = "aiSettings";
    displayName: string = "AI 设置";

    apiKey = new formattingSettings.TextInput({ name: "apiKey", displayName: "API Key", value: "", placeholder: "" });
    apiUrl = new formattingSettings.TextInput({ name: "apiUrl", displayName: "API URL", value: "", placeholder: "" });
    model = new formattingSettings.TextInput({ name: "model", displayName: "模型", value: "", placeholder: "" });

    slices: Array<FormattingSettingsSlice> = [this.apiKey, this.apiUrl, this.model];
}

class WelcomeSettingsCard extends FormattingSettingsCard {
    name: string = "welcomeSettings";
    displayName: string = "欢迎语";

    welcomeMessage = new formattingSettings.TextArea({
        name: "welcomeMessage",
        displayName: "欢迎语内容",
        value: "我是PowerBI星球打造的ABI Chat，欢迎使用。",
        placeholder: ""
    });

    slices: Array<FormattingSettingsSlice> = [this.welcomeMessage];
}

class SuggestionSettingsCard extends FormattingSettingsCard {
    name: string = "suggestionSettings";
    displayName: string = "建议问题";

    question1 = new formattingSettings.TextInput({ name: "question1", displayName: "问题1", value: "分析本页数据", placeholder: "" });
    question2 = new formattingSettings.TextInput({ name: "question2", displayName: "问题2", value: "设计一个分析框架", placeholder: "" });
    question3 = new formattingSettings.TextInput({ name: "question3", displayName: "问题3", value: "推荐合适的可视化方案", placeholder: "" });

    slices: Array<FormattingSettingsSlice> = [this.question1, this.question2, this.question3];
}

class AboutCardSettings extends FormattingSettingsCard {
    name: string = "aboutSettings";
    displayName: string = "关于";

    author = new formattingSettings.TextInput({ name: "author", displayName: "作者", value: "PowerBI星球", placeholder: "" });
    version = new formattingSettings.TextInput({ name: "version", displayName: "版本", value: "2.0.0", placeholder: "" });
    declaration = new formattingSettings.TextArea({ name: "declaration", displayName: "声明", value: "本视觉由 ABI Chat 提供支持。", placeholder: "" });

    slices: Array<FormattingSettingsSlice> = [this.author, this.version, this.declaration];

    public revertToDefault(): void {
        this.author.value = "PowerBI星球";
        this.version.value = "2.0.0";
        this.declaration.value = "本视觉由 ABI Chat 提供支持。";
    }
}

class TmdlSettingsCard extends FormattingSettingsCard {
    name: string = "tmdlSettings";
    displayName: string = "TMDL";

    tmdlCode = new formattingSettings.TextInput({ name: "tmdlCode", displayName: "TMDL Code", value: "", placeholder: "" });
    lastActiveSessionId = new formattingSettings.TextInput({ name: "lastActiveSessionId", displayName: "Last Session", value: "", placeholder: "" });
    chunkCount = new formattingSettings.TextInput({ name: "chunkCount", displayName: "Chunk Count", value: "", placeholder: "" });
    chunkSlices: FormattingSettingsSlice[];
    slices: Array<FormattingSettingsSlice> = [];

    constructor() {
        super();
        this.chunkSlices = Array.from({ length: 50 }).map((_, i) =>
            new formattingSettings.TextInput({ name: `chunk${i}`, displayName: `Chunk ${i}`, value: "", placeholder: "" })
        );
        this.slices = [this.tmdlCode, this.lastActiveSessionId, this.chunkCount, ...this.chunkSlices];
    }
}

export class FormattingSettings extends FormattingSettingsModel {
    aiSettingsCard = new AISettingsCard();
    welcomeSettingsCard = new WelcomeSettingsCard();
    suggestionSettingsCard = new SuggestionSettingsCard();
    aboutCardSettings = new AboutCardSettings();
    tmdlSettingsCard = new TmdlSettingsCard();

    cards = [this.welcomeSettingsCard, this.suggestionSettingsCard, this.aboutCardSettings];
}
