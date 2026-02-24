# 为什么请求负载里会出现两个 user content

在 `模板.js` 的 `callAIAPI(t)` 逻辑中：

1. 先拿到用户原始输入 `t`。
2. 再调用 `prepareDataContext()` 生成当前报表页的数据上下文 `r`。
3. 如果 `r` 存在，就把 `l` 从“原始问题”改写成一个**增强问题**，格式是：

```text
数据上下文：
${r}

用户问题：${t}

请基于提供的数据回答用户问题...
```

4. 请求体最终是：

```ts
messages: [
  { role: "system", content: h },
  ...c,
  { role: "user", content: l }
]
```

其中 `c = getConversationHistoryMessages()`，会把历史对话拼进去（最多 10 条，去掉最后一条 loading/error）。

## 两种常见情况

### 情况 A：非首次提问（已有历史 user）

会出现两个 `user` 很正常：

- 前面的 `user` 来自历史对话（历史问题）。
- 末尾的 `user` 是当前轮增强问题 `l`（数据上下文 + 用户问题 + 输出要求）。

### 情况 B：首次提问（你问的这个）

理论上如果历史里还没有任何 `user`，请求里通常只会有 **1 条 `user`**（即末尾这条增强问题）。

但你抓包里“首次提问也出现 2 条 user”的原因通常是：

- 当前这条原始问题（例如“帮我分析本页数据”）在调用 API 前，已经先写入了 `chatMessages`；
- `getConversationHistoryMessages()` 读取历史时把它也当成了“历史 user”；
- 然后 `callAIAPI(t)` 末尾又追加一次增强后的当前问题 `l`。

结果就是：同一轮里出现“原始 user + 增强 user”两条 `user`。

## 一句话结论

你看到两个 `user` 不一定表示多发了一轮请求，而是**同一轮请求在做“历史拼接 + 当前增强”时把当前问题以两种形态都带上了**。
