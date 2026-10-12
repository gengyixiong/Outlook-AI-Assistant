# Outlook AI Assistant

面向 classic Outlook for Windows 的个人 AI 邮件助手。

## 实施路线

当前正式实施路线是桌面端 COM 加载项，直接运行在用户自己的 Classic Outlook 中：

1. 用户点击按钮后，加载项读取当前邮件或当前 Conversation 的必要内容，生成 Executive Brief。
2. 本地语义搜索：用户用中文描述邮件内容，由本机模型与混合检索找到最相关的原始邮件，并一键在 Outlook 中打开。全部在本机运行。

## 适配环境

- Windows 10 / 11
- .NET Framework 4.8
- Classic Outlook for Windows 2016 或更高版本
- 已验证：classic Outlook 2024 LTSC 64 位，Build `16.0.17932.20574`
- 世纪互联运营的 Microsoft 365 / Exchange Online

## 功能亮点

Outlook AI Assistant 主要解决两个问题：快速看懂长邮件，以及用中文描述找回历史邮件。

### 1. 快速看懂长邮件

自动结合当前邮件和历史 Conversation，生成简短的 Executive Brief，重点告诉你：

- 这件事现在是什么状态
- 最新邮件有什么变化
- 有没有需要你回复、决定或推动的事项
- 是否存在重要风险、日期或关键数字

它适合快速处理长邮件线程，而不是生成冗长的会议纪要。只有用户点击“生成 Executive Brief”后，加载项才会读取邮件；它不会监听新邮件、自动摘要或批量读取邮箱。

### 2. 本地语义搜索

很多时候你只记得邮件的内容大意，却完全不记得邮件里用了什么词。例如：

```text
之前印度客户问过幕墙玻璃 Panel 能不能单独编辑的邮件
```

而邮件原文可能是：

```text
Regarding the curtain wall edit, can we also edit the Panels individually?
```

“语义搜索”页会用完全在本机运行的模型理解这两种描述在语义上的相关性，
返回**真实**的原始邮件：

- 中文描述可以直接找到英文邮件，不需要共享关键词。
- 结合语义向量、SQLite 全文关键词（含中文与短术语）和邮件元数据。
- 结果按相关度排序，显示真实标题、发件人、日期、文件夹和正文片段。
- 点击“打开邮件”直接在 Outlook 中打开对应原始邮件；
  失败时可一键回退到 Outlook 原生搜索。

关键隐私保证：

- `google/embeddinggemma-2` 模型在本机 Python 进程中运行。
- 邮件正文与向量只保存在本机 SQLite，**绝不**发送到云端 Embedding API。
- 本地服务只监听 `127.0.0.1`，并使用本地随机 Token 鉴权。
- 不安装这个可选组件时，原有功能完全不受影响。

详见 [docs/SEMANTIC-SEARCH.md](docs/SEMANTIC-SEARCH.md)。

### 3. Privacy-first

- 语义搜索完全在本机运行，不依赖任何外部 LLM，也不上传邮件内容。
- 本地服务只监听 `127.0.0.1`，不对局域网开放。
- Executive Brief 只在用户主动点击后读取当前邮件及其 Conversation 上下文，
  并发送给用户自己配置的 OpenAI-compatible API。
- 不依赖 Microsoft Graph、EWS 或云端 Vector Database。
- API Key 使用 Windows DPAPI 加密，只能由当前 Windows 用户解密。

## 如何安装

1. 从 GitHub Release 下载最新发布包并解压（**不要在压缩包里直接运行**）。
2. 完全退出 classic Outlook。
3. 双击解压目录中的 **`setup.cmd`**。它会自动完成全部安装：
   - 检查 Python 3.10–3.12
   - 安装 / 更新 Outlook 加载项
   - 安装 / 更新本地语义搜索服务
   - 创建 Python 依赖环境（一次）
   - 下载本地模型（一次，约 1.5 GB）

   `setup.cmd` 可以重复运行：版本相同时只补齐缺失部分，其余自动跳过；
   升级时覆盖代码但保留依赖环境、模型、索引与设置。

4. 重新启动 Outlook，右侧出现“AI 邮件助手”。
5. 在“设置”中填写 OpenAI-compatible Base URL、API Key、Model 和
   Reasoning Effort。
6. 在“语义搜索”页点击“启动服务”，再点击“建立 / 更新语义索引”。
7. 安装完成后可以删除解压出来的文件夹。

安装不需要管理员权限。安装脚本把加载项复制到当前用户的
`%LOCALAPPDATA%\OutlookAiAssistantpp`，服务安装到
`%LOCALAPPDATA%\OutlookAiAssistant\embedding_service`，只写当前用户注册表。

更新：关闭 Outlook 后对新的发布包再次运行 `setup.cmd`。

卸载：关闭 Outlook 后运行 `Uninstall.cmd`（保留设置、密钥与索引）。
彻底删除本地数据请加 `-RemoveSettings`。

## 许可证

本项目采用 [MIT License](LICENSE)。

Copyright (c) 2026 Yixiong Geng
