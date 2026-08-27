# Outlook AI Assistant

面向 classic Outlook for Windows 的个人 AI 邮件助手。

## 实施路线

当前正式实施路线是桌面端 COM 加载项，直接运行在用户自己的 Classic Outlook 中：

1. 用户点击按钮后，加载项读取当前邮件或当前 Conversation 的必要内容，生成 Executive Brief。
2. 用户输入自然语言描述后，AI 负责理解搜索意图；加载项在本地把意图转换为受控的 Outlook AQS 查询。
3. 联系人、别名、Outlook 归类文件夹和附件线索只用于辅助理解搜索目标，最终邮件搜索由 Outlook 本地 Instant Search 执行。
4. 联系人索引只在用户手动点击“重建联系人索引”时更新，不后台扫描、不监听新邮件。

## 适配环境

- Windows 10 / 11
- .NET Framework 4.8
- Classic Outlook for Windows 2016 或更高版本
- 已验证：classic Outlook 2024 LTSC 64 位，Build `16.0.17932.20574`
- 世纪互联运营的 Microsoft 365 / Exchange Online

## 功能亮点

Outlook AI Assistant 主要解决两个问题：快速看懂长邮件，以及用模糊记忆找回邮件。

### 1. 快速看懂长邮件

自动结合当前邮件和历史 Conversation，生成简短的 Executive Brief，重点告诉你：

- 这件事现在是什么状态
- 最新邮件有什么变化
- 有没有需要你回复、决定或推动的事项
- 是否存在重要风险、日期或关键数字

它适合快速处理长邮件线程，而不是生成冗长的会议纪要。只有用户点击“生成 Executive Brief”后，加载项才会读取邮件；它不会监听新邮件、自动摘要或批量读取邮箱。

### 2. 用模糊记忆搜索邮件

很多时候，你只记得：

```text
上个月某个海外代理商发来的邮件
某个同事发来的带附件邮件
之前那个产品版本对比的 Excel
某个项目相关的邮件
```

却记不住准确的人名、邮箱、公司名、附件文件名或搜索关键词。

Outlook AI Assistant 会把这类描述转换为可执行的本地搜索：

```text
自然语言
→ AI 理解搜索意图
→ 联系人 / 日期 / 关键词 / 附件 / 项目
→ 本地生成 Outlook 搜索条件
→ Classic Outlook Instant Search
```

系统会从 Outlook 历史邮件建立轻量 `contact-index.json`，并综合多种线索理解你的模糊记忆，包括：

- 中文名、拼音、英文名和联系人别名
- Display Name、Email、公司、国家、地区和业务关系
- Outlook 左侧栏中 Inbox 旁边的自建归类文件夹名称和路径
- 文件夹中邮件关联的联系人
- 附件名称、文件类型和“是否带附件”条件

这些线索只用于帮助 AI 理解“你说的是谁、哪件事或哪个文件”；AI 返回的联系人只能在本地索引中验证，不能直接变成任意邮箱条件。最终查询仍由 Classic Outlook 本地执行，用户可以在宽松、推荐和精确三档搜索范围之间切换。

联系人索引的位置为：

```text
%LOCALAPPDATA%\OutlookAiAssistant\contact-index.json
```

### 3. Privacy-first

AI 负责理解，Outlook 负责搜索。

- 搜索不会上传整个邮箱、候选邮件正文、附件内容或 Outlook 搜索结果。
- 邮件搜索结果不发送给 AI 做二次排序，最终查询由 Classic Outlook 本地 Instant Search 完成。
- Executive Brief 只在用户主动点击后读取当前邮件及其 Conversation 上下文。
- 联系人索引只保留帮助识别联系人和搜索意图所需的轻量信息，不保存完整邮箱内容。
- 不依赖 Microsoft Graph、EWS、SQLite、Vector Database 或 RAG。
- API Key 使用 Windows DPAPI 加密，只能由当前 Windows 用户解密。

## 如何安装

1. 从 GitHub Release 下载最新发布包并解压。
2. 完全退出 classic Outlook。
3. 双击 `Install.cmd`。
4. 重新启动 Outlook。
5. 在右侧栏“设置”中选择 DeepSeek Flash 或 GLM-5.3 Flash，并填写对应 API Key。
6. 如需联系人别名、公司或 Outlook 归类文件夹辅助搜索，在设置中点击“重建联系人索引”。

安装不需要管理员权限。安装脚本会把插件复制到当前用户的 `%LOCALAPPDATA%\OutlookAiAssistant\app` 并写入当前用户注册表；安装完成后可以删除下载并解压的发布包目录。

更新时，关闭 Outlook 后运行新版本的 `Install.cmd` 即可；已有设置、加密 API Key 和联系人索引会保留。

卸载时，关闭 Outlook 后运行 `Uninstall.cmd`。默认保留设置、加密 API Key 和联系人索引，方便以后重新安装。

## 许可证

本项目采用 [MIT License](LICENSE)。

Copyright (c) 2026 Yixiong Geng
