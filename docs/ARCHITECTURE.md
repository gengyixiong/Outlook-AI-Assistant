# 架构说明

## 总体设计

```text
Outlook UI
  └─ AssistantPaneControl
       ├─ 摘要按钮
       │    └─ OutlookContextService
       │         └─ 当前 MailItem + Conversation.GetTable
       │              └─ ConversationSnapshot
       │                   ├─ 当前 EmailSnapshot（完整原文）
       │                   └─ 历史 EmailSnapshot（排序、去重、上限取样）
       │              └─ Executive Brief Prompt → 配置的 OpenAI-compatible API
       └─ 语义搜索页
            ├─ SemanticSearchFacade
            │    ├─ SemanticSearchClient ──HTTP(127.0.0.1)──► embedding_service
            │    └─ OutlookMailLocator → GetItemFromID → MailItem.Display()
            ├─ SemanticIndexService
            │    └─ OutlookMailScanner（仅 UI 线程访问 COM）
            │         ├─ MAPIFolder.GetTable() 只读元数据
            │         └─ MailItem 正文（仅需新建/更新时）
            └─ SemanticServiceLauncher → start-service.cmd（独立控制台进程）
```

两条功能链路有意分开：

- 摘要链路可以接触当前邮件及同一 Conversation 的历史邮件，但必须由
  按钮触发。
- 本地语义搜索链路不向任何云端发送邮件内容；正文与向量只在本机
  Python 服务与本地 SQLite 之间流动。

自然语言 AQS 搜索和联系人索引已在 0.6.0 中移除。`OutlookSearchService`
保留下来，只用于语义搜索结果卡片上的“Outlook 搜索”回退按钮。

## 分层职责

### `AddIn`

`Connect.cs` 是 COM 入口，负责：

- classic Outlook 加载项生命周期。
- 构建并显示右侧自定义任务窗格。
- 提供 Home / Message 功能区开关。
- 组装各服务，但不承载业务逻辑。
- 创建语义搜索门面（`CreateSemanticFacade`），每次重新读取本地服务 Token，
  使服务重启后无需重启 Outlook。

### `UI`

- `AssistantPaneControl.cs`：摘要、语义搜索、设置三个页面。
  语义搜索页顶部常驻显示服务/模型状态，并提供“启动服务”“下载模型”
  “建立 / 更新语义索引”“停止索引”等操作。
  语义搜索页显示真实结果卡片（标题、发件人、日期、文件夹、片段）
  并提供“打开邮件”和“Outlook 搜索”按钮。
- `SettingsForm.cs`：OpenAI-compatible Base URL、API Key、可编辑 Model、
  Reasoning Effort、摘要身份，以及语义索引范围、增量同步、重新构建、
  删除本地索引，并显示模型/索引/日志目录（可一键打开）与磁盘占用。

UI 只协调用户动作。可测试的转换逻辑必须留在服务层。

### `Outlook`

- `OutlookContextService.cs`：仅在摘要按钮点击后读取当前邮件和同一
  Conversation 中跨文件夹的历史邮件，排序、去重、限制为 50 封后，
  转换成与 COM 脱离的 `ConversationSnapshot`。
- `OutlookSearchService.cs`：接收标题等条件并调用 `Explorer.Search`，
  仅作为语义搜索结果卡片上的“Outlook 搜索”回退路径。
- `ComRelease.cs`：释放临时 Outlook COM 引用。

### `AI`

`OpenAiCompatibleClient.cs` 实现最小化的
`POST /chat/completions` 客户端。所有提供商共用同一传输层：

- Bearer API Key
- `model`
- `reasoning_effort`（`none`、`medium` 或 `max`）
- `messages`
- 非流式响应
- 搜索时优先请求 JSON mode；兼容服务不支持时回退一次

服务地址、模型和 Reasoning Effort 均来自用户设置。新安装默认使用
`https://api.openai.com/v1`、`gpt-5.6-luna` 和 `none`；本版本不自动获取模型列表。

### `Configuration`

- `SettingsStore.cs`：JSON 设置与 Windows DPAPI。
- `AppSettings.cs`：持久化 Base URL、Model、Reasoning Effort、加密密钥和摘要身份。

加载设置时会规范化并验证 Base URL、Model 和 Reasoning Effort，不会覆盖用户输入。
旧版 `deepseek` 与 `zhipu` 设置会保留其地址、模型和加密 API Key，迁移后按普通兼容服务使用。

### `Summary`

- `EmailBodyCleaner.cs`：只清理独立历史邮件中高置信度的重复引用。
  当前邮件原文不清理，以保留只存在于引用中的外部往来。
- `ConversationPromptBuilder.cs`：分配总字符预算，合并 Outlook 自动身份和
  用户别名，并生成最多四段的 Executive Brief 指令。
- `SummaryService.cs`：调用用户配置的 API。

邮件正文被明确标注为不可信数据，不能改变系统任务。

### `SemanticSearch`

本地语义检索层，与云端 AI 完全解耦：

- `SemanticSearchClient.cs`：loopback-only HTTP 客户端，负责与
  Python 服务通信。强制只使用 `127.0.0.1`，使用共享 Token 鉴权，
  服务不可用时返回友好错误而不抛出崩溃性异常。
- `SemanticIndexService.cs`：索引编排。在 UI 线程枚举候选元数据，
  逐封读取正文并发送给本地服务；支持进度、取消与单封失败跳过。
- `OutlookMailScanner.cs`：使用 `MAPIFolder.GetTable()` 优先读取元数据，
  只在需要时读取 `MailItem` 正文；遵守 COM 线程限制并释放所有临时引用。
- `OutlookMailLocator.cs`：`GetItemFromID` 定位原邮件并用
  `InternetMessageId` 等元数据校验身份，绝不打开同标题的其他邮件；
  失败时提示用户并回退到 Outlook 原生搜索。
- `SemanticSearchFacade.cs`：UI 门面，把服务响应转为展示数据，
  并集中处理错误，保证服务未运行时 Outlook 不崩溃。

### `embedding_service/`（Python）

本地语义服务，详见 `docs/SEMANTIC-SEARCH.md`：

- `app/embedding_model.py`：EmbeddingGemma 2 文本编码器封装，仅文本、
  CPU float32 默认、官方非对称 prompt、文档注入真实标题。
- `app/database.py`：SQLite 元数据、片段、向量（BLOB）与 FTS5 全文索引。
- `app/text_utils.py`：分块与 CJK 友好分词（弥补 FTS5 trigram 对中文和
  短术语的不足）。
- `app/search.py`：语义检索 + 关键词检索 + RRF 融合。
- `app/indexing.py`：分块、向量化与每封邮件事务写入。
- `app/api.py`：FastAPI 端点，loopback + Bearer Token。

## COM 注册

安装器在当前用户注册以下对象：

```text
OutlookAiAssistant.Connect
  CLSID {8A435B2A-4E46-4557-9300-BD94F8C92F82}

OutlookAiAssistant.TaskPane
  CLSID {60DF60D1-BB82-4606-B245-F79DD62C99AC}
```

加载项注册位于：

```text
HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookAiAssistant.Connect
```

COM 类位于 `HKCU\Software\Classes`，因此安装不需要管理员权限。

## 兼容性约束

- 生产源码以 .NET Framework 4.8 为目标，并保持 C# 5 兼容。
- Office PIA 从本机 GAC 自动定位。
- 当前实现仅支持 classic Outlook for Windows。
- 本地语义服务使用独立的 Python 虚拟环境；加载项不依赖任何 NuGet 包。
- Python 服务不可用不影响原有功能。

参考：

- [Microsoft ICustomTaskPaneConsumer](https://learn.microsoft.com/en-us/office/vba/api/office.icustomtaskpaneconsumer)
- [Microsoft Conversation.GetTable](https://learn.microsoft.com/en-us/office/vba/api/outlook.conversation.gettable)
- [Microsoft MailItem.GetConversation](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.getconversation)
- [Microsoft Outlook Instant Search](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-use-instant-search-to-search-all-folders-and-all-stores-for-a-phrase-in-the-subject)
