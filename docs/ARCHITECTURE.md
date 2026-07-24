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
       │              └─ SummaryService → 配置的在线 API
       └─ 搜索按钮
            └─ 用户输入文字 → SearchPlannerService → 在线 API
                 └─ SearchPlanParser
                      └─ AqsQueryCompiler
                           └─ OutlookSearchService → Explorer.Search
```

两个功能的数据链路有意完全分开：

- 摘要链路可以接触当前邮件及同一 Conversation 的历史邮件，但必须由
  按钮触发。
- 搜索链路没有 `EmailSnapshot` 或 `MailItem` 参数，因此无法上传邮件。

## 分层职责

### `AddIn`

`Connect.cs` 是 COM 入口，负责：

- classic Outlook 加载项生命周期。
- 构建并显示右侧自定义任务窗格。
- 提供 Home / Message 功能区开关。
- 组装各服务，但不承载业务逻辑。

### `UI`

- `AssistantPaneControl.cs`：摘要、搜索、设置三个页面。
- `SettingsForm.cs`：提供商、地址、模型和 API Key 配置。

UI 只协调用户动作。可测试的转换逻辑必须留在服务层。

### `Outlook`

- `OutlookContextService.cs`：仅在摘要按钮点击后读取当前邮件和同一
  Conversation 中跨文件夹的历史邮件，排序、去重、限制为 50 封后，
  转换成与 COM 脱离的 `ConversationSnapshot`。
- `OutlookSearchService.cs`：只接收已经编译好的 AQS，并调用
  `Explorer.Search`。
- `ComRelease.cs`：释放临时 Outlook COM 引用。

### `AI`

`OpenAiCompatibleClient.cs` 实现最小化的
`POST /chat/completions` 客户端。所有提供商共用同一传输层：

- Bearer API Key
- `model`
- `messages`
- 非流式响应
- 搜索时优先请求 JSON mode；兼容服务不支持时回退一次

### `Configuration`

- `AiProviderPreset.cs`：提供商默认地址和模型。
- `SettingsStore.cs`：JSON 设置与 Windows DPAPI。
- `AppSettings.cs`：持久化模型。

### `Summary`

- `EmailBodyCleaner.cs`：只清理独立历史邮件中高置信度的重复引用。
  当前邮件原文不清理，以保留只存在于引用中的外部往来。
- `ConversationPromptBuilder.cs`：分配总字符预算，并明确区分历史背景和
  当前邮件。
- `SummaryService.cs`：调用用户配置的 API。

邮件正文被明确标注为不可信数据，不能改变系统任务。

### `Search`

1. `SearchPlannerService` 只发送用户输入的描述。
2. `SearchPlanParser` 解析 snake_case JSON、限制数组大小并验证日期。
3. `AqsQueryCompiler` 从白名单字段构建 AQS，清理引号、括号和换行。
4. `SearchPlanFormatter` 把最终条件显示给用户。

模型的原始 AQS 或 DASL 永远不会直接执行。

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

- 生产源码保持 C# 5 兼容。
- 只使用 .NET Framework 4.0 编译时 API；运行环境为 4.8。
- Office PIA 从本机 GAC 自动定位。
- 当前实现仅支持 classic Outlook for Windows。

参考：

- [Microsoft ICustomTaskPaneConsumer](https://learn.microsoft.com/en-us/office/vba/api/office.icustomtaskpaneconsumer)
- [Microsoft Conversation.GetTable](https://learn.microsoft.com/en-us/office/vba/api/outlook.conversation.gettable)
- [Microsoft MailItem.GetConversation](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.getconversation)
- [Microsoft Outlook Instant Search](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-use-instant-search-to-search-all-folders-and-all-stores-for-a-phrase-in-the-subject)
