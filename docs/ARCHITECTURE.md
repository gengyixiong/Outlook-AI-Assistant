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
       ├─ 设置中的“重建联系人索引”按钮
       │    ├─ OutlookContactScanner → 轻量字段聚合 + 自建文件夹标签
       │    ├─ 每人最早 1 + 最近 3 封代表邮件 → 配置的 OpenAI-compatible API
       │    └─ ContactIndexStore → contact-index.json
       └─ 搜索按钮
            └─ 用户输入文字 + Contact Index → SearchPlannerService → 配置的 OpenAI-compatible API
                 └─ SearchPlanParser
                      └─ Contact ID 本地验证 + 附件/关键词条件
                           └─ AqsQueryCompiler
                                ├─ 宽松 AQS（默认）
                                ├─ 推荐 AQS
                                └─ 精确 AQS
                                     └─ OutlookSearchService → Explorer.Search
```

三个功能的数据链路有意分开：

- 摘要链路可以接触当前邮件及同一 Conversation 的历史邮件，但必须由
  按钮触发。
- 联系人索引链路只能由设置按钮触发；首轮只读轻量字段，正文只读每位
  联系人的少量代表邮件。
- 搜索链路没有 `EmailSnapshot` 或 `MailItem` 参数，只发送用户描述和已经
  保存、验证的 Contact Index，不上传搜索候选或结果。

## 分层职责

### `AddIn`

`Connect.cs` 是 COM 入口，负责：

- classic Outlook 加载项生命周期。
- 构建并显示右侧自定义任务窗格。
- 提供 Home / Message 功能区开关。
- 组装各服务，但不承载业务逻辑。

### `UI`

- `AssistantPaneControl.cs`：摘要、搜索、设置三个页面。
- `SettingsForm.cs`：OpenAI-compatible Base URL、API Key、可编辑 Model、Reasoning Effort、摘要身份和联系人索引
  状态/手动重建。

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

### `EntityIndex`

- `OutlookContactScanner.cs`：遍历 Classic Outlook 本地邮件，先按发件人邮箱
  聚合轻量字段，同时把自建邮件文件夹的相对路径及其中联系人记录为人工标签，
  再读取每人最早 1 封和最近最多 3 封代表邮件。
- `ContactProfileExtractor.cs`：用当前配置的 OpenAI-compatible API 提取公司、国家、城市、
  职位、业务关系和 aliases；邮箱与 Contact ID 始终由本地代码决定。
- `ContactIndexStore.cs`：维护唯一正式 JSON、结构验证、旧索引备份和原子替换。
- `ContactIndexService.cs`：仅响应设置按钮，串联上述步骤；不注册定时任务或
  Outlook 事件。

### `Search`

1. `SearchPlannerService` 发送用户描述和已验证的 Contact Index。Contacts 是
   主数据，Folder Context 只是不可作为排除依据的辅助信息。
2. `SearchPlanParser` 解析 snake_case JSON、限制数组大小并验证日期。
   如果模型误把“巴西客户、合作伙伴、主办方”等关系描述放入人员字段，
   解析器会将其降级为精确档软提示，避免生成无效的 `from:` 条件。
3. `ContactIndexResolver` 只接受本地索引中存在的 Contact ID，并解析为真实
   邮箱；Folder Context 永远不能直接生成联系人字段。
4. `AqsQueryCompiler` 从白名单字段构建三档 AQS，清理引号、括号和换行，
   并支持 `attachments:` 与 `hasattachments:`。
   宽松档只使用第一个高价值硬关键词；没有硬关键词时使用第一个概念组。
   推荐档增加概念和明确过滤器；精确档再增加可能不逐字出现的软提示。
5. `SearchPlanFormatter` 显示模型理解和全部三档查询。
6. 用户切换档位时只重新调用本地 `Explorer.Search`，不会再次请求 AI。

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

- 生产源码以 .NET Framework 4.8 为目标，并保持 C# 5 兼容。
- Office PIA 从本机 GAC 自动定位。
- 当前实现仅支持 classic Outlook for Windows。

参考：

- [Microsoft ICustomTaskPaneConsumer](https://learn.microsoft.com/en-us/office/vba/api/office.icustomtaskpaneconsumer)
- [Microsoft Conversation.GetTable](https://learn.microsoft.com/en-us/office/vba/api/outlook.conversation.gettable)
- [Microsoft MailItem.GetConversation](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.getconversation)
- [Microsoft Outlook Instant Search](https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-use-instant-search-to-search-all-folders-and-all-stores-for-a-phrase-in-the-subject)
