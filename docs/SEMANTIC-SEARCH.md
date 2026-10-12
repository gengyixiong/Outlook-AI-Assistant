# 本地语义搜索

本版本在 Classic Outlook COM 加载项中新增了完全在本机运行的语义邮件搜索。
它与原有功能（Executive Brief、联系人索引、自然语言 AQS 搜索）并存，互不影响。

## 目标

用户用中文描述记忆中的邮件内容，系统找到最相关的**真实**原始邮件，显示准确
标题、发件人、日期与正文片段，并可一键在 Outlook 中打开。

示例：

```text
输入：之前印度客户问过幕墙玻璃 Panel 能不能单独编辑的邮件
邮件：Regarding the curtain wall edit, can we also edit the Panels individually?
```

中文描述与英文邮件之间不需要共享任何关键词。

## 架构

```text
AssistantPaneControl（语义搜索页）
  └─ SemanticSearchFacade
       ├─ SemanticSearchClient  ──HTTP(127.0.0.1)──►  embedding_service (Python)
       └─ OutlookMailLocator                          ├─ EmbeddingGemma 2（本地）
            └─ NameSpace.GetItemFromID                ├─ SQLite 元数据 + 向量
                 └─ MailItem.Display()                ├─ FTS5 关键词检索
                                                      └─ RRF 混合排序
索引链路：
AssistantPaneControl（建立/更新索引）
  └─ SemanticIndexService
       └─ OutlookMailScanner（仅 UI 线程访问 COM）
            ├─ MAPIFolder.GetTable()：只读元数据
            └─ MailItem：仅在需要新建/更新时读取正文
                 └─ SemanticSearchClient.IndexMailAsync → 本地服务
```

职责边界：

- **C# 插件**负责 Outlook COM 访问、邮件文本提取、UI、原邮件定位。
- **Python 服务**负责模型加载、文本清洗、分块、向量化、SQLite 存储与检索。
- 插件**从不**直接读写 SQLite，也不把候选邮件或正文发送到云端。

## 模型

- 模型：`google/embeddinggemma-2`（Gemma 4 架构多模态模型）。
- 仅加载文本编码器：`config_kwargs={"vision_config": None, "audio_config": None}`，
  实际加载 270M / 740M 参数。
- 输出维度：768（支持 MRL 截断 512/256/128，默认关闭以保质量）。
- 精度：CPU 使用 `float32`；支持的 GPU 使用 `bfloat16`。**绝不使用 `float16`**，
  因为该模型激活值超出 float16 动态范围，会静默产生 NaN 或退化向量。
- Query 前缀：`task: search result | query: `（官方非对称检索格式）。
- Document 前缀：`title: {真实邮件标题} | text: `。标题由加载项注入，
  因为 `prompt_name="Document"` 只会提供 `title: none`。
- 向量归一化：模型自带 `Normalize` 模块，且在 MRL 截断后重新归一化。

模型首次运行时下载并缓存于 `%LOCALAPPDATA%\OutlookAiAssistant\models`，
之后完全离线可用。

### 依赖版本

模型官方要求（见其 `config_sentence_transformers.json`）：

| 依赖 | 版本 |
| --- | --- |
| sentence-transformers | >= 6.1.0 |
| transformers | ~5.18 |
| torch | 2.14 |

此外，即使关闭视觉编码器，`EmbeddingGemma2Processor` 仍会在导入时
需要 **`torchvision`** 和 **`pillow`**，因此必须一并安装。

## 索引设计

### 范围（默认）

- 最近 12 个月，最多 3,000 封。
- 包含收件箱、已发送邮件和用户自建邮件文件夹（可在设置中调整）。
- 排除已删除邮件、垃圾邮件、草稿和发件箱。
- 安装后**不自动扫描**；必须由用户点击“建立 / 更新语义索引”。

### 每封邮件保存

EntryID、StoreID、InternetMessageId、ConversationID、真实标题、发件人名称与
邮箱、收件人、ReceivedTime/SentOn、LastModificationTime、FolderID 与
FolderPath、附件名称与数量、清洗后的正文、正文 Chunk、Chunk 向量、索引时间、
内容 Hash、索引状态。

### SQLite Schema

- `mails`：每封邮件一行，`UNIQUE(store_id, entry_id)`；
  另按 `internet_message_id`、`conversation_id`、`received_at`、
  `last_modified_at`、`index_status` 建索引。
- `chunks`：`mail_id` 外键 `ON DELETE CASCADE`，`UNIQUE(mail_id, chunk_index)`，
  向量以 float32 BLOB 存储。
- `chunks_fts`：FTS5 外部内容表，通过触发器与 `chunks` 保持同步；
  删除邮件即级联删除片段与 FTS 行，绝不留下孤立向量。

### 索引的一致性与增量

- 内容 Hash = `sha256(subject | sender | cleaned_body | attachment_names)`。
- 增量同步时，若 hash 未变且有向量，则只更新文件夹位置（记录移动），
  不重新计算向量。
- EntryID 会因移动文件夹而改变，因此优先用 `internet_message_id` 匹配同一封邮件。
- 单封邮件失败只记录并跳过，不中断整体索引。
- 每封邮件单独事务写入，中断不会留下半写状态；SQLite 使用 WAL 模式。

### Outlook 扫描性能与线程

- 优先使用 `MAPIFolder.GetTable()` 读取元数据，按 `ReceivedTime` 倒序，
  越过时间下限即停止该文件夹。
- 只对需要新建/更新的邮件读取 `MailItem` 正文。
- 所有 COM 调用都在 Outlook UI 线程上执行；后台任务通过 `Control.Invoke`
  回主线程访问 COM。
- 每个临时 COM 对象都通过 `ComRelease` 释放，不跨线程传递 COM 对象。
- 索引过程分批、显示进度、避免长时间阻塞 UI。

## 邮件正文处理

复用并扩展 `Summary/EmailBodyCleaner.cs` 的思路（Python 侧为
`app/body_cleaner.py`）：

- HTML 转文本（`HTMLBody` 在 `Body` 为空时回退使用）。
- 去除 Outlook 回复引用分隔符（`-----Original Message-----` 等）。
- 通过标题块信号（From/Sent/To/Subject）识别引用。
- 去除 `--` 签名块。
- 对形似提示词注入的行加 `[untrusted-content]` 标记。
- 分块 300–800 tokens（目标 512），相邻块重叠 64 tokens。
- 相同或高度重复的片段去重，避免同一 Conversation 的重复引用被反复向量化。
- 邮件标题作为每个 Chunk 的头部上下文，避免标题关键词被不合理放大。

## 混合检索

1. **语义检索**：Query 向量与所有 Chunk 向量点积，按邮件聚合（取最高分，
   多片段给少量加权），避免一封长邮件占满结果。
2. **关键词检索**：FTS5 + BM25。
   - **重要**：SQLite 内置 trigram 分词器无法匹配中文，也无法匹配
     两字符术语（如“幕墙”“AI”“3D”）。
   - 因此服务对文本做规范化分词（CJK 生成重叠 bigram 与单字，拉丁/数字
     按词切分），用 `unicode61` 索引预分词字段，从而覆盖中英文与短术语。
3. **融合**：RRF（Reciprocal Rank Fusion，k=60），只比较排名，不需要
   校准两路分数。
4. 结构化过滤：发件人、时间范围、是否有附件作为 `mails` 表的可验证条件，
   在检索前缩小候选集。

初始候选：语义 Top 40 + 关键词 Top 40，融合后返回 Top 10。

## 原邮件定位

主路径：

```csharp
session.GetItemFromID(entryId, storeId) → MailItem.Display()
```

失败时：

1. 用默认 Store 再试一次（EntryID 可能因移动而变化）。
2. 按 `InternetMessageId` 跨 Store 查找，并用元数据校验。
3. 严格元数据匹配兜底；绝不打开标题相同但身份不同的邮件。
4. 全部失败时给出明确提示，并提供“Outlook 搜索”按钮，
   复用现有 `OutlookSearchService`。

## 服务安全

- 只监听 `127.0.0.1`，不暴露到局域网。
- 除 `/health` 外都需要 Bearer Token；Token 由服务首次启动时生成，
  保存于 `%LOCALAPPDATA%\OutlookAiAssistant\semantic-service-token`
  （POSIX 上权限 0600）。
- C# 客户端强制只使用 loopback 地址；即使配置了局域网地址也会退回
  `http://127.0.0.1:8765`。
- 服务不可用不会导致 Outlook 崩溃：所有调用都返回友好的错误信息。

## 本地数据清理

- “删除本地索引”会删除本机保存的全部邮件正文副本、片段与向量。
- Outlook 中的原始邮件不受影响。

## 已知限制

- 模拟数据的检索评估只验证管线，不代表真实模型在真实邮箱上的质量；
  真实邮箱评估需用户在自己的邮箱上完成（见 `docs/TEST-CHECKLIST.md`）。
- 首版不含 Reranker、附件内容索引、图片/多模态索引和自动后台同步。
- FTS5 为精确 token 匹配，不做前缀扩展。
