# 项目状态

更新时间：2026-10-12

## 当前路线

Office.js 个人侧载在当前世纪互联租户中不可用，因此正式路线已经切换为
classic Outlook 用户级 COM 加载项。

## 0.5.4 改进：语义服务安装到固定位置

- 新增 `install-service.cmd` / `install-service.ps1`：把语义服务（含
  `.venv`）安装到 `%LOCALAPPDATA%\OutlookAiAssistant\embedding_service`，
  并创建开始菜单快捷方式。
- 修复一个重要可用性问题：之前服务代码与 `.venv` 位于解压目录（往往是
  Downloads），删除该目录就会导致语义搜索失效。现在安装后可安全删除
  解压目录。
- 新增 `uninstall-service.ps1`：默认保留模型缓存与索引，可选用
  `-RemoveModel -RemoveIndex` 一并删除。
- 代码与数据彻底分离：服务代码在 `embedding_service`，数据库/模型/日志/
  令牌在 `%LOCALAPPDATA%\OutlookAiAssistant` 的其他子目录。
- `setup-venv.cmd` 支持 `--no-pause`，供安装脚本非交互调用。
- 新增 4 个路径独立性测试。

## 0.5.3 修复：Python 发现与 Store 别名

- `setup-venv.cmd` 现在优先使用 `py` 启动器，自动跳过 PATH 中指向
  Microsoft Store 占位程序的 `python.exe`，并探测常见安装目录。
- 当真正的 Python 不存在时，给出包含“关闭应用执行别名”步骤的明确指引，
  而不是含糊失败。
- `download-model.cmd` / `start-service.cmd` 在 venv 缺失时也给出同样指引。

## 0.5.2 修复：模型下载体验

- 修复 `download-model.cmd` / `setup-venv.cmd` / `start-service.cmd` 窗口
  一闪就关的问题：现在结束时都会 `pause`，可看到结果。
- 模型下载新增**实时进度显示**（`Fetching N files: xx%` 与下载速度），
  结束时输出 `SUCCESS` 或明确的失败原因。
- 修复一个导致下载步骤直接报错的 bug：
  `huggingface_hub` 需要 `tqdm_class` 是一个可构造的类，旧实现传了实例，
  报 `'function' object is not iterable`。
- 下载时 `max_workers=1`，避免多个进度条互相覆盖。
- CLI 默认静默 httpx/huggingface/transformers 的刷屏日志，输出更可读。
- 新增 4 个下载进度单元测试。

## 0.5.1 修复：模型未下载时的失败行为

- 修复严重缺陷：模型不可用时，`/index/mail` 会静默保存**无向量**的邮件，
  导致一个看似有几千封邮件的索引实际搜不到任何东西。现在服务直接返回
  HTTP 503，不写入任何垃圾数据。
- 新增“下载模型”步骤与界面按钮（`/model/download` 与 `/model/status`），
  在一台从未下载过模型的电脑上给出可操作提示。
- 新增真正的停止索引能力：取消令牌贯穿整个索引过程，界面提供“停止索引”
  按钮，取消后已完成邮件保留。
- 索引前先检查模型就绪状态，避免空跑整个邮箱。
- 连续 10 封邮件失败时自动停止，避免服务中断后继续无效扫描。
- 已建立的有效索引在模型丢失后仍完整保留且可搜索。
- 新增 4 个模型可用性单元测试与 3 个新端点测试。

## 0.5.0 新增：本地语义搜索

- 新增可选的本地 Python 服务 `embedding_service/`，在本机运行
  `google/embeddinggemma-2`（仅文本编码器，CPU `float32`，768 维）。
- 新增 C# `SemanticSearch/` 层：loopback HTTP 客户端、Outlook 邮件扫描、
  索引服务、原邮件定位与 UI 门面。
- Outlook 侧新增“语义搜索”页：中文描述搜索、真实结果卡片、
  “打开邮件”与“Outlook 搜索”回退。
- 设置页新增语义索引管理：索引范围、最大邮件数、增量同步、
  重新构建、删除本地索引、状态查看。
- 混合检索：向量语义 + SQLite FTS5 关键词 + RRF 融合。
- 中文/短术语分词方案（CJK bigram + 拉丁词），弥补 FTS5 trigram 不足。
- 本地服务仅监听 `127.0.0.1`，使用随机 Bearer Token 鉴权。
- 删除本地索引会删除本机保存的邮件正文副本与向量。

## 已完成

- 可维护的 C# 分层项目结构。
- Outlook COM 入口和自定义右侧任务窗格。
- Outlook 功能区“显示助手”开关。
- 按钮触发的当前邮件与同一 Outlook 会话历史读取。
- 会话邮件跨文件夹读取、时间排序、EntryID 去重和 50 封上限取样。
- Executive Brief 固定四段式输出，并结合 Outlook 当前身份、用户邮箱和称呼别名识别相关事项。
- 历史邮件重复引用清理与总字符预算。
- 支持通用 OpenAI-compatible Base URL、API Key、可编辑 Model 和 Reasoning Effort；新安装默认值为 `https://api.openai.com/v1`、`gpt-5.6-luna`、`none`。
- Windows DPAPI API Key 加密。
- 仅由设置页按钮触发的 Outlook 联系人发现和 Contact Index 重建。
- 每个联系人最早 1 封、最近最多 3 封代表邮件取样，以及失败时旧索引保留。
- Outlook 左侧栏中的自建邮件文件夹作为人工归类标签，保存相对路径及其中
  邮件本地解析出的 Contact ID。
- `contact-index.json` 严格校验、稳定联系人 ID、临时文件复验、原子替换和时间戳备份。
- 自然语言转受限 JSON 搜索计划。
- 搜索请求携带已校验 Contact Index，AI 返回联系人 ID 后再在本地精确解析为邮箱。
- 附件名称和扩展名条件解析与安全 AQS 编译。
- 搜索条件分层为硬关键词、概念同义词和软背景提示。
- JSON 校验、字段白名单、值转义和宽松/推荐/精确三档 AQS 编译。
- 默认宽松搜索与无需再次请求 AI 的三档切换按钮。
- Outlook `Explorer.Search` 本地搜索调用。
- 本地语义索引（元数据 + 正文 + 向量 + FTS5），增量同步与内容 Hash。
- 语义 + 关键词混合检索与 RRF 排序。
- 通过 `GetItemFromID` + `InternetMessageId` 校验定位并打开原始邮件。
- 用户级安装和卸载脚本。
- 无第三方 NuGet 依赖的 C# 构建与单元测试。

## 自动验证结果

- C# Release 编译：通过。
- C# 单元测试：40/40 通过。
- Python 单元测试：76/76 通过（含 3 个真实 EmbeddingGemma 2 模型测试）。
- 真实模型验证：在 CPU float32 上加载 768 维文本编码器，
  三个中文→英文检索用例均将正确邮件排在第 1 位。
- C#/Python payload 字段契约测试：通过。
- 发布包：`Outlook-AI-Assistant-v0.5.4.zip` 已生成。

## 尚需用户手动验收

- 在真实 Classic Outlook 中建立 ≥3000 封邮件的索引，确认 UI 保持可操作。
- 在真实邮箱上评估中文查询英文邮件的实际检索质量（Recall@5/Hit@1/MRR）。
- 确认索引中断恢复、增量同步、删除与移动后的定位表现。
- 确认多账户、多 Store、同标题邮件的原邮件定位。
- 确认安装路径包含空格或非 ASCII 字符时仍能运行。

以上测试涉及用户自己的真实邮件或卸载操作，因此未在自动验收中执行。

## 历史原型

根目录的 Office.js 清单、TypeScript 任务窗格和本地 HTTPS 工具保留为历史
参考，不参与桌面版发布。

