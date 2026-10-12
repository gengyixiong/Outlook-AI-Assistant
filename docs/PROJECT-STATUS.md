# 项目状态

更新时间：2026-10-12

## 当前路线

Office.js 个人侧载在当前世纪互联租户中不可用，因此正式路线已经切换为
classic Outlook 用户级 COM 加载项。

## 0.6.6 修复：启动慢/卡住——彻底离线加载

- 定位根因：即使模型已缓存，启动仍会联网并重试，导致界面“正在启动”卡住
  数分钟（日志可见 `WinError 10060` 重试 1/5–5/5）。
- 真正的原因有两层：
  1. `HF_HUB_OFFLINE` 在 `huggingface_hub` **导入时**就被读取并缓存，
     之后再设环境变量无效；我们之前设得太晚。
  2. 即使传了 `local_files_only=True`，`SentenceTransformer` 仍会向 Hub
     查询 prompt/adapter 元数据。
- 修复：在导入任何 Hugging Face 库**之前**，用纯文件系统检查判断模型是否
  已缓存，若已缓存则设置 `HF_HUB_OFFLINE=1` / `TRANSFORMERS_OFFLINE=1`，
  并**直接按本地快照路径加载**，完全不访问网络。
- 兼容旧缓存：同时检查新模型目录与旧的
  `%USERPROFILE%\.cache\huggingface`，因此之前下载的模型无需重新下载。
- 实测：加载从 3–6 分钟降到 **7.5 秒**，网络请求数为 **0**。
- 界面“正在启动”现在显示已等待秒数，超时延长到 180 秒，并在进程退出时
  立即报错。

## 0.6.5 改进：常驻服务、GPU 自动检测、模型目录修复

- 修复模型目录为空：之前虽然配置了 `%LOCALAPPDATA%\OutlookAiAssistant\models`，
  但 `snapshot_download` / `SentenceTransformer` 仍使用 Hugging Face 默认缓存
  （`%USERPROFILE%\.cache\huggingface`），所以配置目录一直是空的。现在设置
  `HF_HOME` 指向该目录，模型真正下载到这里。
- 自动检测计算设备：有可用 NVIDIA CUDA 时用 GPU（bfloat16），否则回退到
  CPU（float32）；`/health` 返回实际设备与选择原因，界面显示。
- `setup-venv.cmd` 现在会检测 NVIDIA GPU：有则装 CUDA 版 PyTorch，
  否则（或安装失败）回退到 CPU 版。
- 服务改为后台常驻、无控制台窗口（用 `pythonw.exe` 启动）。需要排错时
  设置 `OAA_SERVICE_CONSOLE=1` 可恢复可见窗口。
- 语义搜索页新增“重启服务”“停止服务”，可完整控制服务生命周期；
  `/service/shutdown` 端点支持优雅关闭。
- 新增 GPU 检测、设备字段与关闭端点的测试。

## 0.6.4 改进：模型只下载一次，之后完全离线启动

- 澄清：启动时看到的 `Fetching 15 files ... 0.00B` 是**本地缓存校验**，
  不是重新下载；随后的几十秒是从磁盘加载 1.4 GB 权重，不是联网。
- 优化：新增 `EmbeddingModel.is_cached()`，本地已缓存时以
  `local_files_only=True` 加载，**启动不再访问 Hugging Face**，
  断网也能正常启动。
- 实测：在 `HF_HUB_OFFLINE=1` 下启动成功，日志显示 `[offline, from cache]`。
- `/health` 新增 `model_cached` 字段；C# 端解析。
- 新增 2 个测试，并验证移除离线逻辑时测试会失败。

## 0.6.3 改进：卡片双击打开、目录按钮、使用说明

- 结果卡片去除“打开邮件”和“Outlook 搜索”按钮，改为**双击整张卡片**打开
  原邮件；右键菜单仍提供“打开邮件”与“在 Outlook 中搜索此标题”。
- 修复设置页“打开模型/索引/日志目录”无法打开：路径在服务未运行时也会
  回退到确定性默认目录，目录不存在时自动创建后再打开；路径旁会标注
  “（尚未创建）”。
- 设置页新增“使用说明”按钮，弹窗列出完整的部署与每次使用步骤及常见问题。

## 0.6.2 改进：结果卡片自适应面板宽度

- 结果卡片不再使用固定宽度：改为在 100% 宽的表格列中按行 `Dock=Fill`，
  卡片与文字随 Outlook 侧栏宽度一起伸缩，拉宽/收窄都跟随。
- 语义搜索页的状态标签也改为随面板宽度自动换行，不再固定 350px。
- 回归测试 `BuildsVisibleResultCard` 现在会真的改变宿主宽度，断言卡片
  会随之收窄；已验证固定宽度实现在新断言下会失败。

## 0.6.1 修复：语义搜索结果无法显示

- 修复严重 UI 缺陷：搜索结果卡片使用了嵌套 `Dock=Top` 面板，在带
  `AutoScroll` 的 `FlowLayoutPanel` 中会被压缩到约 **2 像素宽**，导致 10 条
  结果全部看不见，也无法点击。
- 重写卡片布局为单列 `AutoSize` 表格面板，固定宽度（330px）并让标签自动
  换行；实测卡片尺寸从 2×36 恢复到 333×152。
- 新增 `BuildsVisibleResultCard` 回归测试，断言卡片宽高可用且按钮可渲染；
  已验证该测试在旧布局下会失败。

## 0.6.0 服务控制集成与移除 AQS 搜索

- 语义搜索页新增“启动服务”按钮，并在顶部常驻显示服务/模型/索引状态。
- 设置页显示模型/索引/日志目录，可一键打开，并显示磁盘占用；路径由
  服务 `/health` 返回（仅路径，无邮件内容）。
- 彻底移除自然语言（AQS）搜索功能与联系人索引功能（代码、UI、文档）。
- 保留 `OutlookSearchService`，仅用于语义结果卡片的“Outlook 搜索”回退。
- 语义检索业务逻辑保持不变。

## 0.5.7 修复：避免在下载目录创建虚拟环境

- 新增保护：如果直接双击解压目录中的 `setup-venv.cmd`（通常位于
  Downloads 或 Temp），脚本会拒绝在下载目录创建 `.venv`，
  而是提示改用 `install-service.cmd`。
- 原因：在下载目录创建的虚拟环境会随文件夹一起被删除，导致语义搜索
  失效。正确入口是 `install-service.cmd`，它把服务安装到固定位置并在
  那里创建环境。
- `setup-venv.cmd --no-pause` 仍可供 `install-service.ps1` 非交互调用。
- 新增 2 个批处理测试覆盖此保护。

## 0.5.6 修复：避免从 ZIP 内直接运行安装脚本

- `Install.cmd` / `Uninstall.cmd` 现在会先检查旁边的 `.ps1` 是否存在。
  如果直接从 ZIP 压缩包内双击，Windows 的临时解压目录可能不完整，
  导致 PowerShell 报 `install.ps1 不存在`。此时会明确提示“请先解压整个
  ZIP 再运行”。
- 修复一个误导性提示：之前 PowerShell 失败后仍显示 `Installation
  completed.`。现在会正确传播退出码，失败时显示 `Installation FAILED`。
- 安装脚本测试扩展到 `desktop/installer/` 目录。

## 0.5.5 修复：拦截过新的 Python

- 修复安装失败：目标电脑上的 Python 3.15 太新，PyTorch/PyYAML 都没有
  对应预编译 wheel，pip 转而从源码编译并报
  `Microsoft Visual C++ 14.0 or greater is required`。
- `setup-venv.cmd` 现在优先选择 3.12 / 3.11 / 3.10，并在检测到 3.13+ 时
  直接拒绝，给出“安装 Python 3.12”的明确指引，不再尝试编译。
- `install-service.ps1` 失败时列出两条最常见原因（Python 过新、
  Microsoft Store 别名）。
- `requirements.txt` 与文档明确标注仅支持 Python 3.10-3.12。

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
- C# 单元测试：24/24 通过。
- Python 单元测试：113/113 通过（含 3 个真实 EmbeddingGemma 2 模型测试）。
- 真实模型验证：在 CPU float32 上加载 768 维文本编码器，
  三个中文→英文检索用例均将正确邮件排在第 1 位。
- C#/Python payload 字段契约测试：通过。
- 发布包：`Outlook-AI-Assistant-v0.6.6.zip` 已生成。

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

