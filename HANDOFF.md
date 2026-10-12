# 交接文档（HANDOFF）

> 面向接手的 Agent / 开发者。读完这一份即可继续工作。
> 最后更新：2026-10-12，对应版本 **0.7.2**，分支 `agent/semantic-search-v0.5.0`

---

## 0. 一句话现状

为 Classic Outlook 的 C# COM 加载项新增了**完全本地运行的语义邮件搜索**
（Python 服务 + EmbeddingGemma 2 + SQLite + 混合检索）。
功能已打通，**但 GPU（CUDA）路径尚未在真机验证通过**，这是当前最优先事项。

---

## 1. 仓库结构（关键部分）

```text
Outlook-AI-Assistant/
├── VERSION                          # ★ 唯一版本源（0.7.2）
├── AGENTS.md                        # 项目约束（隐私/兼容性）
├── README.md                        # 用户向说明
├── docs/
│   ├── ARCHITECTURE.md              # 架构
│   ├── PRIVACY.md                   # 隐私边界
│   ├── INSTALLATION.md              # 安装与排障
│   ├── SEMANTIC-SEARCH.md           # ★ 语义搜索详细设计
│   ├── PROJECT-STATUS.md            # ★ 每个版本的变更记录
│   └── TEST-CHECKLIST.md            # 测试清单
├── desktop/
│   ├── build/build.ps1              # ★ 构建入口（读 VERSION，生成 DLL 版本）
│   ├── installer/
│   │   ├── setup.cmd / setup.ps1    # ★ 唯一安装入口（版本感知）
│   │   ├── Install.cmd / install.ps1        # 仅装加载项（保留，排错用）
│   │   └── Uninstall.cmd / uninstall.ps1
│   ├── src/OutlookAiAssistant/
│   │   ├── AddIn/Connect.cs         # COM 入口，组装所有服务
│   │   ├── SemanticSearch/          # ★ 语义搜索 C# 层
│   │   │   ├── SemanticSearchClient.cs      # loopback HTTP 客户端
│   │   │   ├── SemanticServiceLauncher.cs   # ★ 启停服务（易错点，见 §6）
│   │   │   ├── SemanticIndexService.cs      # 索引编排（COM 线程）
│   │   │   ├── OutlookMailScanner.cs        # GetTable 读元数据
│   │   │   ├── OutlookMailLocator.cs        # 定位并打开原邮件
│   │   │   ├── SemanticSearchFacade.cs      # UI 门面
│   │   │   └── SemanticSearchModels.cs
│   │   ├── UI/AssistantPaneControl.cs       # 三个页签：摘要/语义搜索/设置
│   │   ├── UI/SettingsForm.cs               # 设置 + 索引管理 + 本地路径
│   │   ├── Outlook/OutlookSearchService.cs  # 仅用于结果卡片回退到 Outlook 搜索
│   │   ├── Summary/                         # Executive Brief（未改动）
│   │   └── AI/OpenAiCompatibleClient.cs     # 仅 Executive Brief 使用
│   └── tests/OutlookAiAssistant.UnitTests/Program.cs   # 25 个测试
└── embedding_service/               # ★ Python 本地语义服务
    ├── app/
    │   ├── config.py                # ★ 路径/HF_HOME/设备配置
    │   ├── embedding_model.py       # ★ 模型加载、离线、GPU 检测
    │   ├── api.py                   # FastAPI 端点
    │   ├── database.py              # SQLite + FTS5
    │   ├── search.py                # 语义+关键词+RRF
    │   ├── indexing.py              # 分块+向量化
    │   ├── text_utils.py            # CJK 分词（关键）
    │   ├── body_cleaner.py          # 正文清洗/HTML
    │   └── main.py                  # CLI 入口
    ├── tests/                       # 127 个测试
    ├── setup-venv.cmd               # ★★ 建 venv + 选 torch 通道（当前问题所在）
    ├── download-model.cmd           # 下载模型
    ├── start-service.cmd            # 启动服务（手动排错用）
    ├── install-service.cmd/.ps1     # 装服务到固定位置
    ├── uninstall-service.ps1
    └── diagnose.cmd/.ps1            # ★ 一键诊断（排障首选）
```

---

## 2. 架构与数据流

```text
Outlook UI (AssistantPaneControl)
  └─ 语义搜索页
       ├─ SemanticServiceLauncher ──► pythonw/python 后台进程（无窗口）
       │                                └─ uvicorn on 127.0.0.1:8765
       ├─ SemanticSearchFacade
       │    ├─ SemanticSearchClient ──HTTP(127.0.0.1 + Bearer token)──► 服务
       │    └─ OutlookMailLocator ──► GetItemFromID → MailItem.Display()
       └─ SemanticIndexService
            └─ OutlookMailScanner（仅 UI 线程访问 COM）
                 ├─ MAPIFolder.GetTable() 读元数据
                 └─ MailItem 正文（仅新建/更新时）

Python 服务
  ├─ EmbeddingModel（google/embeddinggemma-2，仅文本编码器，768 维）
  ├─ Database（SQLite: mails / chunks / chunks_fts）
  ├─ HybridSearch（语义 top40 + 关键词 top40 → RRF → top10）
  └─ FastAPI 端点：/health /index/mail /search /stats /model/* /service/shutdown
```

**关键设计约束（勿破坏）**：
- 邮件正文**绝不**离开本机。服务只监听 `127.0.0.1`。
- C# 侧**不**直接读写 SQLite，不引入 NuGet。
- 所有 COM 调用必须在 UI 线程；跨线程用 `Control.Invoke`。
- Python 服务与 Outlook 是**独立进程**，不共享 COM。

---

## 3. 本地数据位置（全部在 `%LOCALAPPDATA%\OutlookAiAssistant\`）

| 路径 | 内容 | 删除影响 |
|---|---|---|
| `app\OutlookAiAssistant.dll` | 加载项 DLL | 用 `Uninstall.cmd` |
| `embedding_service\` | 服务代码 + `.venv` | 用 `uninstall-service.ps1` |
| `embedding_service\.venv\` | Python 依赖（含 torch） | 重装需重下 |
| `models\huggingface\hub\` | 模型缓存 | 需重新下载 |
| `semantic-index\mail.db` | ★ 邮件正文副本+向量 | 索引清空 |
| `logs\` | addin / semantic-service / **service-start**.log | 可删 |
| `semantic-service-token` | 本地 API 令牌 | 自动重建 |
| `.installed-version` | 安装版本标记 | 影响 setup 判断 |
| `settings.json` | 设置 + DPAPI 加密 API Key | 丢失需重配 |

**旧位置**（历史遗留，代码会兼容查找但不写入）：
`%USERPROFILE%\.cache\huggingface\hub\`

---

## 4. 安装与运行（用户视角）

```text
1. 解压发布包（不要在 ZIP 内直接运行）
2. 关掉 Outlook
3. 双击 setup.cmd        ← 唯一入口，可重复运行
4. 打开 Outlook → 语义搜索页 → 启动服务 → 建立语义索引 → 搜索
```

`setup.cmd` 参数：`-CheckOnly`（只检查）/ `-Force` / `-SkipModel` / `-NoPause`

**手工排错入口**（都保留着，很重要）：
```bat
embedding_service\start-service.cmd        :: 前台启动，能看报错
embedding_service\diagnose.cmd             :: 一键诊断，报告写到桌面
embedding_service\setup-venv.cmd           :: 单独建/重建 venv
```

---

## 5. 构建与测试

```powershell
# 全量：编译 + C# 测试 + 打包
powershell -NoProfile -ExecutionPolicy Bypass -File .\desktop\build\build.ps1

# 只编译和测试
.\desktop\build\build.ps1 -SkipPackage
```

```powershell
# Python 测试
cd embedding_service
python -m pytest tests/ -q
```

当前状态：**C# 25/25，Python 127/127 通过**。

版本改动流程（**必须**）：
1. 改 `VERSION` 文件（唯一版本源）
2. `build.ps1` 会自动生成 DLL 版本和包名
3. 同步更新 `docs/PROJECT-STATUS.md`

---

## 6. ★ 当前未完成的问题（最重要）

### GPU（CUDA）未生效 —— 高优先级

**用户环境（已确认）**：
```
NVIDIA GeForce RTX 5060 Laptop GPU
显存 8151 MiB
驱动 610.88
CUDA UMD Version 13.3          ← 驱动完全支持 cu130
```

**现象**：服务运行时报 `no CUDA GPU detected`，实际用的是 CPU。
经查 venv 里是 `torch 2.14.1+cpu`。

**已定位的根因（已修）**：
安装脚本原先选 `cu124` 通道，但 **torch 2.14.x 在 Windows 只发布 `cu130` 和 `cpu`**，
`cu124` 的 wheel 不存在 → pip 失败 → **静默回退 CPU**。
已在 0.7.2 改为 `cu130`。**但尚未在真机验证。**

**★ 下一步必须做的事**：
1. 在用户机器上重跑 `setup.cmd`（或 `setup-venv.cmd`）
2. 确认输出包含：
   ```
   Selected build: NVIDIA CUDA (cu130)
   torch: 2.14.1+cu130
   CUDA available: True
   ```
3. 若 `CUDA available: False`，按下面排查：
   - 删除 `embedding_service\.venv` 后重跑 `setup-venv.cmd`
   - 确认 `pip install torch torchvision --index-url https://download.pytorch.org/whl/cu130` 能成功（约 2.5 GB）
   - 注意：**本仓库开发机是 AMD 显卡，无法复现 NVIDIA 路径**，必须真机验证

**注意**：`cu130` wheel 约 2.5 GB，网络不稳时 pip 会超时。
可考虑在 `setup-venv.cmd` 增加 `--timeout 300 --retries 10`。

### 其他已知限制（非缺陷，属设计范围）

- **未实现**：Reranker、附件内容索引、图片/多模态索引、自动后台同步。
- 本次改动**移除了**：自然语言 AQS 搜索、联系人索引（用户要求）。
- `Summary/`（Executive Brief）链路**未改动**，仍可正常使用。
- 开发机为 AMD 显卡，**NVIDIA/GPU 相关代码从未真机跑过。**

---

## 7. ★ 历史踩坑记录（避免重犯）

这些都是**真实发生过并已修复**的 bug，改动相关代码时务必留意：

| # | 问题 | 根因 | 修复 |
|---|---|---|---|
| 1 | 模型不可用时静默写入无向量数据 | `index_mail` 吞掉异常仍返回成功 | 改为抛 503，不写垃圾数据 |
| 2 | 启动慢 3–6 分钟 | `HF_HUB_OFFLINE` 在 import 时被缓存，设得太晚；且 `local_files_only` 不够 | 导入 HF 前用**文件系统**检查缓存，直接按本地快照路径加载 |
| 3 | 模型目录为空 | 配置了 `model_cache_dir` 但没设 `HF_HOME` | `ensure_directories()` 里设 `HF_HOME` |
| 4 | 搜索结果卡片不可见 | 嵌套 `Dock=Top` 面板在 `AutoScroll+FlowLayoutPanel` 里被压到 2px | 改单列 `AutoSize` 表格，宽度自适应 |
| 5 | **启动服务按钮无效** | `cmd /c "脚本" > 日志` 会把带引号路径当**窗口标题**，什么都不执行 | 改为 `/c call "脚本" > "日志" 2>&1` |
| 6 | 服务报"模型未就绪" | 模型懒加载，`/health` 未触发加载 | `/health` 在已缓存时触发一次加载 |
| 7 | `setup-venv.cmd` 闪退 | 缺 `pause`；且 `if(...)` 块内 `echo` 出现未转义括号会语法错误 | 加 `pause`；括号转义 `^(`；加静态守卫测试 |
| 8 | 在 Downloads 建 venv | 用户双击了 `setup-venv.cmd` | 检测临时目录并引导用 `install-service.cmd` |
| 9 | Python 3.15 装不上 | PyTorch 无 3.13+ 预编译包 | 脚本只接受 3.10–3.12 |
| 10 | **CUDA 未生效** | `cu124` 通道无 torch 2.14 wheel | 改 `cu130`（**待真机验证**） |

---

## 8. 调试工具与常用命令

```powershell
# 查看服务状态（无需令牌）
curl http://127.0.0.1:8765/health

# 服务日志
Get-Content "$env:LOCALAPPDATA\OutlookAiAssistant\logs\semantic-service.log" -Tail 40

# 启动按钮的日志（排查按钮无反应）
Get-Content "$env:LOCALAPPDATA\OutlookAiAssistant\logs\service-start.log"

# 加载项 DLL 实际版本
(Get-Item "$env:LOCALAPPDATA\OutlookAiAssistant\app\OutlookAiAssistant.dll").VersionInfo.FileVersion

# venv 里的 torch 情况
& "$env:LOCALAPPDATA\OutlookAiAssistant\embedding_service\.venv\Scripts\python.exe" -c "import torch;print(torch.__version__, torch.cuda.is_available())"

# 一键诊断（报告写到桌面）
& "$env:LOCALAPPDATA\OutlookAiAssistant\embedding_service\diagnose.cmd"
```

---

## 9. 必须遵守的项目约束（摘自 AGENTS.md）

- 生产代码保持 **C# 5 兼容** 与 **.NET Framework 4.8** 目标。
- **不加 NuGet 依赖**；Office PIA 从 GAC 取。
- 编译用 .NET Framework 的 `csc.exe`，**不需要 .NET SDK**。
- 隐私：邮件正文与向量只留本机；**绝不**上传云端 Embedding；日志不写正文/提示词/密钥。
- 只执行由 `AqsQueryCompiler` 产出的 AQS（注：AQS 功能已移除，此约束仅历史相关）。
- 改动版本时，**同时**更新 `VERSION`、`PROJECT-STATUS.md` 与发布文档。

---

## 10. 建议的后续工作顺序

1. **【最高】真机验证 CUDA**（见 §6），确认 GPU 生效。
2. 在真实邮箱建立 ≥3000 封索引，评估检索质量（Recall@5 / Hit@1 / MRR）。
   目前只有**模拟数据**的评估结果，**不能**当作真实质量。
3. 按用户测试反馈调优：分块大小、RRF 的 k 值、候选池大小。
4. 补 `TEST-CHECKLIST.md` 中标注"需手动验收"的项。
5. 推送分支 / 发布 Release（**需用户明确授权**，目前未推送）。
