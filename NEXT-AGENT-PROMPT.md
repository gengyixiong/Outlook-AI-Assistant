# 接手提示词（复制以下全部内容给下一个 Agent）

---

## 任务背景

你在 `D:\projects_win\Outlook-AI-Assistant` 工作。这是一个 **Classic Outlook 的
C# COM 加载项**，新增了**完全本地运行的语义邮件搜索**功能。

**第一步：务必完整阅读 `HANDOFF.md`**（仓库根目录）。它包含仓库结构、架构、
数据位置、构建命令、当前未完成问题、以及 10 条历史踩坑记录。

## 当前状态

- 版本 **0.7.2**，分支 `agent/semantic-search-v0.5.0`
- 工作区干净，C# 25/25、Python 127/127 测试通过
- 语义搜索功能已能用（用户已确认"可以了"）
- **持久化问题：GPU（CUDA）没有生效，服务仍在用 CPU**

## 首要任务：让用户的 NVIDIA RTX 5060 真正启用 GPU

用户的机器信息（已确认）：
```
NVIDIA GeForce RTX 5060 Laptop GPU, 8151 MiB
驱动 610.88, CUDA UMD Version 13.3     ← 驱动支持 cu130
```

现象：服务运行时报 `no CUDA GPU detected`，venv 里实际是 `torch 2.14.1+cpu`。

**已完成的修复（0.7.2，未真机验证）**：
- 安装脚本的 CUDA 通道从 `cu124` 改为 `cu130`
  （因为 torch 2.14.x 在 Windows 只发布 `cu130` 和 `cpu` 的 wheel，
  `cu124` 的 wheel 不存在 → pip 失败 → 静默回退 CPU）
- `nvidia-smi` 探测除 PATH 外还查 `System32` 与 `NVSMI` 目录
- CUDA 安装失败时打印醒目 WARNING，不再静默降级

**你要做的**：
1. 请用户在他机器上重跑 `setup.cmd`（或 `embedding_service\setup-venv.cmd`）
2. 让用户提供输出，重点确认这几行：
   ```
   Selected build: NVIDIA CUDA (cu130)
   torch: 2.14.1+cu130
   CUDA available: True
   ```
3. 如果 `CUDA available: False`，继续排查：
   - 删除 `%LOCALAPPDATA%\OutlookAiAssistant\embedding_service\.venv` 后重跑
     `setup-venv.cmd`
   - 手工验证 wheel 可装：
     ```
     pip install torch torchvision --index-url https://download.pytorch.org/whl/cu130
     ```
     注意该 wheel 约 **2.5 GB**，网络不稳会超时，建议加
     `--timeout 300 --retries 10`
   - 若确认下载超时是主因，把这两个参数写进 `setup-venv.cmd`

**重要警告**：当前开发机是 **AMD Radeon RX 6800 XT**，**没有 NVIDIA 显卡**。
你无法在本地复现 CUDA 路径，只能靠用户真机反馈。**不要**在没有真机证据的
情况下声称 GPU 已验证通过。

## 工作约束（必须遵守）

- 生产代码保持 **C# 5 兼容**、**.NET Framework 4.8**
- **不加 NuGet 依赖**
- 用户级安装，不需要管理员权限
- 隐私：邮件正文与向量**绝不**上传云端；日志不写正文/提示词/密钥
- 改动版本时同时更新 `VERSION`、`docs/PROJECT-STATUS.md`、发布文档
- **不要推送远程或发布 Release**，除非用户明确授权

## 构建与测试

```powershell
# 编译 + C# 测试 + 打包（会生成 desktop/release/Outlook-AI-Assistant-v<版本>.zip）
powershell -NoProfile -ExecutionPolicy Bypass -File .\desktop\build\build.ps1
```

```powershell
# Python 测试
cd embedding_service
python -m pytest tests/ -q
```

改版本只需改 `VERSION` 文件，`build.ps1` 会自动生成 DLL 版本号和包名。

## 关键文件地图

| 文件 | 作用 |
|---|---|
| `VERSION` | 唯一版本源 |
| `HANDOFF.md` | ★ 交接文档，先读 |
| `desktop/installer/setup.ps1` | 唯一安装入口（版本感知） |
| `desktop/src/OutlookAiAssistant/SemanticSearch/*.cs` | 语义搜索 C# 层 |
| `desktop/src/OutlookAiAssistant/UI/AssistantPaneControl.cs` | 三个页签 UI |
| `embedding_service/app/embedding_model.py` | ★ 模型加载、离线、**GPU 检测** |
| `embedding_service/app/config.py` | 路径、HF_HOME、设备配置 |
| `embedding_service/setup-venv.cmd` | ★★ 选 torch 通道（当前问题所在） |
| `embedding_service/diagnose.cmd` | 一键诊断，报告写桌面 |

## 排障命令（给用户或自己用）

```powershell
# 服务状态（无需令牌）
curl http://127.0.0.1:8765/health

# 启动按钮为什么没反应
Get-Content "$env:LOCALAPPDATA\OutlookAiAssistant\logs\service-start.log"

# venv 里的 torch 情况
& "$env:LOCALAPPDATA\OutlookAiAssistant\embedding_service\.venv\Scripts\python.exe" -c "import torch;print(torch.__version__, torch.cuda.is_available())"

# 一键诊断
& "$env:LOCALAPPDATA\OutlookAiAssistant\embedding_service\diagnose.cmd"
```

## 完成 GPU 后的后续任务

1. 在真实邮箱建立 ≥3000 封索引，**评估检索质量**（Recall@5 / Hit@1 / MRR@10）
   - 目前只有**模拟数据**的评估结果，**绝不能**当作真实质量数据
2. 按用户反馈调优：分块大小、RRF 的 k 值、候选池大小
3. 补齐 `docs/TEST-CHECKLIST.md` 中标注"需手动验收"的项目

## 沟通要求

- 用户是中文使用者，**用中文回复**
- 报告结论时必须**诚实**：区分"已实测验证"与"仅代码正确但未验证"
- 遇到用户报告"还是不行"时，**先拿证据**（日志、版本号、`/health` 输出），
  不要连续盲改代码。必要时让用户跑 `diagnose.cmd` 并把报告发回
- 每次修复都要：加回归测试 → 跑全量测试 → 更新 `PROJECT-STATUS.md` → 提交

---

## 请先这样开始

1. 读 `HANDOFF.md`（必读）
2. 读 `docs/SEMANTIC-SEARCH.md`
3. 跑一次全量构建和测试，确认基线
4. 然后联系用户，索取 GPU 验证所需的信息
