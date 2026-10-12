# Outlook AI Assistant

> **面向 Classic Outlook for Windows 的个人隐私级 AI 邮件助手**  
> 融合大模型长邮件深度提炼（Executive Brief）与 **100% 本地化神经语义邮件检索（EmbeddingGemma-2 + SQLite 混合检索）**。

[功能亮点](#-核心特色与功能亮点) • [本地语义搜索](#1--完全本地化的语义邮件检索跨语言自然语言搜信) • [隐私安全保证](#-隐私与安全性保障-privacy-invariants) • [环境要求](#-适配环境) • [一键安装与使用](#-如何安装与使用) • [架构设计](#-架构概览) • [许可证](#-许可证)

---

## 🌟 核心特色与功能亮点

### 1. 🔍 完全本地化的语义邮件检索（跨语言自然语言搜信）
告别传统 Outlook 搜信“必须精确命中关键词”的痛点。只需凭记忆用中文随手描述邮件大意，即可跨语言精准定位历史邮件！

- **🧠 SOTA 文本表征模型**：基于 Google 开源的 `google/embeddinggemma-2` 稠密向量嵌入模型（768 维高质量多语言表征），采用 Google 官方非对称检索 Prompt 模板进行优化。
- **🌐 强大的跨语言语义对齐**：用日常中文自然语言即可准确检索出全英文业务往来邮件，即使中文描述与英文邮件正文无任何公共关键词，也能高精度命中召回。
- **⚡ 智能硬件适配（CUDA / CPU 双模式）**：
  - **NVIDIA GPU (CUDA) 加速**：自动检测 NVIDIA 显卡与驱动，使用 CUDA 硬件加速及 bfloat16 精度，实现毫秒级高吞吐向量推理与批量建库。
  - **纯 CPU 稳健回退**：无需独立显卡或处于核显/轻薄办公本环境时，全自动无缝回退为纯 CPU 模式（float32 精度，绝无 float16 下溢退化问题），零门槛普惠运行。
- **🚀 混合检索算法（Hybrid Search + RRF）**：
  - 融合向量稠密相似度（Cosine Similarity）+ SQLite FTS5 全文关键词检索 + 元数据过滤（发件人、日期区间、附件等）+ 互惠排序融合（RRF）。
- **🎯 原生无缝联动**：
  - 搜索结果以结构化卡片直观展示（包含匹配度、标题、发件人、接收时间、文件夹及高亮内容摘要）。
  - 双击结果卡片，直接在 Outlook 原生窗口中打开对应的真实邮件，亦可一键调用 Outlook 原生搜索进行全库回溯。
- **🔕 静默后台运行与系统托盘管理**：
  - 启动服务时不弹出任何黑色命令行窗口（CMD/Console），不占据任务栏主窗口。
  - 自动驻留在 Windows 任务栏右下角系统托盘（System Tray），右键托盘图标可随时查看状态或优雅停止服务。

### 2. ⚡ 长邮件深度提炼（Executive Brief）
自动结合当前邮件及其完整会话（Conversation）上下文，生成结构清晰的要务通报：
- **核心态势**：当前事项处于什么进展；
- **最新变化**：最新一封邮件带来了哪些新信息；
- **行动项（Action Items）**：明确指出需要你回复、决策或推进的事项；
- **关键风险与时间节点**：提取关键日期、金额、约束与风险。
- *注：仅在用户主动点击按钮后单次触发，不监听新邮件，不全盘后台抓取。*

### 3. ⚙️ 平铺式一体化配置面板（In-Place Settings）
- 告别层层弹窗：所有配置项（OpenAI-compatible Base URL、API Key、模型名称、Reasoning Effort 思考程度、语言倾向、截断长度、发件人别名等）直接在 Outlook 侧边栏的“设置”主页平铺展示。
- 提供一键打开模型目录、向量索引目录、日志目录的快捷通道。
- 内置“一键下载模型”，首次下载后永久离线运行。

---

## 🔒 隐私与安全性保障 (Privacy Invariants)

本项目遵循严格的隐私第一原则：
1. **邮件内容绝不上云做向量检索**：
   - 所有的邮件正文切片、嵌入编码（Embedding）、向量计算及 SQLite 检索，**100% 在用户本机本地完成**。
   - 绝不向云端发送任何邮件候选列表、邮件正文、标题或收发人用于重排或搜索。
2. **本地回环隔离**：
   - 本地语义搜索后台服务仅监听回环地址 `127.0.0.1:8765`，局域网与外网均无法访问。
   - 每次安装/初始化生成高强度随机 Token，所有内部通信均经 Token 鉴权。
3. **明示交互触发**：
   - 生成 Executive Brief 或重建联系人画像仅在用户点击对应按钮后按需执行。
4. **安全凭据存储**：
   - 用户配置的大模型 API Key 通过 Windows 原生 **DPAPI**（数据保护 API）安全加密存储，且仅当前 Windows 登录用户可解密。
5. **AQS 白名单编译**：
   - 执行 Outlook 原生搜索时，仅通过经过校验的 AQS 查询编译器执行白名单字段检索，绝不直接将模型原始输出带入系统执行。

---

## 💻 适配环境

- **操作系统**：Windows 10 / Windows 11
- **运行时环境**：.NET Framework 4.8（系统内置，**无需安装庞大的 .NET SDK 或新版运行时**）
- **客户端**：Classic Outlook for Windows 2016 / 2019 / 2021 / Office 365 / 2024 LTSC（已实测验证 64 位经典版 Outlook）
- **邮箱账户**：支持 Exchange Online、Microsoft 365（包含世纪互联运营版）以及常见企业邮箱
- **Python 环境（用于本地语义搜索）**：Python 3.10 – 3.12（可选组件，如仅需摘要功能甚至无需安装 Python）

---

## 📦 如何安装与使用

### 一键安装

1. 从 [Releases 页面](../../releases) 下载最新的发布包 ZIP 文件并**解压至任意本地文件夹**（*切勿在压缩包内直接双击运行*）。
2. 完全退出 Classic Outlook。
3. 双击解压目录中的 **`setup.cmd`**：
   - 自动检测并配置 Python 环境；
   - 自动检测是否有 NVIDIA GPU 驱动，智能安装 CUDA / CPU 版依赖；
   - 自动安装 / 升级 Outlook COM 加载项（仅写入当前用户目录与注册表，**无需管理员权限**）；
   - 一键引导下载本地模型（约 1.5 GB，仅需下载一次）。
4. 重新打开 Outlook，界面右侧即可看到“AI 邮件助手”面板！

### 使用指南

1. **配置大模型（用于 Executive Brief）**：
   - 切换到“设置”页，填写兼容 OpenAI 格式的接口地址（Base URL）、API Key、模型名（如 `gpt-4o`、`claude-3-5-sonnet`、`deepseek-chat` 等）及思考强度（Reasoning Effort）。
   - 点击“保存设置”。
2. **使用本地语义搜索**：
   - 在“语义搜索”页面点击“启动服务”（后台静默启动，右下角任务栏出现托盘图标）；
   - 点击“建立语义索引”（首次建立会扫描历史邮件并存入本地数据库，耗时取决于邮件数量与硬件，可随时中断）；
   - 在搜索框中输入任意自然语言（例如：“*上个月讨论关于接口超时的方案*”），回车或点击搜索；
   - 双击匹配的卡片直接打开原邮件！

---

## 🏗️ 架构概览

```text
Classic Outlook (Win32 / COM)
  │
  ├─ OutlookAiAssistant.dll (C# 5 / .NET 4.8 COM Add-in)
  │    ├─ AssistantPaneControl (右侧侧边栏 UI)
  │    ├─ ExecutiveBriefService (邮件会话分析与云端 LLM 摘要)
  │    ├─ AqsQueryCompiler (白名单安全原生搜索编译)
  │    ├─ SemanticServiceLauncher (后台无窗口静默进程管理)
  │    ├─ SemanticServiceTray (Windows 系统托盘通知与交互)
  │    └─ SemanticSearchClient (本地安全 HTTP 客户端)
  │
  └─ [本地回环 127.0.0.1:8765]
       │
       ▼
     embedding_service (本地离线 Python 服务)
       ├─ FastAPI / Uvicorn (轻量 HTTP 微服务，仅监听 127.0.0.1)
       ├─ Google EmbeddingGemma-2 (768 维稠密向量表征)
       │    ├─ CUDA 加速 (NVIDIA GPU / bfloat16)
       │    └─ 纯 CPU 回退 (float32)
       └─ SQLite 本地数据库
            ├─ 向量存储与余弦距离检索
            └─ FTS5 全文索引与 RRF 混合排重
```

---

## 📄 许可证

本项目采用 [MIT License](LICENSE)。

Copyright (c) 2026 Yixiong Geng
