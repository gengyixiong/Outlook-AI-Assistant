# Outlook AI Assistant

面向 classic Outlook for Windows 的个人 AI 邮件助手。

当前正式实施路线是本地 COM 加载项，适配本项目确认过的环境：

- classic Outlook 2024 LTSC，64 位
- Build `16.0.17932.20574`
- 世纪互联运营的 Microsoft 365 / Exchange Online
- Windows 10/11 与 .NET Framework 4.8

## 已实现功能

### 手动会话背景与邮件摘要

- 只有点击“生成会话背景与当前邮件摘要”后才读取邮件。
- 支持阅读窗格选中的邮件和单独打开的邮件窗口。
- 使用 Outlook Conversation 关系读取同一会话中已收、已发的历史邮件。
- 先整理会话背景、关键时间线和既有决定，再单独提炼当前邮件。
- 当前邮件中的外部引用历史会保留，独立历史邮件中的重复引用会清理。
- 最多读取 50 封会话邮件；超限时保留最早起因和最近进展。
- 重点提取与当前 Outlook 用户相关的行动项、决定、日期和风险。
- 不监听新邮件，不自动摘要，不批量读取邮箱。

### 自然语言本地搜索

- 只把用户输入的搜索描述发送给配置的 AI 服务。
- AI 返回受限 JSON 搜索计划；不允许返回并直接执行任意 AQS。
- 搜索计划把条件分成硬关键词、概念同义词和软背景提示。
- 默认只使用最独特的硬关键词执行宽松搜索，优先避免漏掉目标邮件。
- 可在右侧栏切换宽松、推荐和精确三档，不会再次请求 AI。
- 加载项在本机验证、转义并编译为 Outlook AQS。
- 最终由 Outlook Instant Search 在本机索引中执行并显示结果。
- 邮件标题、正文、候选邮件和搜索结果均不上传。

### 多 API 提供商

- DeepSeek
- OpenAI
- 豆包 / 火山方舟
- 自定义 OpenAI Chat Completions 兼容接口

API Key 使用 Windows DPAPI 加密，只能由当前 Windows 用户解密。

## 快速开始

构建：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\desktop\build\build.ps1
```

输出：

```text
desktop\release\Outlook-AI-Assistant-v0.2.1.zip
```

安装：

1. 解压发布包。
2. 完全退出 classic Outlook。
3. 双击 `Install.cmd`。
4. 重新启动 Outlook。
5. 在右侧栏“设置”中配置 API。

不需要管理员权限。

## 项目导航

```text
desktop/
  src/OutlookAiAssistant/
    AddIn/          COM 入口、Ribbon 和右侧栏生命周期
    AI/             OpenAI Chat Completions 兼容客户端
    Configuration/  提供商预设、设置与 DPAPI 加密
    Outlook/        当前会话读取和 Outlook 本地搜索
    Search/         搜索计划解析、校验和 AQS 编译
    Summary/        会话去重、字符预算和结构化摘要提示词
    UI/             右侧栏和设置窗口
  tests/            无第三方依赖的单元测试
  build/            可重复构建脚本
  installer/        用户级安装和卸载脚本
```

根目录原有的 TypeScript、`manifest/`、`src/` 和 `server/` 是早期
Office.js 兼容性探针。由于世纪互联环境无法完成个人侧载，它们被保留为
历史参考，不参与当前桌面版构建。

进一步阅读：

- [架构说明](docs/ARCHITECTURE.md)
- [隐私边界](docs/PRIVACY.md)
- [安装与排障](docs/INSTALLATION.md)
- [维护与发布](docs/MAINTENANCE.md)
- [测试清单](docs/TEST-CHECKLIST.md)
- [项目状态](docs/PROJECT-STATUS.md)
