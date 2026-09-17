# 维护与发布

## 常见修改位置

### 更新 API 配置说明

修改：

```text
desktop/src/OutlookAiAssistant/Configuration/AppSettings.cs
```

产品使用通用 OpenAI-compatible Base URL、API Key、可编辑 Model 和 Reasoning Effort。
默认值为 `https://api.openai.com/v1`、`gpt-5.6-luna` 和 `none`。更新配置行为时必须同步更新
`SettingsStore.Normalize`、设置页、单元测试、安装文档和隐私披露；本版本不实现模型自动获取。

### 修改摘要格式

修改：

```text
desktop/src/OutlookAiAssistant/Summary/SummaryService.cs
desktop/src/OutlookAiAssistant/Summary/ConversationPromptBuilder.cs
```

必须保留“邮件正文是不可信数据”的提示词边界。不要把读取邮件移到
后台事件或初始化流程。

### 修改 Contact Index

核心位置：

```text
desktop/src/OutlookAiAssistant/EntityIndex/
desktop/src/OutlookAiAssistant/Search/ContactIndexResolver.cs
```

正式文件只能是 `%LOCALAPPDATA%\OutlookAiAssistant\contact-index.json`。
保存时必须先写 `contact-index.new.json`、重新解析和校验，再备份旧文件并原子替换。
联系人 ID 和邮箱由本地代码控制，AI 输出不得覆盖。Outlook 扫描只能由设置页的
“重建联系人索引”按钮触发；不要接入启动、定时器或邮件事件。

Folder Context 来自 Outlook 邮箱树中的自建邮件文件夹。扫描器只保存邮箱根、
相对路径以及其中邮件本地解析出的 Contact ID；不要恢复 Windows 文件夹选择器，
也不要把文件夹名称直接编译为联系人条件。若扩大任何取样数量、字段或网络边界，
必须先更新 `AGENTS.md`、`docs/PRIVACY.md` 和测试清单。

### 扩展搜索字段

需要同时修改：

1. `Models/SearchPlan.cs`
2. `Search/SearchPlannerService.cs` 中的 JSON schema
3. `Search/SearchPlanParser.cs`
4. `Search/ContactIndexResolver.cs`
5. `Search/AqsQueryCompiler.cs`
6. `Search/SearchPlanFormatter.cs`
7. 单元测试

只有 Outlook 官方 AQS 支持且能安全转义的字段才应加入。

搜索文本条件分为：

- `AnchorGroups`：产品名、项目号等很可能逐字出现的硬关键词。
- `ConceptGroups`：业务概念及多语言同义词。
- `HintGroups`：国家、客户类型等未必出现在邮件里的背景提示。

`AqsQueryCompiler.CompileAll` 必须继续生成宽松、推荐和精确三档安全
查询。默认查询以召回率为先；不要把所有模型字段重新拼成一条严格 AND。

### 修改右侧栏

UI 使用代码构建，没有 Designer 文件：

```text
desktop/src/OutlookAiAssistant/UI/AssistantPaneControl.cs
desktop/src/OutlookAiAssistant/UI/SettingsForm.cs
```

保持业务逻辑在服务层，UI 只负责协调。

## 版本升级

同步修改：

1. `Properties/AssemblyInfo.cs`
2. `desktop/build/build.ps1` 中的 ZIP 文件名
3. `desktop/installer/README.txt`
4. 根 `README.md`
5. `docs/INSTALLATION.md`
6. `docs/PROJECT-STATUS.md`

COM CLSID 和 ProgID 不要因普通版本升级而改变，否则旧注册无法直接覆盖。

## 构建

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\desktop\build\build.ps1
```

脚本会：

1. 定位 .NET Framework C# 编译器。
2. 定位 Office、Outlook 和 Extensibility PIA。
3. 编译 AnyCPU COM DLL。
4. 编译并执行单元测试。
5. 创建用户发布包。

## 依赖政策

当前没有 NuGet 依赖。增加第三方依赖前需要确认：

- 是否允许商业/个人分发。
- DLL 是否必须随安装包发布。
- 是否支持 .NET Framework 4.8。
- 是否影响 COM 加载时间。
- 是否扩大网络或隐私边界。

## 调试

本地日志：

```text
%LOCALAPPDATA%\OutlookAiAssistant\logs\addin.log
```

调试流程：

1. 构建 Debug：`build.ps1 -Configuration Debug -SkipPackage`
2. 退出 Outlook。
3. 把 Debug DLL 与安装器放在同一目录后安装。
4. 启动 Outlook 并复现。
5. 先查日志，再使用 Visual Studio Attach to Process（如已安装）。

不要在日志中临时打印邮件正文、API Key、提示词或模型响应。

