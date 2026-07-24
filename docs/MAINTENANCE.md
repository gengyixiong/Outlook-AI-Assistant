# 维护与发布

## 常见修改位置

### 添加或更新 API 提供商

修改：

```text
desktop/src/OutlookAiAssistant/Configuration/AiProviderPreset.cs
```

如果仍兼容 `POST /chat/completions`，通常不需要修改传输层。模型名称会
变化，发布前应从提供商官方文档核对。

### 修改摘要格式

修改：

```text
desktop/src/OutlookAiAssistant/Summary/SummaryService.cs
```

必须保留“邮件正文是不可信数据”的提示词边界。不要把读取邮件移到
后台事件或初始化流程。

### 扩展搜索字段

需要同时修改：

1. `Models/SearchPlan.cs`
2. `Search/SearchPlannerService.cs` 中的 JSON schema
3. `Search/SearchPlanParser.cs`
4. `Search/AqsQueryCompiler.cs`
5. `Search/SearchPlanFormatter.cs`
6. 单元测试

只有 Outlook 官方 AQS 支持且能安全转义的字段才应加入。

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
5. `docs/PROJECT-STATUS.md`

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
- 是否支持 .NET Framework 4.x。
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

