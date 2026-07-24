# 安装与排障

## 系统要求

- Windows 10/11
- classic Outlook 2016 或更高版本
- .NET Framework 4.8
- Outlook 安装中包含 Office Primary Interop Assemblies

本项目不支持 new Outlook。

## 安装

1. 解压 `Outlook-AI-Assistant-v0.1.0.zip`。
2. 完全退出 Outlook，在任务管理器中确认没有 `OUTLOOK.EXE`。
3. 双击 `Install.cmd`。
4. 重新启动 classic Outlook。

安装器把 DLL 复制到：

```text
%LOCALAPPDATA%\OutlookAiAssistant\app
```

并只写入当前用户的注册表，因此不需要管理员权限。

## 首次配置

在侧栏“设置”页选择：

- DeepSeek：默认 `https://api.deepseek.com` 和
  `deepseek-v4-flash`。
- OpenAI：默认 `https://api.openai.com/v1` 和
  `gpt-5.6-sol`。
- 豆包：默认方舟地址；模型字段填写控制台提供的模型或推理接入点 ID。
- 自定义：填写兼容服务的根地址和模型名称。

地址既可以填写根地址，也可以直接填写完整
`/chat/completions` 地址。

## 看不到侧栏

依次检查：

1. Outlook 的 `File > Options > Add-ins`。
2. 页面底部选择 `COM Add-ins`，点击 `Go`。
3. 确认 `Outlook AI Assistant` 已勾选。
4. 在 Outlook Home 功能区查找“AI 邮件助手 > 显示助手”。
5. 查看本地日志：

```text
%LOCALAPPDATA%\OutlookAiAssistant\logs\addin.log
```

如果 Outlook 把加载项标记为慢速或禁用，请在
`File > Manage COM Add-ins` 中重新启用。不要直接修改系统策略。

## 搜索结果不完整

加载项使用 Outlook Instant Search，因此结果依赖本机索引：

1. 在 Outlook 搜索框中确认普通搜索能工作。
2. 打开 Windows“索引选项”，确认 Microsoft Outlook 已被索引。
3. 等待首次索引完成。
4. 当前 Office UI 为英文，因此编译器使用英文 AQS 字段名。

## 更新

1. 退出 Outlook。
2. 解压新发布包。
3. 运行新包中的 `Install.cmd`。

用户设置、API Key 和日志位于安装 DLL 之外，会保留。

## 卸载

退出 Outlook 后运行 `Uninstall.cmd`。

默认保留设置和加密 API Key。要同时删除全部本地数据：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\uninstall.ps1 -RemoveSettings
```

该操作会永久删除设置、加密密钥和日志。

