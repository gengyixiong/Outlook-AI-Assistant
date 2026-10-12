# 安装与排障

## 系统要求

- Windows 10/11
- classic Outlook 2016 或更高版本
- .NET Framework 4.8
- Outlook 安装中包含 Office Primary Interop Assemblies
- 语义搜索需要 Python 3.10–3.12（可选；不安装则仅使用原有功能）

本项目不支持 new Outlook。

## 安装

1. 解压 `Outlook-AI-Assistant-v0.5.0.zip`。
2. 完全退出 Outlook，在任务管理器中确认没有 `OUTLOOK.EXE`。
3. 双击 `Install.cmd`。
4. 重新启动 classic Outlook。

安装器把 DLL 复制到：

```text
%LOCALAPPDATA%\OutlookAiAssistant\app
```

并只写入当前用户的注册表，因此不需要管理员权限。

## 首次配置

在侧栏“设置”页填写 OpenAI-compatible API 配置：Base URL、API Key、可编辑的 Model 和 Reasoning Effort。

新安装默认使用 `https://api.openai.com/v1`、`gpt-5.6-luna` 和 `none`。
Reasoning Effort 可选 None、Medium 或 Max，对应 API 值 `none`、`medium` 和 `max`。

Base URL 和 Model 可以手动编辑，程序向 `{Base URL}/chat/completions` 发送请求；
如果已填写完整路径则不会重复追加。HTTP 和 HTTPS 地址都可以使用，适合通过 Tailscale
访问内网服务。修改服务地址后必须重新输入 API Key，旧密钥不会发送到新服务。

升级时会迁移旧版 `deepseek` 或 `zhipu` 设置，保留原有 Base URL、Model 和加密 API Key，
并将 Reasoning Effort 设为 `none`；迁移后按普通 OpenAI-compatible 配置使用。

还可以配置当前用户的其他邮箱地址与称呼别名，帮助 Executive Brief
识别与你相关的事项。

联系人索引不会后台构建。只有点击“重建联系人索引”后，系统才扫描 Outlook
联系人候选及左侧栏中的自建邮件文件夹。自建文件夹的相对路径会作为人工归类标签，
并关联其中邮件在本地解析出的 Contact ID。每位候选最早 1 封、最近最多 3 封
代表邮件的受限正文摘录会发送给当前配置的 API 做结构化提取。生成结果保存在：

```text
%LOCALAPPDATA%\OutlookAiAssistant\contact-index.json
```

## 建立语义索引（可选）

语义搜索需要额外的本地 Python 组件。不安装时，原有功能不受影响。

**按顺序执行以下步骤，模型下载步骤在建立索引前是必需的：**

1. 安装 Python 3.10–3.12（`python.org`）。
2. 进入发布包中的 `embedding_service` 目录。
3. 双击 `setup-venv.cmd`，等待依赖安装完成（CPU 版 PyTorch，约数百 MB）。
4. 双击 `download-model.cmd`，**首次会下载模型（约 1.5 GB，需要联网）**
   并缓存到：

```text
%LOCALAPPDATA%\OutlookAiAssistant\models
```

5. 双击 `start-service.cmd` 启动服务；它会监听 `127.0.0.1:8765`。
6. 在 Outlook 侧栏“语义搜索”页点击“建立 / 更新语义索引”。

如果忘记下载模型，索引不会偷偷写入无用数据：界面会明确提示
“本地语义模型尚未就绪”，并提供“下载模型”按钮。也可以直接在界面点击
“下载模型”完成下载（需要联网）。

索引范围和最大邮件数可在“设置 > 语义索引”中调整。模型下载完成后，
语义搜索可完全离线运行。

### 停止索引

建立索引时，“停止索引”按钮可用。点击后会在当前这一封邮件处理完成后
停止；已经完成的邮件会保留，可稍后继续增量同步。索引过程不会阻塞
Outlook UI。

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

## 语义搜索不可用

1. 打开“语义搜索”页，点击“检查本地服务”，确认服务已在运行。
2. 如果提示服务未运行，重新双击 `embedding_service\start-service.cmd`。
3. 查看服务日志：

```text
%LOCALAPPDATA%\OutlookAiAssistant\logs\semantic-service.log
```

4. 如果提示“模型尚未就绪”或“模型不可用”，点击“下载模型”
   （或运行 `download-model.cmd`）完成一次下载。
5. 如果提示“没有找到匹配的邮件”，可能是索引尚未建立，或描述过于宽泛。
6. 服务或模型不可用不会影响 Executive Brief、联系人索引和 Outlook 原生搜索。

## 搜索结果不完整

加载项使用 Outlook Instant Search，因此结果依赖本机索引：

1. 在 Outlook 搜索框中确认普通搜索能工作。
2. 打开 Windows“索引选项”，确认 Microsoft Outlook 已被索引。
3. 等待首次索引完成。
4. 当前 Office UI 为英文，因此编译器使用英文 AQS 字段名。
5. 如果自然语言使用的是联系人称呼而不是邮箱，请在设置中手动重建联系人索引。

## 更新

1. 退出 Outlook。
2. 解压新发布包。
3. 运行新包中的 `Install.cmd`。

用户设置、API Key、联系人索引、语义索引和日志位于安装 DLL 之外，会保留。

## 卸载

退出 Outlook 后运行 `Uninstall.cmd`。

默认保留设置、加密 API Key、联系人索引和语义索引。要同时删除全部本地数据：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\uninstall.ps1 -RemoveSettings
```

该操作会永久删除设置、加密密钥、联系人索引、语义索引（含本地邮件正文副本
与向量）、索引备份和日志。Outlook 中的原始邮件不受影响。

