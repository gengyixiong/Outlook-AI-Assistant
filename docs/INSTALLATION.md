# 安装与排障

## 最简单的方式：一个脚本

解压发布包后，双击 **`setup.cmd`** 即可。它会按顺序完成全部安装，并且
**可以重复运行**：

1. 检查 Python 3.10–3.12
2. 安装 / 更新 Outlook 加载项（COM 注册）
3. 安装 / 更新语义搜索服务代码
4. 创建 Python 依赖环境（`venv`）—— 只需一次
5. 下载本地模型（约 1.5 GB）—— 只需一次

它会先对比包内版本与已安装版本：

- **版本相同**：只补齐缺失的组件，其余全部跳过（几秒完成）
- **版本较新**：覆盖代码，但**保留**依赖环境、模型缓存、语义索引、设置和加密 API Key

常用参数（在 cmd 中运行 `setup.cmd -参数名`）：

| 参数 | 作用 |
| --- | --- |
| `-CheckOnly` | 只检查状态，不修改任何文件 |
| `-Force` | 强制重做全部步骤 |
| `-SkipModel` | 跳过模型下载 |
| `-NoPause` | 结束时不等待按键 |

安装完成后即可删除解压出来的文件夹。

### setup.cmd 不做的事

**建立语义索引不在安装脚本里**。它需要读取几千封邮件、耗时较长，且要求
Outlook 正在运行，因此请在插件界面里点击“建立 / 更新语义索引”。

## 系统要求

- Windows 10/11
- classic Outlook 2016 或更高版本
- .NET Framework 4.8
- Outlook 安装中包含 Office Primary Interop Assemblies
- 语义搜索需要 Python 3.10–3.12（可选；不安装则仅使用原有功能）

本项目不支持 new Outlook。安装不需要管理员权限。

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

1. 安装 Python 3.10–3.12（`python.org`），安装时勾选 **Add python.exe to PATH**。
2. 进入发布包中的 `embedding_service` 目录。
3. 双击 `install-service.cmd`，把服务安装到固定位置：

```text
%LOCALAPPDATA%\OutlookAiAssistant\embedding_service
```

   这一步会同时创建虚拟环境并装好依赖（包含 CPU 版 PyTorch）。完成后，**下载解压的发布包文件夹就可以删除了**，卸载、清理 Downloads 都不会影响它。

4. 进入安装目录，双击 `download-model.cmd`，**首次会下载模型（约 1.5 GB，
   需要联网）**：

```text
%LOCALAPPDATA%\OutlookAiAssistant\models
```

   该窗口会显示实时进度（`Fetching N files: xx%` 与下载速度），
   结束时显示 `SUCCESS`，并**不会自动关闭窗口**。
   （如果模型已缓存，这一步几秒就完成。）

5. 双击 `start-service.cmd`（或使用开始菜单中的“Start semantic service”
   快捷方式）启动服务；它监听 `127.0.0.1:8765`。**此窗口需保持打开。**
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

### 目录结构（哪些能删）

| 路径 | 内容 | 能否删除 |
| --- | --- | --- |
| 解压出的发布包文件夹 | 安装脚本与服务源码 | 安装服务后可删 |
| `%LOCALAPPDATA%\OutlookAiAssistant\app` | 加载项 DLL | 通过 `Uninstall.cmd` |
| `%LOCALAPPDATA%\OutlookAiAssistant\embedding_service` | 已安装的服务与 `.venv` | 通过 `uninstall-service.ps1` |
| `%LOCALAPPDATA%\OutlookAiAssistant\semantic-index` | 邮件索引（含正文副本与向量） | 可删（会清空索引） |
| `%LOCALAPPDATA%\OutlookAiAssistant\models` | 模型缓存 | 可删（需重新下载） |
| `%LOCALAPPDATA%\OutlookAiAssistant\logs` | 日志 | 可删 |

卸载语义服务（保留模型与索引）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\uninstall-service.ps1
```

如需连模型、索引一起删除，加上 `-RemoveModel -RemoveIndex`。

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

### 第一步就报 “Microsoft Visual C++ 14.0 or greater is required”

这说明电脑上的 Python **太新**（0.5.4 日志里的 Python 3.15）。PyTorch 目前
**没有为 Python 3.13 及以上提供预编译安装包**，pip 只好从源码编译，
而机器上没有 C++ 编译器，于是失败。

解决：**安装 Python 3.12**（不要装 3.13/3.14/3.15）：

1. 到 https://www.python.org/downloads/ 下载 **Python 3.12**。
2. 安装时勾选 **Add python.exe to PATH**。
3. 关闭并重新打开命令行窗口（让 PATH 生效）。
4. 重新运行 `install-service.cmd`。

`setup-venv.cmd` 会优先选用 3.12 / 3.11 / 3.10，并在检测到不受支持的版本时
直接拒绝并给出提示，而不是尝试编译。

### 第一步就报 “Python was not found / Microsoft Store”

这说明 `python` 指向了 Microsoft Store 的占位程序，而不是真正的安装。
按以下顺序处理：

1. 从 https://www.python.org/downloads/ 安装 Python 3.10–3.12，安装时勾选
   **Add python.exe to PATH**。
2. 如果已经安装但仍报这个错，打开
   **设置 > 应用 > 高级应用设置 > 应用执行别名**，把 `python.exe` 和
   `python3.exe` 两个别名**关掉**。
3. 重新运行 `setup-venv.cmd`。

`setup-venv.cmd` 会优先使用 `py` 启动器，并在 PATH 中自动跳过 Store
占位程序，也会尝试常见的安装目录，因此只要真正装了 Python 就能找到。

### 服务或模型问题

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

