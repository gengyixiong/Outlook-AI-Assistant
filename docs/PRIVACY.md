# 隐私边界

## 摘要

触发条件：用户点击“生成 Executive Brief”。

发送到用户所配置 AI 服务的数据：

- 当前这一封邮件的主题、发件人、收件人、抄送和接收时间。
- Outlook Conversation 找到的同一会话历史邮件的上述元数据。
- 当前邮件原文，以及独立历史邮件去除明显重复引用后的正文。
- 所有正文合计受设置中的“本次发送正文字符上限”约束。
- 当前 Outlook 用户的显示名和邮箱，用于识别“与我相关”的事项。
- 用户在设置中额外填写的本人邮箱和别名/称呼。
- 当前邮件和历史邮件的附件数量，不发送附件文件。

不发送：

- 当前 Outlook 会话之外的其他邮件。
- 邮箱文件夹结构。
- 搜索结果。
- 附件内容。

每次最多读取当前邮件和 49 封独立历史邮件。超过上限时保留最早和最近
的会话区段，并在界面和提示词中标注。Outlook Conversation 不会返回
“已删除邮件”中的项目。加载项不监听 `NewMailEx`、选择变化或收件箱事件。

## 搜索

当前版本只有一种搜索：本机语义搜索（见下一节）。它完全在本机运行，
不向任何云端 AI 发送邮件内容。

自然语言 AQS 搜索和联系人索引已在 0.6.0 中移除，不再有相关数据上传。

## 本地语义搜索（新增）

触发条件：用户点击语义搜索页的“建立 / 更新语义索引”，或点击“语义搜索”。

本功能**完全在本地运行**，不会把任何邮件内容发送到云端：

- 邮件正文、标题、发件人和附件名只用于在**本机**生成向量并写入本地
  SQLite 数据库。
- 向量化模型 `google/embeddinggemma-2` 在本机 Python 进程中运行。
- 语义搜索只把用户输入的自然语言描述在本机编码为向量，并与本地
  索引向量比较。
- 不使用任何云端 Embedding API，不把候选邮件或结果发送给任何 LLM。
- 语义搜索默认不依赖外部 LLM。

本地存储位置：

```text
%LOCALAPPDATA%\OutlookAiAssistant\semantic-index\mail.db   # 元数据、正文、向量
%LOCALAPPDATA%\OutlookAiAssistant\models\                  # 模型缓存
%LOCALAPPDATA%\OutlookAiAssistant\semantic-service-token   # 本地访问令牌
%LOCALAPPDATA%\OutlookAiAssistant\logs\semantic-service.log # 服务日志
```

本地服务：

- 只监听 `127.0.0.1`，不对局域网开放。
- 除健康检查外都需要本地随机 Bearer Token，避免其他本地进程随意读取
  数据库中的邮件内容。
- 服务不可用或模型不可用时给出明确提示，不会让 Outlook 崩溃。
- `/health` 会返回模型、索引、日志目录的**本地文件路径**，供设置页显示和
  一键打开；这些只是路径字符串，不包含任何邮件内容。

数据清理：

- 设置页“删除本地索引”会删除本机保存的邮件正文副本、片段和向量数据。
- Outlook 中的原始邮件不受影响。
- 卸载时默认保留上述数据；使用 `Uninstall.cmd -RemoveSettings` 会一并删除。

正文清洗：

- 去除 Outlook 回复引用、签名和重复引用片段。
- 形似提示词注入的行会被标记为不可信内容，防止邮件正文影响系统行为。
- 邮件正文被视为不可信数据，不能改变系统任务。

## API Key

API Key 使用 Windows DPAPI 的 `CurrentUser` 范围加密，存放于：

```text
%LOCALAPPDATA%\OutlookAiAssistant\settings.json
```

密钥：

- 不以明文写入磁盘。
- 不写入项目或安装目录。
- 不写入日志。
- 不在设置窗口中回显已保存的值。
- 与保存时的 API 请求地址绑定；更改服务地址后必须重新输入，旧密钥不会自动发送到新服务。
- 复制到另一台电脑后不能解密，需要重新填写。

## 日志

本地日志：

```text
%LOCALAPPDATA%\OutlookAiAssistant\logs\addin.log
```

日志只记录生命周期、操作类型和异常类型/消息。代码约束禁止传入：

- 邮件正文或主题。
- 用户搜索描述。
- AI 请求或响应内容。
- API Key。

日志达到 1 MB 后轮换一次。

## 网络

- Base URL 支持 HTTP 和 HTTPS；请求发送到配置的 OpenAI-compatible Chat Completions API。
- 本地语义搜索服务只监听 loopback（`127.0.0.1`），不产生任何外部网络请求。
- 加载项没有遥测、自动更新或其他后台网络请求。
- 模型首次下载需要访问 Hugging Face；下载完成后语义搜索可完全离线运行。
