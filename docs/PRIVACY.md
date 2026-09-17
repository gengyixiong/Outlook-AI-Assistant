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

## 联系人索引

触发条件：用户点击设置中的“重建联系人索引”。加载项不会在启动、定时器
或 Outlook 事件中自动调用该功能。

本地扫描：

- 扫描 Classic Outlook 已同步的邮件文件夹，但跳过已删除、垃圾邮件、草稿、
  发件箱和已发送邮件等默认文件夹。
- 首轮只读取发件人显示名、发件人邮箱、接收时间和 Entry ID。
- 按邮箱聚合，每位联系人只加载最早 1 封和最近最多 3 封代表邮件正文。
- Outlook 左侧栏中的自建邮件文件夹会被视为人工归类标签；保存邮箱根名称、
  相对文件夹路径以及该文件夹内邮件本地解析出的 Contact ID。
- 单个异常邮件或不可访问文件夹不会中断其余联系人发现。

发送到当前 AI 服务的数据：

- 本地确认的发件人显示名、邮箱和域名。
- 与该联系人相关的 Outlook 自建归类文件夹路径。
- 上述少量代表邮件的主题、日期及经过引用清理和长度限制的正文片段。

AI 返回的邮箱和 Contact ID 不被信任；正式邮箱来自 Outlook，本地程序生成
稳定 Contact ID。正式索引仅保存必要的联系人 Profile，不保存原始邮件正文。

索引位置：

```text
%LOCALAPPDATA%\OutlookAiAssistant\contact-index.json
```

新索引先写入 `contact-index.new.json` 并重新解析验证。成功后旧索引备份到
`backup/` 再替换；任何失败都保留旧正式索引并清理临时文件。

## Outlook 归类文件夹上下文

这里的 Folder Context 只指 Classic Outlook 左侧邮箱树中由用户创建、可将
邮件拖入的邮件文件夹，不是 Windows 磁盘目录。扫描保存：

- Outlook 邮箱根名称。
- 自建邮件文件夹的相对路径。
- 文件夹内邮件在本地解析出的 Contact ID 列表。

文件夹路径和联系人关联与 Contacts 保存在同一个 `contact-index.json`。
它们是可选且不完整的人工标签，不能直接解析成 `from:`、`to:` 或 `cc:`；
必须先命中 Contacts 中已验证的 Contact ID。

## 搜索

发送到 AI 服务的数据：

- 用户在搜索输入框中输入的自然语言描述。
- 当前本地日期，用于解析“上周”“上个月”等相对时间。
- 已验证的 `contact-index.json`，包括 Contacts 和可选 Folder Context。

不发送：

- 搜索候选邮件的主题、正文、发件人或收件人元数据。
- 候选邮件。
- Outlook 搜索结果。
- 附件内容或完整邮箱。

AI 只生成受限 JSON。Contact ID 必须在本地索引中精确验证后才能解析为
邮箱；不存在的 ID 会被丢弃。Folder Context 不能直接成为联系人条件。
加载项在本机从同一计划编译宽松、推荐和精确三档 AQS；切换档位不会再次
请求 AI。搜索由本机 Outlook Instant Search 执行。

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

- 远程 Base URL 必须使用 HTTPS；localhost 和 127.0.0.1 可使用 HTTP。请求发送到配置的 OpenAI-compatible Chat Completions API。
- 加载项没有遥测、自动更新或其他后台网络请求。
