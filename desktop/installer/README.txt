Outlook AI Assistant 0.4
==========================

系统要求
--------
- Windows 10/11
- Microsoft classic Outlook 2016 或更高版本
- .NET Framework 4.8（Windows 10/11 通常已包含）

安装
----
1. 完全退出 classic Outlook。
2. 双击 Install.cmd。
3. 不需要管理员权限。
4. 重新启动 Outlook。
5. 在右侧栏“设置”页填写 OpenAI-compatible Base URL、API Key、Model 和 Reasoning Effort。
   新安装默认使用 https://api.openai.com/v1、gpt-5.6-luna 和 None；模型可手动输入，Reasoning Effort 可选 None、Medium 或 Max。
6. 如需联系人称呼搜索，可在设置中点击“重建联系人索引”；该操作不会后台运行。

更新
----
完全退出 Outlook 后，直接运行新版本中的 Install.cmd。
已有设置、加密 API Key 和联系人索引会保留。

卸载
----
完全退出 Outlook 后，双击 Uninstall.cmd。
默认保留设置、加密 API Key 和联系人索引，便于以后重新安装。

隐私
----
- 只有点击“生成 Executive Brief”后，当前邮件及受限会话上下文才会发送给所选 API。
- 只有点击“重建联系人索引”后，才会扫描 Outlook；每位候选最多发送最早 1 封、最近 3 封代表邮件的受限正文摘录。
- Outlook 左侧栏中的自建邮件文件夹会作为人工归类标签，并关联其中邮件的联系人。
- 搜索功能只发送用户输入的搜索描述和已校验的 Contact Index，不发送候选邮件、搜索结果或邮件正文。
- AI 返回的联系人 ID 必须在本地索引中精确命中，才能转换为邮箱条件。
- 同一搜索计划可在宽松、推荐、精确三档之间本地切换，无需再次请求 API。
- 邮件搜索、索引和结果显示均由本机 Outlook 完成。
- API Key 通过 Windows DPAPI 为当前 Windows 用户加密。
