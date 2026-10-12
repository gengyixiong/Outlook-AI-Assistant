Outlook AI Assistant 0.6.2
==========================

系统要求
--------
- Windows 10/11
- Microsoft classic Outlook 2016 或更高版本
- .NET Framework 4.8（Windows 10/11 通常已包含）
- 语义搜索（可选）：Python 3.10-3.12

安装
----
1. 完全退出 classic Outlook。
2. 双击 Install.cmd。
3. 不需要管理员权限。
4. 重新启动 Outlook。
5. 在右侧栏“设置”页填写 OpenAI-compatible Base URL、API Key、Model 和 Reasoning Effort。
   新安装默认使用 https://api.openai.com/v1、gpt-5.6-luna 和 None；模型可手动输入，Reasoning Effort 可选 None、Medium 或 Max。
6. 如需联系人称呼搜索，可在设置中点击“重建联系人索引”；该操作不会后台运行。

本地语义搜索（可选）
--------------------
重要：必须先安装服务并下载模型，否则无法建立索引。
1. 安装 Python 3.10-3.12（勾选 Add python.exe to PATH）。
2. 进入 embedding_service 目录，运行 install-service.cmd。
   服务会被安装到 %%LOCALAPPDATA%%\OutlookAiAssistant\embedding_service，
   并自动创建虚拟环境。之后下载解压的文件夹可以安全删除。
3. 进入安装目录，运行 download-model.cmd 下载模型（约 1.5 GB，需要联网）。
4. 运行 start-service.cmd（或开始菜单快捷方式）启动本地服务，
   仅监听 127.0.0.1。此窗口需保持打开。
5. 在 Outlook 侧栏“语义搜索”页点击“建立 / 更新语义索引”。
也可直接在界面点击“下载模型”。索引过程中可用“停止索引”中断。
卸载服务（保留模型与索引）：powershell -File uninstall-service.ps1
不安装此组件时，原有功能完全不受影响。

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
- 本地语义搜索完全在本机运行：邮件正文与向量只保存在本地 SQLite，不上传云端；本地服务仅监听 127.0.0.1 并使用随机 Token 鉴权。
- 邮件搜索、索引和结果显示均由本机 Outlook 完成。
- API Key 通过 Windows DPAPI 为当前 Windows 用户加密。
