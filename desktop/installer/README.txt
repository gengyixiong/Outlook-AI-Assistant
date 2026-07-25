Outlook AI Assistant 0.2.1
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
5. 在右侧栏“设置”页填写 API 地址、模型和 API Key。

更新
----
完全退出 Outlook 后，直接运行新版本中的 Install.cmd。
已有设置和加密 API Key 会保留。

卸载
----
完全退出 Outlook 后，双击 Uninstall.cmd。
默认保留设置和加密 API Key，便于以后重新安装。

隐私
----
- 只有点击“生成当前邮件摘要”后，当前邮件才会发送给配置的 API。
- 搜索功能只发送用户输入的搜索描述，并默认执行高召回的宽松查询。
- 同一搜索计划可在宽松、推荐、精确三档之间本地切换，无需再次请求 API。
- 邮件搜索、索引和结果显示均由本机 Outlook 完成。
- API Key 通过 Windows DPAPI 为当前 Windows 用户加密。
