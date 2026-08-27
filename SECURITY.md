# 安全与隐私约束

本仓库不得提交任何真实邮件内容、API Key、访问令牌、个人设置或运行日志。

## 禁止提交

- `.env`、`settings.json`、`secrets*.json`、`credentials*.json`
- API Key、GitHub Token、私钥或证书私钥
- `%LOCALAPPDATA%\OutlookAiAssistant` 下的设置、联系人索引、索引备份与日志
- `contact-index.json`、`contact-index*.json` 和其中的联系人、邮箱或文件夹上下文
- 邮件导出文件、邮件正文、搜索结果或调试抓包
- 编译产物、安装包和包含本地路径的临时文件

## 提交前检查

1. 查看 `git status --short` 和 `git diff --cached`。
2. 搜索常见密钥前缀、`Bearer` Token 和私钥头。
3. 确认 `.env.example` 只包含空值或明确的占位符。
4. 确认日志、用户配置、联系人索引及其备份仍由 `.gitignore` 排除。

API Key 只能由插件运行时使用 Windows DPAPI 加密后写入当前用户的
`%LOCALAPPDATA%\OutlookAiAssistant\settings.json`，不得复制到仓库。
