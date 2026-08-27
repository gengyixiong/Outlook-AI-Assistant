# 项目状态

更新时间：2026-08-27

## 当前路线

Office.js 个人侧载在当前世纪互联租户中不可用，因此正式路线已经切换为
classic Outlook 用户级 COM 加载项。

## 已完成

- 可维护的 C# 分层项目结构。
- Outlook COM 入口和自定义右侧任务窗格。
- Outlook 功能区“显示助手”开关。
- 按钮触发的当前邮件与同一 Outlook 会话历史读取。
- 会话邮件跨文件夹读取、时间排序、EntryID 去重和 50 封上限取样。
- Executive Brief 固定四段式输出，并结合 Outlook 当前身份、用户邮箱和称呼别名识别相关事项。
- 历史邮件重复引用清理与总字符预算。
- 仅开放 DeepSeek Flash 与 GLM-5.3 Flash，API 地址和模型固定不可编辑。
- Windows DPAPI API Key 加密。
- 仅由设置页按钮触发的 Outlook 联系人发现和 Contact Index 重建。
- 每个联系人最早 1 封、最近最多 3 封代表邮件取样，以及失败时旧索引保留。
- Outlook 左侧栏中的自建邮件文件夹作为人工归类标签，保存相对路径及其中
  邮件本地解析出的 Contact ID。
- `contact-index.json` 严格校验、稳定联系人 ID、临时文件复验、原子替换和时间戳备份。
- 自然语言转受限 JSON 搜索计划。
- 搜索请求携带已校验 Contact Index，AI 返回联系人 ID 后再在本地精确解析为邮箱。
- 附件名称和扩展名条件解析与安全 AQS 编译。
- 搜索条件分层为硬关键词、概念同义词和软背景提示。
- JSON 校验、字段白名单、值转义和宽松/推荐/精确三档 AQS 编译。
- 默认宽松搜索与无需再次请求 AI 的三档切换按钮。
- Outlook `Explorer.Search` 本地搜索调用。
- 用户级安装和卸载脚本。
- 无第三方依赖的构建与单元测试。

## 自动验证结果

- C# Release 编译：通过。
- COM 元数据 / RegAsm 解析：通过。
- 单元测试：29/29 通过。
- 发布包：`Outlook-AI-Assistant-v0.3.0.zip` 已生成。

## 真实 Outlook 验收结果

- 用户级安装与 COM 注册：通过。
- classic Outlook 首次加载：通过。
- 右侧任务窗格停靠：通过。
- Home 功能区“显示助手”按钮：通过。
- 侧栏显示 / 隐藏切换：通过。
- 旧版 API 设置窗口首次打开与提供商列表初始化：通过。
- 旧版摘要按钮和隐私说明加载：通过。
- 未选择邮件时保持“尚未读取邮件”且不自动发送请求：通过。

## 0.3.0 尚需真实 Outlook 与用户 API Key 验收

- 分别使用 DeepSeek Flash、GLM-5.3 Flash 测试 API Key 验证真实请求。
- 验证 Executive Brief 四段结构、身份别名和“不需处理”的兜底文案。
- 手动重建 Contact Index，检查候选取样、进度、备份和单联系人失败回退。
- 验证 Outlook 自建邮件文件夹的相对路径和 Contact ID 关联。
- 使用 Ma、Jason、Richard 等称呼验证 Contact ID 到邮箱的本地解析。
- 使用当前 Outlook 索引验证三档 AQS、日期、附件名称、扩展名和中英文关键词搜索。
- 确认搜索请求期间不读取候选 MailItem，也不发送搜索结果或邮件正文。
- 确认关闭 Outlook 后更新与卸载流程。

以上测试涉及用户自己的 API Key、真实邮件或卸载操作，因此未在自动验收中执行。

## 历史原型

根目录的 Office.js 清单、TypeScript 任务窗格和本地 HTTPS 工具保留为历史
参考，不参与桌面版发布。
