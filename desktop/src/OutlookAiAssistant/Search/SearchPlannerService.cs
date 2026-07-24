using System;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Models;

namespace OutlookAiAssistant.Search
{
    /// <summary>
    /// Sends only the user's natural-language search description to the AI.
    /// It never receives or transmits any Outlook message or search result.
    /// </summary>
    public sealed class SearchPlannerService
    {
        private readonly OpenAiCompatibleClient _client;
        private readonly SettingsStore _settingsStore;
        private readonly SearchPlanParser _parser;

        public SearchPlannerService(
            OpenAiCompatibleClient client,
            SettingsStore settingsStore,
            SearchPlanParser parser)
        {
            _client = client;
            _settingsStore = settingsStore;
            _parser = parser;
        }

        public async Task<SearchPlan> CreatePlanAsync(
            string naturalLanguageQuery,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(naturalLanguageQuery))
            {
                throw new InvalidOperationException("请先输入要查找的邮件描述。");
            }

            string apiKey = _settingsStore.ReadApiKey(settings);
            string response = await _client.CompleteAsync(
                settings,
                apiKey,
                BuildSystemPrompt(DateTime.Today),
                naturalLanguageQuery.Trim(),
                true,
                cancellationToken);
            return _parser.Parse(response);
        }

        private static string BuildSystemPrompt(DateTime today)
        {
            StringBuilder prompt = new StringBuilder();
            prompt.AppendLine(
                "你把用户的自然语言描述转换成 Outlook 本地搜索计划。");
            prompt.AppendLine(
                "你看不到邮箱，也绝不能要求邮件正文、候选邮件或搜索结果。");
            prompt.AppendLine(
                "目标是优先召回目标邮件：宁可多返回一些结果，也不要因为条件过严而漏掉。");
            prompt.AppendLine("今天是 " + today.ToString("yyyy-MM-dd") + "。");
            prompt.AppendLine(
                "把“上周、上个月、今年”等相对日期换成含首尾日期的绝对日期。");
            prompt.AppendLine("只返回一个 JSON 对象，不要 Markdown，不要解释。");
            prompt.AppendLine("严格使用以下字段：");
            prompt.AppendLine("{");
            prompt.AppendLine("  \"from\": [\"姓名或邮箱\"],");
            prompt.AppendLine("  \"to\": [\"姓名或邮箱\"],");
            prompt.AppendLine("  \"cc\": [\"姓名或邮箱\"],");
            prompt.AppendLine("  \"to_me\": false,");
            prompt.AppendLine(
                "  \"anchor_groups\": [[\"独特硬关键词\", \"同一词的其他拼写\"]],");
            prompt.AppendLine(
                "  \"concept_groups\": [[\"概念同义词A\", \"概念同义词B\"]],");
            prompt.AppendLine(
                "  \"hint_groups\": [[\"背景线索A\", \"背景线索B\"]],");
            prompt.AppendLine(
                "  \"subject_groups\": [[\"同义词A\", \"同义词B\"]],");
            prompt.AppendLine(
                "  \"body_groups\": [[\"同义词A\", \"同义词B\"]],");
            prompt.AppendLine("  \"received_from\": \"YYYY-MM-DD\",");
            prompt.AppendLine("  \"received_through\": \"YYYY-MM-DD\",");
            prompt.AppendLine("  \"has_attachments\": null,");
            prompt.AppendLine("  \"is_unread\": null,");
            prompt.AppendLine(
                "  \"scope\": \"all_folders\" 或 \"current_folder\"");
            prompt.AppendLine("}");
            prompt.AppendLine(
                "每个内层数组是同一个词或概念的 OR 别名；不同外层组才是 AND。");
            prompt.AppendLine(
                "anchor_groups 只放很可能逐字出现在邮件里的稀有标识，"
                    + "例如产品名、项目名、订单号、明确公司名或邮箱。");
            prompt.AppendLine(
                "concept_groups 放业务概念及中英或其他相关语言同义词，"
                    + "例如“大赛/competition/contest/competição/concurso”。");
            prompt.AppendLine(
                "hint_groups 放可能只是用户背景知识、未必逐字出现在邮件中的线索，"
                    + "例如国家、地区、客户类型、合作关系。");
            prompt.AppendLine(
                "例如“巴西客户发来的 GstarBIM 大赛邮件”：GstarBIM 是 anchor，"
                    + "大赛及其多语言同义词是 concept，巴西/Brazil/Brasil 是 hint；"
                    + "绝不能把“巴西客户”写入 from。");
            prompt.AppendLine(
                "from、to、cc 只放真实姓名、公司名、域名或邮箱地址；"
                    + "不要放“某国客户、合作伙伴、主办方、以前的客户”等关系描述。");
            prompt.AppendLine(
                "用户在自己的邮箱里泛称“发给我的”时不要设置 to_me；"
                    + "只有明确要求收件人栏为本人时才设置。");
            prompt.AppendLine(
                "只有用户明确限定主题或正文时才使用 subject_groups 或 body_groups；"
                    + "否则使用 anchor_groups、concept_groups 或 hint_groups。");
            prompt.AppendLine(
                "最多输出 2 个 anchor 组、2 个 concept 组、2 个 hint 组；"
                    + "避免把一句描述拆成许多必须同时命中的条件。");
            prompt.AppendLine(
                "如果有多个 anchor 组，把最独特、最可能实际出现的那个排在第一位。");
            prompt.AppendLine(
                "未提到的字符串和数组字段使用空值，未提到的布尔字段使用 null。");
            prompt.AppendLine(
                "不要添加用户没说过的人员、项目、日期或业务条件。");
            return prompt.ToString();
        }
    }
}
