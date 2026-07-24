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
                "  \"text_groups\": [[\"同义词A\", \"同义词B\"]],");
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
                "外层关键词组之间是 AND；每个内层数组中的词是 OR 同义词。");
            prompt.AppendLine(
                "只有用户明确限定主题或正文时才使用 subject_groups 或 body_groups；"
                    + "否则使用 text_groups。");
            prompt.AppendLine(
                "未提到的字符串和数组字段使用空值，未提到的布尔字段使用 null。");
            prompt.AppendLine(
                "不要添加用户没说过的人员、项目、日期或业务条件。");
            return prompt.ToString();
        }
    }
}
