using System.Collections.Generic;
using System.Text;
using OutlookAiAssistant.Models;

namespace OutlookAiAssistant.Search
{
    public sealed class SearchPlanFormatter
    {
        public string Format(SearchPlan plan, string aqs)
        {
            StringBuilder text = new StringBuilder();
            AppendList(text, "发件人", plan.From);
            AppendList(text, "收件人", plan.To);
            AppendList(text, "抄送", plan.Cc);
            if (plan.ToMe)
            {
                text.AppendLine("收件条件：包含我");
            }

            AppendGroups(text, "任意位置关键词", plan.TextGroups);
            AppendGroups(text, "主题关键词", plan.SubjectGroups);
            AppendGroups(text, "正文关键词", plan.BodyGroups);
            if (!string.IsNullOrWhiteSpace(plan.ReceivedFrom))
            {
                text.AppendLine("开始日期：" + plan.ReceivedFrom);
            }

            if (!string.IsNullOrWhiteSpace(plan.ReceivedThrough))
            {
                text.AppendLine("结束日期：" + plan.ReceivedThrough);
            }

            if (plan.HasAttachments.HasValue)
            {
                text.AppendLine(
                    "附件：" + (plan.HasAttachments.Value ? "有附件" : "无附件"));
            }

            if (plan.IsUnread.HasValue)
            {
                text.AppendLine(
                    "阅读状态：" + (plan.IsUnread.Value ? "未读" : "已读"));
            }

            text.AppendLine(
                "范围："
                    + (plan.Scope == "current_folder" ? "当前文件夹" : "所有文件夹"));
            text.AppendLine();
            text.AppendLine("本地 Outlook AQS：");
            text.AppendLine(aqs);
            return text.ToString().Trim();
        }

        private static void AppendList(
            StringBuilder text,
            string label,
            IList<string> values)
        {
            if (values != null && values.Count > 0)
            {
                text.AppendLine(label + "：" + string.Join(" / ", values));
            }
        }

        private static void AppendGroups(
            StringBuilder text,
            string label,
            IList<List<string>> groups)
        {
            if (groups == null)
            {
                return;
            }

            for (int index = 0; index < groups.Count; index++)
            {
                text.AppendLine(
                    label + " " + (index + 1) + "："
                        + string.Join(" 或 ", groups[index]));
            }
        }
    }
}

